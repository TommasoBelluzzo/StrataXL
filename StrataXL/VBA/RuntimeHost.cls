VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RuntimeHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'====================================='
' Copyright (C) 2019 Tommaso Belluzzo '
'          Part of StrataXL           '
'====================================='

' SETTINGS

Option Explicit

' IMPORTS

Private Declare PtrSafe Function AddDllDirectory Lib "Kernel32.dll" _
( _
    ByVal newDirectory As LongPtr _
) As LongPtr

Private Declare PtrSafe Function GetShortPathName Lib "Kernel32.dll" Alias "GetShortPathNameW" _
( _
    ByVal lpszLongPath As LongPtr, _
    ByVal lpszShortPath As LongPtr, _
    ByVal cchBuffer As Long _
) As Long

Private Declare PtrSafe Function SetDefaultDllDirectories Lib "Kernel32.dll" _
( _
    ByVal directoryFlags As Long _
) As Long

Private Declare PtrSafe Function CLRCreateInstance Lib "MSCoreE.dll" _
( _
    ByRef rclsid As Any, _
    ByRef riid As Any, _
    ByRef ppInterface As IUnknown _
) As Long

Private Declare PtrSafe Function VariantCopy Lib "OleAut32.dll" _
( _
    ByRef pvargDest As Variant, _
    ByRef pvargSrc As Variant _
) As Long

#If Win64 Then
    Private Declare PtrSafe Function LoadRuntime Lib "RuntimeLoader64.dll" _
    ( _
        ByRef pCorRuntimeHost As IUnknown _
    ) As Long
#Else
    Private Declare PtrSafe Function LoadRuntime Lib "RuntimeLoader32.dll" _
    ( _
        ByRef pCorRuntimeHost As IUnknown _
    ) As Long
#End If

' MEMBERS

Dim m_Initialized As Boolean
Dim m_TypesCache As Collection
Dim m_ManagedHelper As Object
Dim m_LibrariesPath As String
Dim m_LogFile As String
Dim m_Suffix As String
Dim m_AssemblyCore As mscorlib.Assembly
Dim m_AssemblyStrata As mscorlib.Assembly
Dim m_AssemblySystem As mscorlib.Assembly
Dim m_Domain As mscorlib.AppDomain
Dim m_Host As mscoree.CorRuntimeHost

' CONSTRUCTOR

Private Sub Class_Initialize()

    m_Initialized = False
    Set m_TypesCache = New Collection
    m_LibrariesPath = ThisWorkbook.Path & "\Libraries"
    m_LogFile = ThisWorkbook.Path & "\StrataXL.log"
    m_Suffix = Format(Now, "yyyymmdd_hhnnss")

End Sub

' DESTRUCTOR

Private Sub Class_Terminate()

    Call PerformCleanup

End Sub

' PROPERTY
' Gets the initialization status of the instance.

Property Get Initialized() As Variant

    Initialized = m_Initialized

End Property

' METHOD
' Creates an instance of a .NET class capable of handling delegates and method calls.

Private Function CreateManagedHelper(ByVal domain As mscorlib.AppDomain) As Object

    Dim code As String: code = ""

    code = code + _
        "using System;" + vbCrLf + _
        "using System.IO;" + vbCrLf + _
        "using System.Reflection;" + vbCrLf + _
        "using System.Runtime.InteropServices;" + vbCrLf + _
        "using System.Text;" + vbCrLf + _
        "" + vbCrLf

    Dim dw As String: dw = "DelegateWrapper_" + m_Suffix

    code = code + _
        "public sealed class " + dw + vbCrLf + _
        "{" + vbCrLf

    code = code + _
        "    private Object m_Instance;" + vbCrLf + _
        "    private String m_MethodName;" + vbCrLf + _
        "    private Type m_InstanceType;" + vbCrLf + _
        "" + vbCrLf

    code = code + _
        "    private " + dw + "(Object instance, Type instanceType, String methodName)" + vbCrLf + _
        "    {" + vbCrLf + _
        "        m_Instance = instance;" + vbCrLf + _
        "        m_MethodName = methodName;" + vbCrLf + _
        "        m_InstanceType = instanceType;" + vbCrLf + _
        "    }" + vbCrLf + _
        "" + vbCrLf
        
    code = code + _
        "    private Object Method(Object[] methodParameters)" + vbCrLf + _
        "    {" + vbCrLf + _
        "        return m_InstanceType.InvokeMember(m_MethodName, BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, m_Instance, methodParameters);" + vbCrLf + _
        "    }" + vbCrLf _

    Dim params As String: params = ""
    Dim vars As String: vars = ""

    Dim i As Integer

    For i = 0 To 20
    
        Dim iString As String
        iString = CStr(i)
    
        If (i = 1) Then
            params = "Object p1"
            vars = "p1 "
        ElseIf (i >= 2) Then
            params = params + ", Object p" + iString
            vars = Left(vars, Len(vars) - 1) + ", p" + iString + " "
        End If
        
        code = code + "" + vbCrLf + _
            "    private Object Method" + iString + "(" + params + ") { return Method(new Object[] { " + vars + "}); }" + vbCrLf + _
            "    private void VirtualMethod" + iString + "(" + params + ") { Method(new Object[] { " + vars + "}); }" + vbCrLf

    Next i
    
    code = code + vbCrLf + _
        "    public static Delegate CreateDelegate(Type delegateType, Object instance, String methodName)" + vbCrLf + _
        "    {" + vbCrLf + _
        "        if (!delegateType.IsSubclassOf(typeof(Delegate)))" + vbCrLf + _
        "            throw new ArgumentException(""Invalid delegate type specified."");" + vbCrLf + _
        "" + vbCrLf + _
        "        var miInvoke = delegateType.GetMethod(""Invoke"");" + vbCrLf + _
        "" + vbCrLf + _
        "        if (miInvoke == null)" + vbCrLf + _
        "            throw new ArgumentException(""Invalid delegate type specified."");" + vbCrLf + _
        "" + vbCrLf + _
        "        var parametersLength = miInvoke.GetParameters().Length;" + vbCrLf + _
        "" + vbCrLf + _
        "        if (parametersLength > 20)" + vbCrLf + _
        "            throw new NotSupportedException(""The number of parameters exceeds the maximum threshold of 20."");" + vbCrLf + _
        "" + vbCrLf + _
        "        var wrapper = new " + dw + "(instance, instance.GetType(), methodName);" + vbCrLf + _
        "        var miWrapper = typeof(" + dw + ").GetMethod(((miInvoke.ReturnType == typeof(void)) ? ""VirtualMethod"" : ""Method"") + parametersLength.ToString(), BindingFlags.Instance | BindingFlags.NonPublic);" + vbCrLf + _
        "" + vbCrLf + _
        "        return Delegate.CreateDelegate(delegateType, wrapper, miWrapper);" + vbCrLf + _
        "    }" + vbCrLf

    code = code + "}" + vbCrLf + _
        "" + vbCrLf

    Dim mh As String: mh = "ManagedHelper_" + m_Suffix

    code = code + _
        "[ComVisible(true)]" + vbCrLf + _
        "public sealed class " + mh + vbCrLf + _
        "{" + vbCrLf
        
    code = code + _
        "    private static void LogException(Exception e)" + vbCrLf + _
        "    {" + vbCrLf + _
        "        try" + vbCrLf + _
        "        {" + vbCrLf + _
        "            StringBuilder builder = new StringBuilder();" + vbCrLf + _
        "            builder.AppendLine(String.Concat(""########## EXCEPTION THROWN AT "", DateTime.Now.ToString(""dd/MM/yyyy HH:mm""), "" ##########"", Environment.NewLine));" + vbCrLf + _
        "            builder.AppendLine(String.Concat(""[TYPE] "", e.GetType().ToString(), Environment.NewLine, ""[MESSAGE] "", e.Message, Environment.NewLine, ""[STACK TRACE]"", Environment.NewLine, e.StackTrace, Environment.NewLine));" + vbCrLf + _
        "" + vbCrLf + _
        "            while ((e = e.InnerException) != null)" + vbCrLf + _
        "                builder.AppendLine(String.Concat(""[TYPE] "", e.GetType().ToString(), "" (Inner)"", Environment.NewLine, ""[MESSAGE] "", e.Message, Environment.NewLine, ""[STACK TRACE]"", Environment.NewLine, e.StackTrace, Environment.NewLine));" + vbCrLf + _
        "" + vbCrLf + _
        "            builder.AppendLine();" + vbCrLf + _
        "            builder.AppendLine();" + vbCrLf + _
        "" + vbCrLf + _
        "            File.AppendAllText(@""" + m_LogFile + """, builder.ToString());" + vbCrLf + _
        "        }" + vbCrLf + _
        "        catch { }" + vbCrLf + _
        "    }" + vbCrLf

    code = code + _
        "    public Delegate CreateDelegate(Type delegateType, Object instance, String methodName)" + vbCrLf + _
        "    {" + vbCrLf + _
        "        try" + vbCrLf + _
        "        {" + vbCrLf + _
        "            return " + dw + ".CreateDelegate(delegateType, instance, methodName);" + vbCrLf + _
        "        }" + vbCrLf + _
        "        catch (Exception e)" + vbCrLf + _
        "        {" + vbCrLf + _
        "            LogException(e);" + vbCrLf + _
        "            throw;" + vbCrLf + _
        "        }" + vbCrLf + _
        "    }" + vbCrLf + _
        "" + vbCrLf
        
    code = code + _
        "    public Delegate CreateDelegateWithFunction(Type delegateType, Object functionPointer)" + vbCrLf + _
        "    {" + vbCrLf + _
        "        try" + vbCrLf + _
        "        {" + vbCrLf + _
        "            IntPtr pointer;" + vbCrLf + _
        "" + vbCrLf + _
        "            if (functionPointer is Int32)" + vbCrLf + _
        "                pointer = new IntPtr((Int32)functionPointer);" + vbCrLf + _
        "            else if (functionPointer is Int64)" + vbCrLf + _
        "                pointer = new IntPtr((Int64)functionPointer);" + vbCrLf + _
        "            else" + vbCrLf + _
        "                pointer = (IntPtr)functionPointer;" + vbCrLf + _
        "" + vbCrLf + _
        "            return Marshal.GetDelegateForFunctionPointer(pointer, delegateType);" + vbCrLf + _
        "        }" + vbCrLf + _
        "        catch (Exception e)" + vbCrLf + _
        "        {" + vbCrLf + _
        "            LogException(e);" + vbCrLf + _
        "            throw;" + vbCrLf + _
        "        }" + vbCrLf + _
        "    }" + vbCrLf + _
        "" + vbCrLf

    code = code + _
        "    public Object Invoke(Object instance, MethodBase method, Object[] methodParameters, Object[] isPointer)" + vbCrLf + _
        "    {" + vbCrLf + _
        "        Object[] finalParameters = new Object[methodParameters.Length];" + vbCrLf + _
        "" + vbCrLf + _
        "        for (Int32 i = 0; i < finalParameters.Length; ++i)" + vbCrLf + _
        "        {" + vbCrLf + _
        "            if ((isPointer[i] is Boolean) && (Boolean)isPointer[i])" + vbCrLf + _
        "            {" + vbCrLf + _
        "                Object parameter = methodParameters[i];" + vbCrLf + _
        "" + vbCrLf + _
        "                if (parameter is Int32)" + vbCrLf + _
        "                    finalParameters[i] = new IntPtr((Int32)parameter);" + vbCrLf + _
        "                else if (parameter is Int64)" + vbCrLf + _
        "                    finalParameters[i] = new IntPtr((Int64)parameter);" + vbCrLf + _
        "                else" + vbCrLf + _
        "                    finalParameters[i] = (IntPtr)parameter;" + vbCrLf + _
        "            }" + vbCrLf + _
        "            else" + vbCrLf + _
        "                finalParameters[i] = methodParameters[i];" + vbCrLf + _
        "        }" + vbCrLf + _
        "" + vbCrLf + _
        "        return method.Invoke(instance, finalParameters);" + vbCrLf + _
        "    }" + vbCrLf + _
        "" + vbCrLf
        
    code = code + _
        "    public Object GetField(Type fieldOwnerType, String fieldName)" + vbCrLf + _
        "    {" + vbCrLf + _
        "        try" + vbCrLf + _
        "        {" + vbCrLf + _
        "            return fieldOwnerType.GetField(fieldName).GetValue(null);" + vbCrLf + _
        "        }" + vbCrLf + _
        "        catch (Exception e)" + vbCrLf + _
        "        {" + vbCrLf + _
        "            LogException(e);" + vbCrLf + _
        "            throw;" + vbCrLf + _
        "        }" + vbCrLf + _
        "    }" + vbCrLf + _
        "" + vbCrLf

    code = code + _
        "    public Object InvokeMember(Type methodOwnerType, String methodName, BindingFlags bindingFlags, Object instance, Object[] methodParameters)" + vbCrLf + _
        "    {" + vbCrLf + _
        "        try" + vbCrLf + _
        "        {" + vbCrLf + _
        "            return methodOwnerType.InvokeMember(methodName, bindingFlags, null, instance, methodParameters);" + vbCrLf + _
        "        }" + vbCrLf + _
        "        catch (Exception e)" + vbCrLf + _
        "        {" + vbCrLf + _
        "            LogException(e);" + vbCrLf + _
        "            throw;" + vbCrLf + _
        "        }" + vbCrLf + _
        "    }" + vbCrLf + _
        "" + vbCrLf

    code = code + _
        "    public Object InvokeInstanceMember(Object instance, String methodName, Object[] methodParameters)" + vbCrLf + _
        "    {" + vbCrLf + _
        "        try" + vbCrLf + _
        "        {" + vbCrLf + _
        "            return InvokeMember(instance.GetType(), methodName, BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public, instance, methodParameters);" + vbCrLf + _
        "        }" + vbCrLf + _
        "        catch (Exception e)" + vbCrLf + _
        "        {" + vbCrLf + _
        "            LogException(e);" + vbCrLf + _
        "            throw;" + vbCrLf + _
        "        }" + vbCrLf + _
        "    }" + vbCrLf + _
        "" + vbCrLf

    code = code + _
        "    public String GetTypeName(Type type)" + vbCrLf + _
        "    {" + vbCrLf + _
        "        try" + vbCrLf + _
        "        {" + vbCrLf + _
        "            return type.ToString();" + vbCrLf + _
        "        }" + vbCrLf + _
        "        catch (Exception e)" + vbCrLf + _
        "        {" + vbCrLf + _
        "            LogException(e);" + vbCrLf + _
        "            throw;" + vbCrLf + _
        "        }" + vbCrLf + _
        "    }" + vbCrLf + _
        "" + vbCrLf

    code = code + _
        "    public Type GetObjectType(Object instance)" + vbCrLf + _
        "    {" + vbCrLf + _
        "        try" + vbCrLf + _
        "        {" + vbCrLf + _
        "            return instance.GetType();" + vbCrLf + _
        "        }" + vbCrLf + _
        "        catch (Exception e)" + vbCrLf + _
        "        {" + vbCrLf + _
        "            LogException(e);" + vbCrLf + _
        "            throw;" + vbCrLf + _
        "        }" + vbCrLf + _
        "    }" + vbCrLf + _
        "" + vbCrLf
        
    code = code + _
        "    public void SetAssemblyResolver(AppDomain domain, String path)" + vbCrLf + _
        "    {" + vbCrLf + _
        "        if (domain == null)" + vbCrLf + _
        "            throw new ArgumentException(""Invalid domain specified."");" + vbCrLf + _
        "" + vbCrLf + _
        "        if (String.IsNullOrWhiteSpace(path) || !Directory.Exists(path))" + vbCrLf + _
        "            throw new ArgumentException(""Invalid path specified."");" + vbCrLf + _
        "" + vbCrLf + _
        "        domain.AssemblyResolve += (sender, args) =>" + vbCrLf + _
        "        {" + vbCrLf + _
        "            if (String.IsNullOrWhiteSpace(args.Name) || args.Name.Contains("".resources""))" + vbCrLf + _
        "                return null;" + vbCrLf + _
        "" + vbCrLf + _
        "            foreach (Assembly assembly in domain.GetAssemblies())" + vbCrLf + _
        "            {" + vbCrLf + _
        "                if (assembly.FullName == args.Name)" + vbCrLf + _
        "                    return assembly;" + vbCrLf + _
        "            }" + vbCrLf + _
        "" + vbCrLf + _
        "            String assemblyPath = Path.Combine(path, String.Concat(args.Name.Split(new Char[] {','})[0], "".dll""));" + vbCrLf + _
        "            try { return Assembly.LoadFrom(assemblyPath); } catch { return null; }" + vbCrLf + _
        "        };" + vbCrLf + _
        "    }" + vbCrLf

    code = code + "}"

    Dim asm As mscorlib.Assembly: Set asm = CompileAssembly(code)
    Set CreateManagedHelper = ToObject(asm.CreateInstance(mh))

End Function

' METHOD
' Creates a .NET Runtime Host session through the RuntimeLoader library.

Private Function CreateRuntimeHost() As mscoree.CorRuntimeHost

    Dim shortPath As String: shortPath = String$(260, vbNull)
    Dim pathLength As Long: pathLength = GetShortPathName(StrPtr(m_LibrariesPath), StrPtr(shortPath), 260)
    
    If (pathLength = 0) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateRuntimeHost", "The call to the native function GetShortPathName failed.")
    End If
    
    shortPath = Left$(shortPath, pathLength)

    Dim hr As Long: hr = SetDefaultDllDirectories(&H1000)
    
    If (hr = 0) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateRuntimeHost", "The call to the native function SetDefaultDllDirectories failed.")
    End If
    
    Call AddDllDirectory(StrPtr(shortPath))

    Dim crh As mscoree.CorRuntimeHost
    hr = LoadRuntime(crh)
    
    If (hr < 0) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateRuntimeHost", "The creation of the .NET Runtime Host session failed.")
    End If

    Set CreateRuntimeHost = crh
    
End Function

' METHOD
' Loads the .NET Strata Wrapper into the current domain.

Private Function LoadStrata() As mscorlib.Assembly

    Dim asm As mscorlib.Assembly: Set asm = Nothing
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

    Dim ikrPath As String: ikrPath = m_LibrariesPath & "\IKVM.Runtime.dll"
    
    If Not fso.FileExists(ikrPath) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.LoadStrata", "The library IKVM.Runtime.dll could not be found.")
    End If

    Dim ikrAssemblyName As mscorlib.AssemblyName: Set ikrAssemblyName = InvokeMethodStaticFromName("System.Reflection.AssemblyName", "GetAssemblyName", ikrPath)
    
    If ikrAssemblyName Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.LoadStrata", "The assembly name of the library IKVM.Runtime.dll could not be retrieved.")
    End If
    
    Dim ikrFullName As String: ikrFullName = ToObject(ikrAssemblyName).FullName
    
    If (ikrFullName = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.LoadStrata", "The fully qualified name of the library IKVM.Runtime.dll could not be retrieved.")
    End If

    Set asm = m_Domain.Load_2(ikrFullName)

    Dim ikcPath As String: ikcPath = m_LibrariesPath & "\IKVM.OpenJDK.Core.dll"
    
    If Not fso.FileExists(ikcPath) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.LoadStrata", "The library IKVM.OpenJDK.Core.dll could not be found.")
    End If

    Dim ikcAssemblyName As mscorlib.AssemblyName: Set ikcAssemblyName = InvokeMethodStaticFromName("System.Reflection.AssemblyName", "GetAssemblyName", ikcPath)
    
    If ikcAssemblyName Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.LoadStrata", "The assembly name of the library IKVM.OpenJDK.Core.dll could not be retrieved.")
    End If
    
    Dim ikcFullName As String: ikcFullName = ToObject(ikcAssemblyName).FullName
    
    If (ikcFullName = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.LoadStrata", "The fully qualified name of the library IKVM.OpenJDK.Core.dll could not be retrieved.")
    End If

    Set asm = m_Domain.Load_2(ikcFullName)

    Dim strataPath As String

    #If Win64 Then
        strataPath = m_LibrariesPath & "\Strata64.dll"
    #Else
        strataPath = m_LibrariesPath & "\Strata32.dll"
    #End If

    If Not fso.FileExists(strataPath) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.LoadStrata", "The Strata wrapper could not be found.")
    End If

    Dim strataAssemblyName As mscorlib.AssemblyName: Set strataAssemblyName = InvokeMethodStaticFromName("System.Reflection.AssemblyName", "GetAssemblyName", strataPath)

    If strataAssemblyName Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.LoadStrata", "The assembly name of the Strata wrapper could not be retrieved.")
    End If

    Dim strataFullName As String: strataFullName = ToObject(strataAssemblyName).FullName
    
    If (strataFullName = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.LoadStrata", "The fully qualified name of the Strata wrapper could not be retrieved.")
    End If

    Set asm = m_Domain.Load_2(strataFullName)
    Call InvokeMethodStaticFromName("ikvm.runtime.Startup", "addBootClassPathAssembly", asm)

    Set LoadStrata = asm

End Function

' METHOD
' Checks whether the given object is a .NET Type.

Private Function IsType(ByRef v As Variant) As Boolean

    IsType = (VarType(v) = vbDataObject) And TypeOf v Is mscorlib.Type

End Function

' METHOD
' Performs a cleaup of the instance members.

Private Sub PerformCleanup()

    m_Initialized = False
    
    m_LibrariesPath = vbNullString
    m_LogFile = vbNullString
    m_Suffix = vbNullString

    Set m_TypesCache = Nothing
    Set m_ManagedHelper = Nothing

    Set m_AssemblyCore = Nothing
    Set m_AssemblyStrata = Nothing
    Set m_AssemblySystem = Nothing

    If Not m_Domain Is Nothing Then
    
        If Not m_Host Is Nothing Then
            Call m_Host.UnloadDomain(m_Domain)
        End If

        Set m_Domain = Nothing

    End If
    
    If Not m_Host Is Nothing Then

        Call m_Host.Stop
        Set m_Host = Nothing

    End If

End Sub

' METHOD
' Sets the value of the given .NET object instance property.

Public Sub SetProperty(ByVal instance As mscorlib.Object, ByVal propertyName As String, ByVal propertyValue As Variant, ParamArray propertyParameters() As Variant)

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.SetProperty", "This instance has not been initialized yet.")
    End If

    If instance Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.SetProperty", "The ""instance"" parameter is null.")
    End If

    propertyName = Trim$(propertyName)

    If (propertyName = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.SetProperty", "The ""propertyName"" parameter is null or empty.")
    End If

    Dim lb As Long: lb = LBound(propertyParameters)
    Dim ub As Long: ub = UBound(propertyParameters)
    Dim c As Long: c = ub - lb + 1
    
    Dim params() As Variant: ReDim params(0 To c)
    Call VariantCopy(params(0), propertyValue)

    Dim i As Long: i = 0
    Dim j As Long: j = lb

    For i = 1 To c
        Call VariantCopy(params(i), propertyParameters(j))
        j = j + 1
    Next i

    Call m_ManagedHelper.InvokeMember _
    ( _
        GetType(instance), _
        propertyName, _
        BindingFlags_SetProperty Or BindingFlags_Instance Or BindingFlags_Public, _
        instance, _
        params _
    )

End Sub

' METHOD
' Resolves the given class, name, binding flags and argument types into the correct .NET method overload.

Private Function ResolveWrapper(ByVal methodOwnerType As mscorlib.Type, ByVal methodName As String, ByVal bf As mscorlib.BindingFlags, ByRef argumentTypes() As Variant) As mscorlib.MethodInfo

    Dim i As Long, j As Long

    Dim lb As Long: lb = LBound(argumentTypes)
    Dim ub As Long: ub = UBound(argumentTypes)
    Dim c As Long

    For i = lb To ub
        If IsType(argumentTypes(i)) Then
            c = c + 1
        End If
    Next i

    Dim mis() As mscorlib.MethodInfo: mis = methodOwnerType.GetMethods(bf)
    Dim mi As mscorlib.MethodInfo
    Dim matched As Boolean

    For i = LBound(mis) To UBound(mis)

        Set mi = mis(i)

        If (mi.Name = methodName) Then

            Dim pi() As mscorlib.ParameterInfo: pi = mi.GetParameters()

            If ((UBound(pi) - LBound(pi) + 1) = c) Then

                matched = True
                
                Dim k As Long: k = LBound(pi)

                For j = lb To ub

                    If IsType(argumentTypes(j)) Then

                        Dim piObject As mscorlib.Object: Set piObject = pi(k)
                        Dim o As Object: Set o = piObject
                        Dim piType As mscorlib.Type: Set piType = o.ParameterType

                        If Not piType.Equals(argumentTypes(j)) Then
                            matched = False
                            Exit For
                        End If
 
                        k = k + 1

                    End If

                Next j

                If matched Then
                    Exit For
                End If

            End If

        End If

        Set mi = Nothing

    Next i

    Set ResolveWrapper = mi

End Function

' METHOD
' Creates a .NET assembly from the given source code.

Public Function CompileAssembly(ByVal code As String, ParamArray referencedAssemblyNames() As Variant) As mscorlib.Assembly

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CompileAssembly", "This instance has not been initialized yet.")
    End If

    code = Trim$(code)

    If (code = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CompileAssembly", "The ""code"" parameter is null or empty.")
    End If

    Dim o As Object:

    Dim compilerParameters As mscorlib.Object: Set compilerParameters = m_AssemblySystem.CreateInstance("System.CodeDom.Compiler.CompilerParameters")
    Set o = compilerParameters
    o.CompilerOptions = "/optimize"
    o.GenerateInMemory = True

    Dim refAssemblies As Object: Set refAssemblies = ToObject(o.ReferencedAssemblies)

    Dim s As String, v As Variant

    For Each v In referencedAssemblyNames
    
        If (VarType(v) <> vbString) Then
            Call Err.Raise(vbObjectError + 1, "RuntimeHost.CompileAssembly", "The ""referencedAssemblyNames"" contains invalid values.")
        End If
        
        s = Trim$(CStr(v))
        
        If (s = vbNullString) Then
            Call Err.Raise(vbObjectError + 1, "RuntimeHost.CompileAssembly", "The ""referencedAssemblyNames"" contains invalid values.")
        End If
    
        Call refAssemblies.Add(s)

    Next v
    
    Dim codeParameters() As String: ReDim codeParameters(0)
    codeParameters(0) = code
    
    Dim compiler As mscorlib.Object: Set compiler = m_AssemblySystem.CreateInstance("Microsoft.CSharp.CSharpCodeProvider")
    Set o = compiler

    Dim compilerResult As mscorlib.Object: Set compilerResult = o.CompileAssemblyFromSource(compilerParameters, codeParameters)
    Set o = compilerResult

    Dim compilerErrors As Object: Set compilerErrors = ToObject(o.Errors)

    If compilerErrors.HasErrors Then

        Dim i As Long
        Dim compilerError As Object

        For i = 0 To compilerErrors.Count - 1
            Set compilerError = ToObject(compilerErrors.Item(i))
            Debug.Print compilerError.ErrorText; Replace$(" [Line <#>]", "<#>", CStr(compilerError.Line))
        Next i

        Exit Function

    End If

    Set CompileAssembly = o.CompiledAssembly

End Function

' METHOD
' Creates a .NET delegate instance which calls the given method of the given instance.

Public Function CreateDelegateFromName(ByVal delegateType As String, ByVal instance As Object, ByVal instanceMethod As String) As mscorlib.Delegate

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateDelegateFromName", "This instance has not been initialized yet.")
    End If

    delegateType = Trim$(delegateType)

    If (delegateType = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateDelegateFromName", "The ""delegateType"" parameter is null or empty.")
    End If
    
    Dim resolvedType As mscorlib.Type: Set resolvedType = ResolveType(delegateType)
    
    If resolvedType Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateDelegateFromName", "The type specified by the ""delegateType"" parameter could not be resolved.")
    End If

    If instance Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateDelegateFromName", "The ""instance"" parameter is null.")
    End If
    
    instanceMethod = Trim$(instanceMethod)
    
    If (instanceMethod = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateDelegateFromName", "The ""instanceMethod"" parameter is null or empty.")
    End If

    Set CreateDelegateFromName = m_ManagedHelper.CreateDelegate(resolvedType, instance, instanceMethod)

End Function

' METHOD
' Creates a .NET delegate instance which calls the given method of the given instance.

Public Function CreateDelegateFromType(ByVal delegateType As mscorlib.Type, ByVal instance As Object, ByVal instanceMethod As String) As mscorlib.Delegate

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateDelegateFromType", "This instance has not been initialized yet.")
    End If

    If delegateType Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateDelegateFromType", "The ""delegateType"" parameter is null.")
    End If
    
    If instance Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateDelegateFromType", "The ""instance"" parameter is null.")
    End If
    
    instanceMethod = Trim$(instanceMethod)
    
    If (instanceMethod = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateDelegateFromType", "The ""instanceMethod"" parameter is null or empty.")
    End If

    Set CreateDelegateFromType = m_ManagedHelper.CreateDelegate(delegateType, instance, instanceMethod)

End Function

' METHOD
' Creates a .NET delegate instance which calls the given function.
' The "functionPointer" parameter must be passed using the AddressOf operator.
' The signature of the function pointer must match the delegate type.

Public Function CreateDelegateWithFunctionFromName(ByVal delegateType As String, ByVal functionPointer As LongPtr) As mscorlib.Delegate

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateDelegateWithFunctionFromName", "This instance has not been initialized yet.")
    End If

    delegateType = Trim$(delegateType)

    If (delegateType = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateDelegateWithFunctionFromName", "The ""delegateType"" parameter is null or empty.")
    End If

    Dim resolvedType As mscorlib.Type: Set resolvedType = ResolveType(delegateType)
    
    If resolvedType Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateDelegateWithFunctionFromName", "The type specified by the ""delegateType"" parameter could not be resolved.")
    End If

    Set CreateDelegateWithFunctionFromName = m_ManagedHelper.CreateDelegateWithFunction(resolvedType, functionPointer)

End Function

' METHOD
' Creates a .NET delegate instance which calls the given function.
' The "functionPointer" parameter must be passed using the AddressOf operator.
' The signature of the function pointer must match the delegate type.

Public Function CreateDelegateWithFunctionFromType(ByVal delegateType As mscorlib.Type, ByVal functionPointer As LongPtr) As mscorlib.Delegate

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateDelegateWithFunctionFromType", "This instance has not been initialized yet.")
    End If

    If delegateType Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateDelegateWithFunctionFromType", "The ""delegateType"" parameter is null.")
    End If

    Set CreateDelegateWithFunctionFromType = m_ManagedHelper.CreateDelegateWithFunction(delegateType, functionPointer)

End Function
 
' METHOD
' Creates a .NET enumerator instance of given type and value.

Public Function CreateEnumFromName(ByVal enumType As String, ByVal enumValue As Variant) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateEnumFromName", "This instance has not been initialized yet.")
    End If

    enumType = Trim$(enumType)

    If (enumType = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateEnumFromName", "The ""enumType"" parameter is null or empty.")
    End If

    Dim resolvedType As mscorlib.Type: Set resolvedType = ResolveType(enumType)
    
    If resolvedType Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateEnumFromName", "The type specified by the ""enumType"" parameter could not be resolved.")
    End If

    Select Case VarType(enumValue)

        Case vbByte
            CreateEnumFromName = InvokeMethodStaticFromName("System.Enum", "ToObject", resolvedType, CByte(enumValue))

        Case vbInteger
            CreateEnumFromName = InvokeMethodStaticFromName("System.Enum", "ToObject", resolvedType, CInt(enumValue))

        Case vbLong
            CreateEnumFromName = InvokeMethodStaticFromName("System.Enum", "ToObject", resolvedType, CLngPtr(enumValue))

        Case vbString

            Dim valueString As String: valueString = Trim$(CStr(enumValue))
            
            If (valueString = vbNullString) Then
                Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateEnumFromName", "The ""enumValue"" parameter is invalid.")
            End If
            
            CreateEnumFromName = InvokeMethodStaticFromName("System.Enum", "Parse", resolvedType, valueString)

        Case Else
            Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateEnumFromName", "The ""enumValue"" parameter is invalid.")

    End Select

End Function

' METHOD
' Creates a .NET enumerator instance of given type and value.

Public Function CreateEnumFromType(ByVal enumType As mscorlib.Type, ByVal enumValue As Variant) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateEnumFromType", "This instance has not been initialized yet.")
    End If

    If enumType Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateEnumFromType", "The ""enumType"" parameter is null.")
    End If

    Select Case VarType(enumValue)

        Case vbByte
            CreateEnumFromType = InvokeMethodStaticFromName("System.Enum", "ToObject", enumType, CByte(enumValue))

        Case vbInteger
            CreateEnumFromType = InvokeMethodStaticFromName("System.Enum", "ToObject", enumType, CInt(enumValue))

        Case vbLong
            CreateEnumFromType = InvokeMethodStaticFromName("System.Enum", "ToObject", enumType, CLngPtr(enumValue))

        Case vbString

            Dim valueString As String: valueString = Trim$(CStr(enumValue))
            
            If (valueString = vbNullString) Then
                Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateEnumFromType", "The ""enumValue"" parameter is invalid.")
            End If

            CreateEnumFromType = InvokeMethodStaticFromName("System.Enum", "Parse", enumType, valueString)
 
        Case Else
            Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateEnumFromType", "The ""enumValue"" parameter is invalid.")

    End Select

End Function

' METHOD
' Creates a .NET object instance of the given type.

Public Function CreateObjectFromName(ByVal objectType As String) As mscorlib.Object

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateObjectFromName", "This instance has not been initialized yet.")
    End If

    objectType = Trim$(objectType)

    If (objectType = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateEnumFromName", "The ""objectType"" parameter is null or empty.")
    End If

    Dim resolvedType As mscorlib.Type: Set resolvedType = ResolveType(objectType)
    
    If resolvedType Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateEnumFromName", "The type specified by the ""objectType"" parameter could not be resolved.")
    End If

    Set CreateObjectFromName = ToObject(resolvedType.Assembly).CreateInstance_2(objectType, False)

End Function

' METHOD
' Creates a .NET object instance of the given type.

Public Function CreateObjectFromType(ByVal objectType As mscorlib.Type) As mscorlib.Object

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateObjectFromType", "This instance has not been initialized yet.")
    End If

    If objectType Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateObjectFromType", "The ""objectType"" parameter is null.")
    End If

    Set CreateObjectFromType = ToObject(objectType.Assembly).CreateInstance_2(objectType.ToString(), False)

End Function

' METHOD
' Creates a .NET object instance of the given type, with the given parameters.

Public Function CreateObjectWithParamsFromName(ByVal objectType As String, ParamArray objectParameters() As Variant) As mscorlib.Object

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateObjectWithParamsFromName", "This instance has not been initialized yet.")
    End If

    objectType = Trim$(objectType)

    If (objectType = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateObjectWithParamsFromName", "The ""objectType"" parameter is null or empty.")
    End If

    Dim resolvedType As mscorlib.Type: Set resolvedType = ResolveType(objectType)
    
    If resolvedType Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateObjectWithParamsFromName", "The type specified by the ""objectType"" parameter could not be resolved.")
    End If

    Dim params() As Variant: params = objectParameters

    Set CreateObjectWithParamsFromName = ToObject(resolvedType.Assembly).CreateInstance_3 _
    ( _
        objectType, _
        False, _
        BindingFlags_CreateInstance Or BindingFlags_Instance Or BindingFlags_Public, _
        Nothing, _
        params, _
        Nothing, _
        Array() _
    )

End Function

' METHOD
' Creates a .NET object instance of the given type, with the given parameters.

Public Function CreateObjectWithParamsFromType(ByVal objectType As mscorlib.Type, ParamArray objectParameters() As Variant) As mscorlib.Object

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateObjectWithParamsFromType", "This instance has not been initialized yet.")
    End If

    If objectType Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.CreateObjectWithParamsFromType", "The ""objectType"" parameter is null.")
    End If

    Dim params() As Variant: params = objectParameters

    Set CreateObjectWithParamsFromType = ToObject(objectType.Assembly).CreateInstance_3 _
    ( _
        objectType.ToString(), _
        False, _
        BindingFlags_CreateInstance Or BindingFlags_Instance Or BindingFlags_Public, _
        Nothing, _
        params, _
        Nothing, _
        Array() _
    )

End Function

' METHOD
' Gets the value of the given .NET static field, defined on the given type.

Public Function GetFieldStaticByName(ByVal fieldOwner As String, ByVal fieldName As String) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetFieldStaticByName", "This instance has not been initialized yet.")
    End If

    fieldOwner = Trim$(fieldOwner)

    If (fieldOwner = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetFieldStaticByName", "The ""fieldOwner"" parameter is null or empty.")
    End If
    
    Dim fieldOwnerType As mscorlib.Type: Set fieldOwnerType = ResolveType(fieldOwner)
    
    If fieldOwnerType Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetFieldStaticByName", "The type specified by the ""fieldOwner"" parameter could not be resolved.")
    End If
    
    fieldName = Trim$(fieldName)

    If (fieldName = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetFieldStaticByName", "The ""fieldName"" parameter is null or empty.")
    End If

    Dim hr As Long: hr = VariantCopy _
    ( _
        GetFieldStaticByName, m_ManagedHelper.GetField _
        ( _
            fieldOwnerType, _
            fieldName _
        ) _
    )

    If (hr < 0) Then
        Call Err.Raise(hr)
    End If

End Function

' METHOD
' Gets the value of the given .NET static field, defined on the given type.

Public Function GetFieldStaticByType(ByVal fieldOwnerType As mscorlib.Type, ByVal fieldName As String) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetFieldStaticByType", "This instance has not been initialized yet.")
    End If

    If fieldOwnerType Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetFieldStaticByType", "The ""fieldOwnerType"" parameter is null.")
    End If
    
    fieldName = Trim$(fieldName)

    If (fieldName = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetFieldStaticByType", "The ""fieldName"" parameter is null or empty.")
    End If

    Dim hr As Long: hr = VariantCopy _
    ( _
        GetFieldStaticByType, m_ManagedHelper.GetField _
        ( _
            fieldOwnerType, _
            fieldName _
        ) _
    )

    If (hr < 0) Then
        Call Err.Raise(hr)
    End If

End Function

' METHOD
' Gets the value of the given .NET property, defined on the given instance.

Public Function GetProperty(ByVal instance As mscorlib.Object, ByVal propertyName As String, ParamArray propertyParameters() As Variant) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetProperty", "This instance has not been initialized yet.")
    End If

    If instance Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetProperty", "The ""instance"" parameter is null.")
    End If

    propertyName = Trim$(propertyName)

    If (propertyName = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetProperty", "The ""propertyName"" parameter is null or empty.")
    End If

    Dim lb As Long: lb = LBound(propertyParameters)
    Dim ub As Long: ub = UBound(propertyParameters)
    Dim c As Long: c = ub - lb + 1
    
    Dim params() As Variant: params = propertyParameters

    Dim hr As Long: hr = VariantCopy _
    ( _
        GetProperty, m_ManagedHelper.InvokeMember _
        ( _
            GetType(instance), _
            propertyName, _
            BindingFlags_GetProperty Or BindingFlags_Instance Or BindingFlags_Public, _
            instance, _
            params _
        ) _
    )

    If (hr < 0) Then
        Call Err.Raise(hr)
    End If

End Function

' METHOD
' Gets the value of the given .NET static property, defined on the given type.

Public Function GetPropertyStaticFromName(ByVal propertyOwner As String, ByVal propertyName As String, ParamArray propertyParameters() As Variant) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetPropertyStaticFromName", "This instance has not been initialized yet.")
    End If

    propertyOwner = Trim$(propertyOwner)

    If (propertyOwner = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetPropertyStaticFromName", "The ""propertyOwner"" parameter is null or empty.")
    End If
    
    Dim propertyOwnerType As mscorlib.Type: Set propertyOwnerType = ResolveType(propertyOwner)
    
    If propertyOwnerType Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetPropertyStaticFromName", "The type specified by the ""propertyOwner"" parameter could not be resolved.")
    End If
    
    propertyName = Trim$(propertyName)

    If (propertyName = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetPropertyStaticFromName", "The ""propertyName"" parameter is null or empty.")
    End If

    Dim params() As Variant: params = propertyParameters

    Dim hr As Long: hr = VariantCopy _
    ( _
        GetPropertyStaticFromName, m_ManagedHelper.InvokeMember _
        ( _
            propertyOwnerType, _
            propertyName, _
            BindingFlags_GetProperty Or BindingFlags_Public Or BindingFlags_Static, _
            Nothing, _
            params _
        ) _
    )

    If (hr < 0) Then
        Call Err.Raise(hr)
    End If

End Function

' METHOD
' Gets the value of the given .NET static property, defined on the given type.

Public Function GetPropertyStaticFromType(ByVal propertyOwnerType As mscorlib.Type, ByVal propertyName As String, ParamArray propertyParameters() As Variant) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetPropertyStaticFromType", "This instance has not been initialized yet.")
    End If

    If propertyOwnerType Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetPropertyStaticFromType", "The ""propertyOwnerType"" parameter is null.")
    End If
    
    propertyName = Trim$(propertyName)

    If (propertyName = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetPropertyStaticFromType", "The ""propertyName"" parameter is null or empty.")
    End If

    Dim params() As Variant: params = propertyParameters

    Dim hr As Long: hr = VariantCopy _
    ( _
        GetPropertyStaticFromType, m_ManagedHelper.InvokeMember _
        ( _
            propertyOwnerType, _
            propertyName, _
            BindingFlags_GetProperty Or BindingFlags_Public Or BindingFlags_Static, _
            Nothing, _
            params _
        ) _
    )

    If (hr < 0) Then
        Call Err.Raise(hr)
    End If

End Function

' METHOD
' Retrieves the type of the given .NET object instance.

Public Function GetType(ByVal instance As mscorlib.Object) As mscorlib.Type

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetType", "This instance has not been initialized yet.")
    End If

    If instance Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetType", "The ""instance"" parameter is null.")
    End If

    Set GetType = m_ManagedHelper.GetObjectType(instance)

End Function

' METHOD
' Retrieves the fully qualified type name of the given .NET object instance.

Public Function GetTypeNameFromObject(ByVal instance As mscorlib.Object) As String

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetTypeNameFromObject", "This instance has not been initialized yet.")
    End If

    If instance Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetTypeNameFromObject", "The ""instance"" parameter is null.")
    End If

    Dim objectType As mscorlib.Type: Set objectType = m_ManagedHelper.GetObjectType(instance)
    GetTypeNameFromObject = m_ManagedHelper.GetTypeName(objectType)

End Function

' METHOD
' Retrieves the fully qualified name of the given .NET type.

Public Function GetTypeNameFromType(ByVal targetType As mscorlib.Type) As String

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetTypeNameFromType", "This instance has not been initialized yet.")
    End If

    If targetType Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.GetTypeNameFromType", "The ""targetType"" parameter is null.")
    End If

    GetTypeNameFromType = m_ManagedHelper.GetTypeName(targetType)

End Function

' METHOD
' Initializes an instance of the RuntimeHost class.

Public Sub Initialize()

    If m_Initialized Then
        Exit Sub
    End If
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(m_LibrariesPath) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.Initialize", "The ""Libraries"" folder could not be found.")
    End If

    On Error GoTo ErrorHandler

    m_Initialized = True

    Set m_Host = CreateRuntimeHost()
    Call m_Host.Start

    Call m_Host.CreateDomain("StrataXL_" + m_Suffix, Nothing, m_Domain)
    Set m_AssemblyCore = m_Domain.Load_2("mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    Set m_AssemblySystem = m_Domain.Load_2("System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")

    Set m_ManagedHelper = CreateManagedHelper(m_Domain)
    Call m_ManagedHelper.SetAssemblyResolver(m_Domain, m_LibrariesPath)

    Set m_AssemblyStrata = LoadStrata()

    Exit Sub

ErrorHandler:

    Call PerformCleanup
    Call Err.Raise(Err.Number, Err.Source, Err.Description)

End Sub

' METHOD
' Calls the given method on the given .NET object instance.

Public Function InvokeMethod(ByVal instance As mscorlib.Object, ByVal methodName As String, ParamArray methodParameters() As Variant) As Variant
    
    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.InvokeMethod", "This instance has not been initialized yet.")
    End If
    
    If instance Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.InvokeMethod", "The ""instance"" parameter is null.")
    End If
    
    methodName = Trim$(methodName)

    If (methodName = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.InvokeMethod", "The ""methodName"" parameter is null or empty.")
    End If
    
    Dim params() As Variant: params = methodParameters

    Dim hr As Long: hr = VariantCopy _
    ( _
        InvokeMethod, m_ManagedHelper.InvokeInstanceMember _
        ( _
            instance, _
            methodName, _
            params _
        ) _
    )

    If (hr < 0) Then
        Call Err.Raise(hr)
    End If

End Function

' METHOD
' Calls the given static method, defined on the given type.

Public Function InvokeMethodStaticFromName(ByVal methodOwner As String, ByVal methodName As String, ParamArray methodParameters() As Variant) As Variant

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.InvokeMethodStaticFromName", "This instance has not been initialized yet.")
    End If

    methodOwner = Trim$(methodOwner)

    If (methodOwner = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.InvokeMethodStaticFromName", "The ""methodOwner"" parameter is null or empty.")
    End If
    
    Dim methodOwnerType As mscorlib.Type: Set methodOwnerType = ResolveType(methodOwner)
    
    If methodOwnerType Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.InvokeMethodStaticFromName", "The type specified by the ""methodOwnerType"" parameter could not be resolved.")
    End If
    
    methodName = Trim$(methodName)

    If (methodName = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.InvokeMethodStaticFromName", "The ""methodName"" parameter is null or empty.")
    End If

    Dim params() As Variant: params = methodParameters

    Dim hr As Long: hr = VariantCopy _
    ( _
        InvokeMethodStaticFromName, m_ManagedHelper.InvokeMember _
        ( _
            methodOwnerType, _
            methodName, _
            BindingFlags_InvokeMethod Or BindingFlags_Public Or BindingFlags_Static, _
            Nothing, _
            params _
        ) _
    )

    If (hr < 0) Then
        Call Err.Raise(hr)
    End If

End Function

' METHOD
' Calls the given static method, defined on the given type.

Public Function InvokeMethodStaticFromType(ByVal methodOwnerType As mscorlib.Type, ByVal methodName As String, ParamArray methodParameters() As Variant) As mscorlib.Object

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.InvokeMethodStaticFromType", "This instance has not been initialized yet.")
    End If

    If methodOwnerType Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.InvokeMethodStaticFromType", "The ""methodOwnerType"" parameter is null.")
    End If
    
    methodName = Trim$(methodName)

    If (methodName = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.InvokeMethodStaticFromType", "The ""methodName"" parameter is null or empty.")
    End If

    Dim params() As Variant: params = methodParameters
    
    Dim hr As Long: hr = VariantCopy _
    ( _
        InvokeMethodStaticFromType, m_ManagedHelper.InvokeMember _
        ( _
            methodOwnerType, _
            methodName, _
            BindingFlags_InvokeMethod Or BindingFlags_Public Or BindingFlags_Static, _
            Nothing, _
            params _
        ) _
    )

End Function

' METHOD
' Loads a .NET assembly into the current domain, from the given full name.

Public Sub LoadAssembly(ByVal asmFullName As String)

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.LoadAssembly", "This instance has not been initialized yet.")
    End If

    asmFullName = Trim$(asmFullName)

    If (asmFullName = vbNullString) Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.LoadAssembly", "The ""asmFullName"" parameter is null or empty.")
    End If

    Dim asm As mscorlib.Assembly: Set asm = m_Domain.Load_2(asmFullName)

End Sub

' METHOD
' Resolves a fully qualified type name into a .NET type.

Public Function ResolveType(ByVal typeName As String) As mscorlib.Type

    Set ResolveType = Nothing

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.ResolveType", "This instance has not been initialized yet.")
    End If

    typeName = Trim$(typeName)

    If (typeName = vbNullString) Then
        Exit Function
    End If

    On Error Resume Next

    Set ResolveType = m_TypesCache.Item(typeName)

    On Error GoTo 0

    If Not ResolveType Is Nothing Then
        Exit Function
    End If

    Dim asms() As mscorlib.Assembly: asms = m_Domain.GetAssemblies()
    
    Dim a As Variant
    Dim asm As mscorlib.Assembly

    For Each a In asms

        Set asm = a
        Set ResolveType = asm.GetType_2(typeName)

        If Not ResolveType Is Nothing Then
            Call m_TypesCache.Add(ResolveType, typeName)
            Exit Function
        End If

    Next a

End Function

' METHOD
' Resolves the given .NET origin type, member name and member type into a .NET type.

Public Function ResolveTypeByInheritance(ByVal originType As mscorlib.Type, ByVal memberName As String, ByVal memberType As mscorlib.MemberTypes) As mscorlib.Type

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.ResolveTypeByInheritance", "This instance has not been initialized yet.")
    End If

    If originType Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.ResolveTypeByInheritance", "The ""originType"" parameter is null.")
    End If

    Dim mis() As mscorlib.MemberInfo: mis = originType.GetMembers_2()
    
    Dim i As Long
    Dim mi As mscorlib.MemberInfo

    For i = LBound(mis) To UBound(mis)

        Set mi = mis(i)

        If ((mi.Name = memberName) And ((mi.memberType And MemberTypes_All) <> 0)) Then
            Set ResolveTypeByInheritance = originType
            Exit Function
        End If

    Next i

    If originType.BaseType Is Nothing Then
        Set ResolveTypeByInheritance = Nothing
        Exit Function
    End If

    Set ResolveTypeByInheritance = ResolveTypeByInheritance(originType.BaseType, memberName, memberType)

End Function

' METHOD
' Converts the given .NET object instance into a .NET enumerable.

Public Function ToEnumerable(ByVal instance As mscorlib.Object) As mscorlib.IEnumerable

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.ToEnumerable", "This instance has not been initialized yet.")
    End If

    If instance Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.ToEnumerable", "The ""instance"" parameter is null.")
    End If

    Set ToEnumerable = instance

End Function

' METHOD
' Converts the given .NET object into a VB object.

Public Function ToObject(ByVal instance As mscorlib.Object) As Object

    If Not m_Initialized Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.ToObject", "This instance has not been initialized yet.")
    End If

    If instance Is Nothing Then
        Call Err.Raise(vbObjectError + 1, "RuntimeHost.ToObject", "The ""instance"" parameter is null.")
    End If

    Set ToObject = instance

End Function
