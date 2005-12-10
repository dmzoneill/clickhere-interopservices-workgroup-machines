Imports System.Runtime.InteropServices

Module Module1

    Structure Computer_info_101

        Public Platform_ID As Integer

        <MarshalAsAttribute(UnmanagedType.LPWStr)> Public Name As String

        Public Version_Major As Integer

        Public Version_Minor As Integer

        Public Type As Integer

        <MarshalAsAttribute(UnmanagedType.LPWStr)> Public Comment As String

    End Structure

    Declare Unicode Function NetServerEnum Lib "Netapi32.dll" (ByVal Servername As Integer, ByVal level As Integer, ByRef buffer As Integer, ByVal PrefMaxLen As Integer, ByRef EntriesRead As Integer, ByRef TotalEntries As Integer, ByVal ServerType As Integer, ByVal DomainName As String, ByRef ResumeHandle As Integer) As Integer

    Declare Function NetApiBufferFree Lib "Netapi32.dll" (ByVal lpBuffer As Integer) As Integer

    Private Const SV_TYPE_SERVER As Integer = &H2 ' All Servers

    Sub Main()

        Dim ComputerInfo As Computer_info_101

        Dim i, MaxLenPref, level, ret, EntriesRead, TotalEntries, ResumeHandle As Integer

        Dim BufPtr As Integer

        Dim iPtr As IntPtr

        MaxLenPref = -1

        level = 101

        'replace "MKOSLOF" with  your workgroup name

        ret = NetServerEnum(0, level, BufPtr, MaxLenPref, EntriesRead, TotalEntries, SV_TYPE_SERVER, "CLICKHERE", ResumeHandle)

        If ret <> 0 Then
            Console.WriteLine("An Error has occured")
            Return
        End If

        ' loop thru the entries

        For i = 0 To EntriesRead - 1

            ' copy into our structure

            Dim ptr As IntPtr = New IntPtr(BufPtr)

            ComputerInfo = CType(Marshal.PtrToStructure(ptr, GetType(Computer_info_101)), Computer_info_101)

            BufPtr = BufPtr + Len(ComputerInfo)

            Console.WriteLine(ComputerInfo.Name)

        Next

        NetApiBufferFree(BufPtr)

        Console.Write("Press Enter to End")

        Dim s As String = Console.ReadLine()

    End Sub


End Module