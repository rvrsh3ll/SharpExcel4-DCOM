# SharpExcel4-DCOM
 Port of Invoke-Excel4DCOM https://github.com/outflanknl/Excel4-DCOM

 As an example, one could execute SharpExcel4-DCOM.exe through Cobalt Strike's Beacon "execute-assembly" module.

#### Needed Info
You'll need to replace the calc shellcode with your target architechture shellcode. If it 64bit office that you are targeting, be sure to change "bool office64 = false;" to true prior to compilation.

#### Example usage
beacon>execute-assembly /root/SharpExcel4-DCOM/SharpExcel4-DCOM.exe --computername remotehost.domain.local

