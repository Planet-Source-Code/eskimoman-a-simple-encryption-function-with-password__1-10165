<div align="center">

## A Simple encryption function with Password


</div>

### Description

A simple encryption function with Password. The slightest change with the password will effect the text so that it is unreadable. This function also encrypts and decrypts depending on if you set EnDeCrypt Boolean to True or False.

I am sure you could crack a message from this code after a while but i think it is stronger than other simple cryption codes you will find on planet source code.

Sorry that I ain't REMed it but it is quite simple so you should understand it. I hope you like.
 
### More Info
 
Text for encryption, Password and either true or false for encrypt or decrypt.

All you need to do is paste the code into a project and create a Command button. I useally put the function into a module.

The En/Decrypted text.

None I know off


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Eskimoman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/eskimoman.md)
**Level**          |Intermediate
**User Rating**    |4.1 (29 globes from 7 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Encryption](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/encryption__1-48.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/eskimoman-a-simple-encryption-function-with-password__1-10165/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
Dim Text As String, Password As String
Text = "Hello"
Password = "Pass"
Print Text
Text = Crypt(Text, Password, True)
Print Text
Text = Crypt(Text, Password, False)
Print Text
End Sub
Public Function Crypt(Source As String, strPassword As String, EnDeCrypt As Boolean) As String
'EnDeCrypt True = Encrypt
'EnDeCrypt False = Decrypt
Dim intPassword As Long
Dim intCrypt As Long
For x = 1 To Len(strPassword)
 intPassword = intPassword + Asc(Mid$(strPassword, x, 1))
Next x
For x = 1 To Len(Source)
If EnDeCrypt = True Then
 intCrypt = Asc(Mid$(Source, x, 1)) + intPassword + x
 Do Until intCrypt <= 255
 intCrypt = intCrypt - 255
 Loop
Else
 intCrypt = Asc(Mid$(Source, x, 1)) - intPassword - x
 Do Until intCrypt > 0
 intCrypt = intCrypt + 255
 Loop
End If
Crypt = Crypt & Chr(intCrypt)
Next x
End Function
```

