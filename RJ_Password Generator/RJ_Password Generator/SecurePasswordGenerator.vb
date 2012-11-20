'Copyright 2012, Raymond C. Borges Hink and Jarilyn M. Hernandez Jimenez
'This program is copyright under GNU GENERAL PUBLIC LICENSE Version 3

'This program is free software: you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation, either version 3 of the License, or
'(at your option) any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with this program.  If not, see <http://www.gnu.org/licenses/>.

Imports System
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography

Public Class SecurePasswordGenerator

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        If System.IO.Directory.Exists("C:\SPG") = False Then
            System.IO.Directory.CreateDirectory("C:\SPG")
            Dim TextFile As New StreamWriter("C:\SPG\passwordProfiles.txt")
            TextFile.Close()
        End If


        'If passwordProfile file is empty show message create new profile
        Dim srFileReader As System.IO.StreamReader
        Dim sInputLine As String

        If System.IO.File.Exists("C:\SPG\passwordProfiles.txt") = True Then
            Dim a As IO.FileInfo = New IO.FileInfo("C:\SPG\passwordProfiles.txt")
            If a.Length = 0 Then

                Dim a1 As String = ("Copyright 2012" & ControlChars.NewLine & "Raymond C. Borges Hink and Jarilyn M. Hernandez Jimenez")
                Dim b As String = (ControlChars.NewLine & "This program is copyright under GNU GENERAL PUBLIC LICENSE Version 3." & ControlChars.NewLine)
                Dim c As String = ("This program is free software: you can redistribute it and/or modify")
                Dim d As String = ("it under the terms of the GNU General Public License as published by")
                Dim e1 As String = ("the Free Software Foundation, either version 3 of the License, or")
                Dim f As String = ("(at your option) any later version.")

                Dim g As String = ("This program is distributed in the hope that it will be useful,")
                Dim h As String = ("but WITHOUT ANY WARRANTY; without even the implied warranty of")
                Dim i1 As String = ("MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE." & ControlChars.NewLine)
                Dim j As String = ("See the GNU General Public License for more details.")

                Dim k As String = ("You should have received a copy of the GNU General Public License")
                Dim l As String = ("along with this program.  If not, see <http://www.gnu.org/licenses/>.")

                MessageBox.Show(a1 & vbNewLine & b & vbNewLine & c & vbNewLine & d & vbNewLine & e1 & vbNewLine & f & vbNewLine & g & vbNewLine & h & vbNewLine & i1 & vbNewLine & j & vbNewLine & k & vbNewLine & l, "Copyright Information")
                MessageBox.Show("Create a new password profile to activate password profiles!")
            End If
            srFileReader = System.IO.File.OpenText("C:\SPG\passwordProfiles.txt")
            sInputLine = srFileReader.ReadLine()
            Dim I As Integer = 0
            Do Until sInputLine Is Nothing
                pwdProBox.Items.Insert(I, sInputLine)
                sInputLine = srFileReader.ReadLine()
                I = I + 1
            Loop
            srFileReader.Close()
        End If

        'Disable all fields
        GeneratePwdPassProfileButton.Enabled = False
        CopyPasswdFromPassProfileButton.Enabled = False
        pwdProBox.Enabled = True
        CheckedListBox1.Enabled = False
        GeneratePWDnewProfileButton.Enabled = False
        txtBoxResult.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        txtbox1.Enabled = False
        txtbox2.Enabled = False
        txtbox3.Enabled = False
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False

    End Sub


    Private Sub GeneratePWD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GeneratePWDnewProfileButton.Click

        Dim password As String
        Dim strongPWD As String

        'Password variable holds concatenation of text boxes
        password = TextBox1.Text & TextBox2.Text & TextBox3.Text

        'start string to byte encoding conversion
        'encodes password from string to byte type and stores in bytePassword variable
        Dim encText As New System.Text.UTF8Encoding()
        Dim bytePassword() As Byte

        'password in byte form
        bytePassword = encText.GetBytes(password)

        'shows lentgh in bytes of bytepassword 
        'not used in conversion
        bytePassword.Length.ToString()
        'end of string to byte encoding conversion


        'START salt value generator
        ' Define min and max salt sizes.
        Dim minSaltSize As Integer
        Dim maxSaltSize As Integer

        minSaltSize = 510
        maxSaltSize = 512

        ' Generate a random number for the size of the salt.
        Dim random As Random
        random = New Random()

        Dim saltSize As Integer
        saltSize = random.Next(minSaltSize, maxSaltSize)

        ' Allocate a byte array, which will hold the salt.
        Dim saltBytes = New Byte(saltSize - 1) {}

        ' Initialize a random number generator.
        Dim rng As RNGCryptoServiceProvider
        rng = New RNGCryptoServiceProvider()

        ' Fill the salt with cryptographically strong byte values.
        rng.GetNonZeroBytes(saltBytes)
        'END of salt value generation

        'CONCATENATE salt with password then hash with SHA512
        ' Allocate array, which will hold plain text and salt.
        Dim passwordWithSaltBytes() As Byte = New Byte(bytePassword.Length + saltBytes.Length - 1) {}

        ' Copy plain text bytes into resulting array.
        Dim I As Integer
        For I = 0 To bytePassword.Length - 1
            passwordWithSaltBytes(I) = bytePassword(I)
        Next I

        ' Append salt bytes to the resulting array.
        For I = 0 To saltBytes.Length - 1
            passwordWithSaltBytes(bytePassword.Length + I) = saltBytes(I)
        Next I
        'End concatenation of salt and hashed password


        'START SHA 512 conversion
        Dim finalresult() As Byte
        Dim shaM As New SHA512Managed()

        'computes SHA-512 for bytePassword and stores in result variable of type byte
        finalresult = shaM.ComputeHash(passwordWithSaltBytes)
        'END of SHA-512 conversion

        'Convert byte array to base64 string and print to textbox
        strongPWD = Convert.ToBase64String(finalresult)

        'Trim password
        Dim MyChar() As Char = {"=", "="}
        strongPWD = strongPWD.TrimEnd(MyChar)


        ' Start of Checked List box code
        Dim illegalChars As Char() = ""
        Dim sb As New System.Text.StringBuilder

        'If all selected
        If CheckedListBox1.GetItemChecked(0) = True And CheckedListBox1.GetItemChecked(1) = True And CheckedListBox1.GetItemChecked(2) = True And CheckedListBox1.GetItemChecked(3) = True Then
            'Leave the same
        End If
        'End of all selected


        'If only alpha-numeric 
        If CheckedListBox1.GetItemChecked(0) = True And CheckedListBox1.GetItemChecked(1) = True And CheckedListBox1.GetItemChecked(2) = True And CheckedListBox1.GetItemChecked(3) = False Then
            illegalChars = "+/".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of alpha-numeric



        'If only alphabetic and symbols
        If CheckedListBox1.GetItemChecked(0) = True And CheckedListBox1.GetItemChecked(1) = True And CheckedListBox1.GetItemChecked(2) = False And CheckedListBox1.GetItemChecked(3) = True Then
            illegalChars = "0123456789".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of alphabetic and symbols




        'If only alphabetic uppercase
        If CheckedListBox1.GetItemChecked(0) = True And CheckedListBox1.GetItemChecked(1) = False And CheckedListBox1.GetItemChecked(2) = False And CheckedListBox1.GetItemChecked(3) = False Then
            strongPWD = strongPWD.ToUpper()
            illegalChars = "+/0123546789".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of alphabetic uppercase


        'If only alphabetic lowercase
        If CheckedListBox1.GetItemChecked(0) = False And CheckedListBox1.GetItemChecked(1) = True And CheckedListBox1.GetItemChecked(2) = False And CheckedListBox1.GetItemChecked(3) = False Then
            strongPWD = strongPWD.ToLower()
            illegalChars = "+/0123546789".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of alphabetic lowercase



        'If only alphabetic 
        If CheckedListBox1.GetItemChecked(0) = True And CheckedListBox1.GetItemChecked(1) = True And CheckedListBox1.GetItemChecked(2) = False And CheckedListBox1.GetItemChecked(3) = False Then
            illegalChars = "+/0123546789".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of alphabetic 



        'If only alphabetic uppercase and numbers
        If CheckedListBox1.GetItemChecked(0) = True And CheckedListBox1.GetItemChecked(1) = False And CheckedListBox1.GetItemChecked(2) = True And CheckedListBox1.GetItemChecked(3) = False Then
            strongPWD = strongPWD.ToUpper()
            illegalChars = "+/".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of only alphabetic uppercase and numbers



        'If only alphabetic uppercase, symbols and numbers
        If CheckedListBox1.GetItemChecked(0) = True And CheckedListBox1.GetItemChecked(1) = False And CheckedListBox1.GetItemChecked(2) = True And CheckedListBox1.GetItemChecked(3) = True Then
            strongPWD = strongPWD.ToUpper()
            illegalChars = "".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of only alphabetic uppercase, symbols and numbers



        'If only alphabetic lowercase, symbols and numbers
        If CheckedListBox1.GetItemChecked(0) = False And CheckedListBox1.GetItemChecked(1) = True And CheckedListBox1.GetItemChecked(2) = True And CheckedListBox1.GetItemChecked(3) = True Then
            strongPWD = strongPWD.ToLower()
            illegalChars = "".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of only alphabetic lowercase, symbols and numbers



        'If only alphabetic lowercase and numbers
        If CheckedListBox1.GetItemChecked(0) = False And CheckedListBox1.GetItemChecked(1) = True And CheckedListBox1.GetItemChecked(2) = True And CheckedListBox1.GetItemChecked(3) = False Then
            strongPWD = strongPWD.ToLower()
            illegalChars = "+/".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of alphabetic lowercase and numbers



        'If only numeric
        If CheckedListBox1.GetItemChecked(0) = False And CheckedListBox1.GetItemChecked(1) = False And CheckedListBox1.GetItemChecked(2) = True And CheckedListBox1.GetItemChecked(3) = False Then
            illegalChars = "+/abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of only numeric



        'If only symbols
        If CheckedListBox1.GetItemChecked(0) = False And CheckedListBox1.GetItemChecked(1) = False And CheckedListBox1.GetItemChecked(2) = False And CheckedListBox1.GetItemChecked(3) = True Then
            illegalChars = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of only symbols


        'If only symbols and numbers
        If CheckedListBox1.GetItemChecked(0) = False And CheckedListBox1.GetItemChecked(1) = False And CheckedListBox1.GetItemChecked(2) = True And CheckedListBox1.GetItemChecked(3) = True Then
            illegalChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of symbols and numbers



        'If none
        If CheckedListBox1.GetItemChecked(0) = False And CheckedListBox1.GetItemChecked(1) = False And CheckedListBox1.GetItemChecked(2) = False And CheckedListBox1.GetItemChecked(3) = False Then
            strongPWD = ""
        End If
        'End of none



        'If only symbols and uppercase
        If CheckedListBox1.GetItemChecked(0) = True And CheckedListBox1.GetItemChecked(1) = False And CheckedListBox1.GetItemChecked(2) = False And CheckedListBox1.GetItemChecked(3) = True Then
            strongPWD = strongPWD.ToUpper()
            illegalChars = "0123456789".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of symbols and uppercase



        'If only symbols and lowercase
        If CheckedListBox1.GetItemChecked(0) = False And CheckedListBox1.GetItemChecked(1) = True And CheckedListBox1.GetItemChecked(2) = False And CheckedListBox1.GetItemChecked(3) = True Then
            strongPWD = strongPWD.ToLower()
            illegalChars = "0123456789".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of symbols and lowercase


        'Write password to textbox
        txtBoxResult.Text = strongPWD
        'End of password generating code


        ' Create a file and write the byte data SALT to a file.
        Dim oFileStream As System.IO.FileStream
        oFileStream = New System.IO.FileStream("C:\SPG\" + pwdProfileBox.Text + ".dat", System.IO.FileMode.Create)
        oFileStream.Write(saltBytes, 0, saltBytes.Length)
        oFileStream.Close()
        'END of salt file writing


        'Write the new password profile to the list of profiles.
        If System.IO.File.Exists("C:\SPG\passwordProfiles.txt") = True Then
            Dim sw As StreamWriter
            sw = File.AppendText("C:\SPG\passwordProfiles.txt")
            sw.WriteLine(pwdProfileBox.Text)
            sw.Flush()
            sw.Close()
        Else
            MsgBox("File Does Not Exist")
        End If


        'Updates profile box lists
        pwdProBox.Items.Clear()
        Dim srFileReader As StreamReader
        srFileReader = System.IO.File.OpenText("C:\SPG\passwordProfiles.txt")
        I = 0
        Do Until srFileReader.ReadLine() Is Nothing
            pwdProBox.Items.Insert(I, srFileReader.ReadLine())
            I = I + 1
        Loop
        srFileReader.Close()
        'Finished updating password profile combobox list



        'Create pass profile configuration file
        If File.Exists("C:\SPG\" + pwdProfileBox.Text + ".txt") Then
            MessageBox.Show("Password Profile exists, please select another name")
        Else
            Dim oWrite As System.IO.StreamWriter
            oWrite = File.CreateText("C:\SPG\" + pwdProfileBox.Text + ".txt")
            oWrite.WriteLine(CheckedListBox1.GetItemChecked(3))
            oWrite.WriteLine(CheckedListBox1.GetItemChecked(2))
            oWrite.WriteLine(CheckedListBox1.GetItemChecked(1))
            oWrite.WriteLine(CheckedListBox1.GetItemChecked(0))
            oWrite.WriteLine(ComboBox1.SelectedIndex())
            oWrite.WriteLine(ComboBox2.SelectedIndex())
            oWrite.WriteLine(ComboBox3.SelectedIndex())
            oWrite.Close()
        End If
        'END of writing password profile settings



        MessageBox.Show("Password profile succesfully created!")

        'Reset settings to default
        pwdProfileBox.Text = ""
        CheckedListBox1.SetItemChecked(0, False)
        CheckedListBox1.SetItemChecked(1, False)
        CheckedListBox1.SetItemChecked(2, False)
        CheckedListBox1.SetItemChecked(3, False)
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""

        'Disable all fields
        pwdProBox.Enabled = True
        CheckedListBox1.Enabled = False
        GeneratePWDnewProfileButton.Enabled = False
        txtBoxResult.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False


    End Sub

    Private Sub btnGeneratePwd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GeneratePwdPassProfileButton.Click

        Dim pwdProfile As String
        Dim password As String
        Dim strongPWD As String

        'Password variable holds concatenation of text boxes
        password = txtbox1.Text & txtbox2.Text & txtbox3.Text

        'start string to byte encoding conversion
        'encodes password from string to byte type and stores in bytePassword variable
        Dim encText As New System.Text.UTF8Encoding()
        Dim bytePassword() As Byte

        'password in byte form
        bytePassword = encText.GetBytes(password)
        'END of string to byte encoding conversion


        'Read salt value file
        Dim oFile As System.IO.FileInfo
        pwdProfile = ("C:\SPG\" + pwdProBox.SelectedItem() + ".dat")
        oFile = New System.IO.FileInfo(pwdProfile)
        Dim oFileStream2 As System.IO.FileStream = oFile.OpenRead()
        Dim lBytes As Long = oFileStream2.Length
        Dim saltBytes(lBytes - 1) As Byte
        If (lBytes > 0) Then
            ' Read the file into a byte array
            oFileStream2.Read(saltBytes, 0, lBytes)
            oFileStream2.Close()
        End If

        'CONCATENATE salt with hashed password

        ' Allocate array, which will hold plain text and salt.
        Dim passwordWithSaltBytes() As Byte = New Byte(bytePassword.Length + saltBytes.Length - 1) {}

        ' Copy plain text bytes into resulting array.
        Dim I As Integer
        For I = 0 To bytePassword.Length - 1
            passwordWithSaltBytes(I) = bytePassword(I)
        Next I

        ' Append salt bytes to the resulting array.
        For I = 0 To saltBytes.Length - 1
            passwordWithSaltBytes(bytePassword.Length + I) = saltBytes(I)
        Next I
        'End concatenation of salt and hashed password


        'START SHA 512 conversion
        Dim finalresult() As Byte
        Dim shaM As New SHA512Managed()

        'computes SHA-512 for bytePassword and stores in result variable of type byte
        finalresult = shaM.ComputeHash(passwordWithSaltBytes)
        'END of SHA-512 conversion

        'Convert byte array to base64 string and print to textbox
        strongPWD = Convert.ToBase64String(finalresult)
        'Trim password
        Dim MyChar() As Char = {"=", "="}
        strongPWD = strongPWD.TrimEnd(MyChar)



        ' Start of Checked List box code
        Dim illegalChars As Char() = ""
        Dim sb As New System.Text.StringBuilder

        'If all selected
        If CheckedListBoxMain.GetItemChecked(0) = True And CheckedListBoxMain.GetItemChecked(1) = True And CheckedListBoxMain.GetItemChecked(2) = True And CheckedListBoxMain.GetItemChecked(3) = True Then
            'Leave the same
        End If
        'End of all selected


        'If only alpha-numeric 
        If CheckedListBoxMain.GetItemChecked(0) = True And CheckedListBoxMain.GetItemChecked(1) = True And CheckedListBoxMain.GetItemChecked(2) = True And CheckedListBoxMain.GetItemChecked(3) = False Then
            illegalChars = "+/".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of alpha-numeric



        'If only alphabetic and symbols
        If CheckedListBoxMain.GetItemChecked(0) = True And CheckedListBoxMain.GetItemChecked(1) = True And CheckedListBoxMain.GetItemChecked(2) = False And CheckedListBoxMain.GetItemChecked(3) = True Then
            illegalChars = "0123456789".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of alphabetic and symbols




        'If only alphabetic uppercase
        If CheckedListBoxMain.GetItemChecked(0) = True And CheckedListBoxMain.GetItemChecked(1) = False And CheckedListBoxMain.GetItemChecked(2) = False And CheckedListBoxMain.GetItemChecked(3) = False Then
            strongPWD = strongPWD.ToUpper()
            illegalChars = "+/0123546789".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of alphabetic uppercase


        'If only alphabetic lowercase
        If CheckedListBoxMain.GetItemChecked(0) = False And CheckedListBoxMain.GetItemChecked(1) = True And CheckedListBoxMain.GetItemChecked(2) = False And CheckedListBoxMain.GetItemChecked(3) = False Then
            strongPWD = strongPWD.ToLower()
            illegalChars = "+/0123546789".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of alphabetic lowercase



        'If only alphabetic 
        If CheckedListBoxMain.GetItemChecked(0) = True And CheckedListBoxMain.GetItemChecked(1) = True And CheckedListBoxMain.GetItemChecked(2) = False And CheckedListBoxMain.GetItemChecked(3) = False Then
            illegalChars = "+/0123546789".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of alphabetic 



        'If only alphabetic uppercase and numbers
        If CheckedListBoxMain.GetItemChecked(0) = True And CheckedListBoxMain.GetItemChecked(1) = False And CheckedListBoxMain.GetItemChecked(2) = True And CheckedListBoxMain.GetItemChecked(3) = False Then
            strongPWD = strongPWD.ToUpper()
            illegalChars = "+/".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of only alphabetic uppercase and numbers



        'If only alphabetic uppercase, symbols and numbers
        If CheckedListBoxMain.GetItemChecked(0) = True And CheckedListBoxMain.GetItemChecked(1) = False And CheckedListBoxMain.GetItemChecked(2) = True And CheckedListBoxMain.GetItemChecked(3) = True Then
            strongPWD = strongPWD.ToUpper()
            illegalChars = "".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of only alphabetic uppercase, symbols and numbers



        'If only alphabetic lowercase, symbols and numbers
        If CheckedListBoxMain.GetItemChecked(0) = False And CheckedListBoxMain.GetItemChecked(1) = True And CheckedListBoxMain.GetItemChecked(2) = True And CheckedListBoxMain.GetItemChecked(3) = True Then
            strongPWD = strongPWD.ToLower()
            illegalChars = "".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of only alphabetic lowercase, symbols and numbers



        'If only alphabetic lowercase and numbers
        If CheckedListBoxMain.GetItemChecked(0) = False And CheckedListBoxMain.GetItemChecked(1) = True And CheckedListBoxMain.GetItemChecked(2) = True And CheckedListBoxMain.GetItemChecked(3) = False Then
            strongPWD = strongPWD.ToLower()
            illegalChars = "+/".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of alphabetic lowercase and numbers



        'If only numeric
        If CheckedListBoxMain.GetItemChecked(0) = False And CheckedListBoxMain.GetItemChecked(1) = False And CheckedListBoxMain.GetItemChecked(2) = True And CheckedListBoxMain.GetItemChecked(3) = False Then
            illegalChars = "+/abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of only numeric



        'If only symbols
        If CheckedListBoxMain.GetItemChecked(0) = False And CheckedListBoxMain.GetItemChecked(1) = False And CheckedListBoxMain.GetItemChecked(2) = False And CheckedListBoxMain.GetItemChecked(3) = True Then
            illegalChars = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of only symbols


        'If only symbols and numbers
        If CheckedListBoxMain.GetItemChecked(0) = False And CheckedListBoxMain.GetItemChecked(1) = False And CheckedListBoxMain.GetItemChecked(2) = True And CheckedListBoxMain.GetItemChecked(3) = True Then
            illegalChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of symbols and numbers



        'If none
        If CheckedListBoxMain.GetItemChecked(0) = False And CheckedListBoxMain.GetItemChecked(1) = False And CheckedListBoxMain.GetItemChecked(2) = False And CheckedListBoxMain.GetItemChecked(3) = False Then
            strongPWD = ""
        End If
        'End of none



        'If only symbols and uppercase
        If CheckedListBoxMain.GetItemChecked(0) = True And CheckedListBoxMain.GetItemChecked(1) = False And CheckedListBoxMain.GetItemChecked(2) = False And CheckedListBoxMain.GetItemChecked(3) = True Then
            strongPWD = strongPWD.ToUpper()
            illegalChars = "0123456789".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of symbols and uppercase



        'If only symbols and lowercase
        If CheckedListBoxMain.GetItemChecked(0) = False And CheckedListBoxMain.GetItemChecked(1) = True And CheckedListBoxMain.GetItemChecked(2) = False And CheckedListBoxMain.GetItemChecked(3) = True Then
            strongPWD = strongPWD.ToLower()
            illegalChars = "0123456789".ToCharArray()
            For Each ch As Char In strongPWD
                If Array.IndexOf(illegalChars, ch) = -1 Then
                    sb.Append(ch)
                End If
            Next
            strongPWD = sb.ToString()
        End If
        'End of symbols and lowercase

        'END of ' Start of Checked List box code modifications to password


        'Write password to textbox
        txtboxPwdGenerated.Text = strongPWD
        'End of password generating code


        MessageBox.Show("Password profile succesfully created!")


        'Reset settings to default
        pwdProfileBox.Text = ""
        CheckedListBoxMain.SetItemChecked(0, False)
        CheckedListBoxMain.SetItemChecked(1, False)
        CheckedListBoxMain.SetItemChecked(2, False)
        CheckedListBoxMain.SetItemChecked(3, False)
        txtbox1.Text = ""
        txtbox2.Text = ""
        txtbox3.Text = ""

        'Disable all fields
        CheckedListBoxMain.Enabled = False
        GeneratePwdPassProfileButton.Enabled = False
        txtboxPwdGenerated.Enabled = False
        ComboBoxPID1.Enabled = False
        ComboBoxPID2.Enabled = False
        ComboBoxPID3.Enabled = False
        txtbox1.Enabled = False
        txtbox2.Enabled = False
        txtbox3.Enabled = False

    End Sub


    Private Sub btnCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyPasswdFromPassProfileButton.Click
        Clipboard.SetText(txtboxPwdGenerated.Text)
        'Reset password to blank to default
        txtboxPwdGenerated.Text = ""

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyPasswdFromNewProfileButton.Click
        Clipboard.SetText(txtBoxResult.Text)

        'Reset password to blank to default
        txtBoxResult.Text = ""

    End Sub

    Private Sub pwdProBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pwdProBox.SelectedIndexChanged

        'Reset settings to default
        CheckedListBoxMain.SetItemChecked(0, False)
        CheckedListBoxMain.SetItemChecked(1, False)
        CheckedListBoxMain.SetItemChecked(2, False)
        CheckedListBoxMain.SetItemChecked(3, False)
        txtbox1.Text = ""
        txtbox2.Text = ""
        txtbox3.Text = ""
        txtboxPwdGenerated.Text = ""



        'Read password profile file and sets settings
        If (Not System.IO.Directory.Exists("C:\SPG")) Then
            System.IO.Directory.CreateDirectory("C:\SPG")
        End If



        'Load password profile settings
        If System.IO.File.Exists("C:\SPG\" + pwdProBox.SelectedItem() + ".txt") = True Then
            txtbox1.Enabled = True
            txtbox2.Enabled = True
            txtbox3.Enabled = True
            GeneratePwdPassProfileButton.Enabled = True
            CopyPasswdFromPassProfileButton.Enabled = True

            Dim srFileReader As System.IO.StreamReader
            Dim sInputLine As String
            srFileReader = System.IO.File.OpenText("C:\SPG\" + pwdProBox.SelectedItem() + ".txt")
            sInputLine = srFileReader.ReadLine()
            Dim I As Integer = 0
            Do Until sInputLine Is Nothing
                If I = 0 Then
                    If "True" = sInputLine Then
                        CheckedListBoxMain.SetItemChecked(3, True)
                    End If

                ElseIf I = 1 Then
                    If "True" = sInputLine Then
                        CheckedListBoxMain.SetItemChecked(2, True)
                    End If

                ElseIf I = 2 Then
                    If sInputLine = "True" Then
                        CheckedListBoxMain.SetItemChecked(1, True)
                    End If

                ElseIf I = 3 Then
                    If sInputLine = "True" Then
                        CheckedListBoxMain.SetItemChecked(0, True)
                    End If

                ElseIf I = 4 Then
                    ComboBoxPID1.SelectedIndex() = Convert.ToInt32(sInputLine)

                ElseIf I = 5 Then
                    ComboBoxPID2.SelectedIndex() = Convert.ToInt32(sInputLine)

                ElseIf I = 6 Then
                    ComboBoxPID3.SelectedIndex() = Convert.ToInt32(sInputLine)
                End If

                sInputLine = srFileReader.ReadLine()
                I = I + 1
            Loop
            srFileReader.Close()
        End If


    End Sub

    Private Sub pwdProfileBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pwdProfileBox.TextChanged

        'Reset settings to default
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
        CheckedListBox1.SetItemChecked(0, False)
        CheckedListBox1.SetItemChecked(1, False)
        CheckedListBox1.SetItemChecked(2, False)
        CheckedListBox1.SetItemChecked(3, False)
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""

        'Enable all fields
        CheckedListBox1.Enabled = True
        GeneratePWDnewProfileButton.Enabled = True
        txtBoxResult.Enabled = True
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        ComboBox3.Enabled = True
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True


    End Sub

End Class
