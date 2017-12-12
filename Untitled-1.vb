Đoạn mã 1

'Hàm ChangeProperty thay đổi các thuộc tính của CSDL

Function ChangeProperty(strPropName, varPropType, varPropValue)

  Dim dbs As Database, prp As Property

  Const conPropNotFoundError = 3270

  Set dbs = CurrentDb

  On Error GoTo Change_XuLyLoi

  dbs.Properties(strPropName) = varPropValue

  ChangeProperty = True

Change_KetThuc:

  Exit Function

Change_XuLyLoi:

  'Thuộc tính không thấy

  If Err = conPropNotFoundError Then 

  Set prp = dbs.CreateProperty(strPropName, _

  varPropType, varPropValue)

  dbs.Properties.Append prp

  Resume Next

  Else

  'Không biết lỗi gì

  ChangeProperty = False

  Resume Change_KetThuc

  End If

End Function

'Xử lý tình huống chọn nút [Khóa database]

Private Sub cmdLock_Click()

  ‘Biểu mẫu này được nạp trước

  ChangeProperty "StartupForm", dbText, "frmKhoiDong"

  ChangeProperty "StartupShowDBWindow", dbBoolean, False

  ChangeProperty "StartupShowStatusBar", dbBoolean, False

  ChangeProperty "AllowBuiltinToolbars", dbBoolean, False

  ChangeProperty "AllowFullMenus", dbBoolean, False

  ChangeProperty "AllowBreakIntoCode", dbBoolean, False

  ChangeProperty "AllowSpecialKeys", dbBoolean, False

 

  ‘Không cho xài phím Shift để bỏ qua biểu mẫu frmKhoiDong

  ChangeProperty "AllowBypassKey", dbBoolean, False

 

  MsgBox "Cơ sở dữ liệu đã được khóa! Đóng cơ sở dữ liệu, _

  rồi mở lại mới có ép-phê.", vbOKOnly, "eChip Security"

  cmdExit.SetFocus

  cmdUnlock.Visible = True

  cmdLock.Visible = False

End Sub

'Xử lý tình huống chọn nút [Mở database]

Private Sub cmdUnlock_Click()

  ‘Không cần biểu mẫu khởi động nữa

  ChangeProperty "StartupForm", dbText, ""

  ChangeProperty "StartupShowDBWindow", dbBoolean, True

  ChangeProperty "StartupShowStatusBar", dbBoolean, True

  ChangeProperty "AllowBuiltinToolbars", dbBoolean, True

  ChangeProperty "AllowFullMenus", dbBoolean, True

  ChangeProperty "AllowBreakIntoCode", dbBoolean, True

  ChangeProperty "AllowSpecialKeys", dbBoolean, True

  ChangeProperty "AllowBypassKey", dbBoolean, True

  MsgBox "Cơ sở dữ liệu đã được mở khóa ! _

  Đóng cơ sở dữ liệu, rồi mở lại mới có ép-phê.", _

  vbOKOnly, "eChip Security"

  cmdExit.SetFocus

  txtPassword = ""

  cmdLock.Visible = True

  cmdUnlock.Visible = False

  txtPassword.Visible = False

End Sub

'Xử lý tình huống khi mở biểu mẫu

Private Sub Form_Open(Cancel As Integer)

  Dim dbs As Database

  Set dbs = CurrentDb

  On Error GoTo KhongCoThuocTinh_Err

  If dbs.Properties("AllowBypassKey") Then

    cmdLock.Visible = True

    txtPassword.Visible = False

  Else

    cmdLock.Visible = False

    txtPassword.Visible = True

  End If

  Exit Sub

  KhongCoThuocTinh_Err:

  cmdLock.Visible = True

  txtPassword.Visible = False

End Sub

'Khi người ta gõ mật khẩu và nhấn phím Enter

Private Sub txtPassword_LostFocus()

  If txtPassword = "echip" Then

    cmdUnlock.Visible = True

  End If

End Sub
