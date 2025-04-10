object FormSetting: TFormSetting
  Left = 0
  Top = 0
  Caption = 'FormSetting'
  ClientHeight = 179
  ClientWidth = 328
  Color = 16635850
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  OnCreate = FormCreate
  TextHeight = 15
  object lblDebugSQL: TLabel
    Left = 24
    Top = 24
    Width = 59
    Height = 15
    Caption = 'Debug SQL'
  end
  object lblUserRESTAPI: TLabel
    Left = 27
    Top = 56
    Width = 69
    Height = 15
    Caption = 'Use REST API'
  end
  object lblUser: TLabel
    Left = 200
    Top = 56
    Width = 23
    Height = 15
    Caption = 'User'
  end
  object lblPass: TLabel
    Left = 200
    Top = 88
    Width = 23
    Height = 15
    Caption = 'Pass'
  end
  object tgsDebugSQL: TToggleSwitch
    Left = 104
    Top = 19
    Width = 73
    Height = 20
    TabOrder = 0
  end
  object tgsUseRESTAPI: TToggleSwitch
    Left = 104
    Top = 56
    Width = 73
    Height = 20
    TabOrder = 1
    OnClick = tgsUseRESTAPIClick
  end
  object btnSave: TButton
    Left = 136
    Top = 136
    Width = 75
    Height = 25
    Caption = 'Save'
    TabOrder = 2
    OnClick = btnSaveClick
  end
  object btnCancel: TButton
    Left = 232
    Top = 136
    Width = 75
    Height = 25
    Caption = 'Cancel'
    TabOrder = 3
    OnClick = btnCancelClick
  end
  object edtUser: TEdit
    Left = 242
    Top = 53
    Width = 78
    Height = 23
    TabOrder = 4
  end
  object edtPass: TEdit
    Left = 242
    Top = 82
    Width = 78
    Height = 23
    PasswordChar = '*'
    TabOrder = 5
  end
end
