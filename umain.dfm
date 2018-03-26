object Form1: TForm1
  Left = 192
  Top = 118
  Width = 1254
  Height = 850
  Caption = 'Excel2MySQL'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 1246
    Height = 289
    Align = alTop
    TabOrder = 0
    object GroupBox1: TGroupBox
      Left = 8
      Top = 8
      Width = 409
      Height = 105
      Caption = 'Excel File'
      TabOrder = 0
      object Label1: TLabel
        Left = 16
        Top = 48
        Width = 103
        Height = 13
        Caption = 'Header Row Number:'
      end
      object edtExcelFile: TEdit
        Left = 16
        Top = 24
        Width = 353
        Height = 21
        TabOrder = 0
      end
      object btnSelectExcelFile: TButton
        Left = 368
        Top = 23
        Width = 25
        Height = 21
        Caption = '...'
        TabOrder = 1
        OnClick = btnSelectExcelFileClick
      end
      object edtHeaderRowNumber: TEdit
        Left = 16
        Top = 64
        Width = 89
        Height = 21
        TabOrder = 2
        Text = '1'
      end
    end
    object GroupBox2: TGroupBox
      Left = 424
      Top = 8
      Width = 409
      Height = 105
      Caption = 'MySQL Options'
      TabOrder = 1
      object Label2: TLabel
        Left = 40
        Top = 24
        Width = 25
        Height = 13
        Caption = 'Host:'
      end
      object Label3: TLabel
        Left = 43
        Top = 48
        Width = 22
        Height = 13
        Caption = 'Port:'
      end
      object Label4: TLabel
        Left = 216
        Top = 24
        Width = 51
        Height = 13
        Caption = 'Username:'
      end
      object Label5: TLabel
        Left = 218
        Top = 48
        Width = 49
        Height = 13
        Caption = 'Password:'
      end
      object Label6: TLabel
        Left = 206
        Top = 72
        Width = 61
        Height = 13
        Caption = 'Table Name:'
      end
      object Label7: TLabel
        Left = 16
        Top = 72
        Width = 49
        Height = 13
        Caption = 'Database:'
      end
      object edtMySQLHost: TEdit
        Left = 72
        Top = 24
        Width = 121
        Height = 21
        TabOrder = 0
        Text = 'localhost'
      end
      object edtMySQLPort: TEdit
        Left = 72
        Top = 48
        Width = 121
        Height = 21
        TabOrder = 1
        Text = '3306'
      end
      object edtMySQLUsername: TEdit
        Left = 272
        Top = 24
        Width = 121
        Height = 21
        TabOrder = 3
        Text = 'root'
      end
      object edtMySQLPassword: TEdit
        Left = 272
        Top = 48
        Width = 121
        Height = 21
        PasswordChar = '*'
        TabOrder = 4
      end
      object edtMySQLDatabase: TEdit
        Left = 72
        Top = 72
        Width = 121
        Height = 21
        TabOrder = 2
      end
      object edtMySQLTableName: TEdit
        Left = 272
        Top = 72
        Width = 121
        Height = 21
        TabOrder = 5
      end
    end
    object btnReadHeaders: TButton
      Left = 333
      Top = 120
      Width = 83
      Height = 25
      Caption = 'Read Headers'
      TabOrder = 2
      OnClick = btnReadHeadersClick
    end
    object btnReadTableHeaders: TButton
      Left = 725
      Top = 120
      Width = 107
      Height = 25
      Caption = 'Show Table Fields'
      TabOrder = 3
      OnClick = btnReadTableHeadersClick
    end
    object GroupBox6: TGroupBox
      Left = 8
      Top = 160
      Width = 825
      Height = 113
      Caption = 'Relation'
      TabOrder = 4
      object Label8: TLabel
        Left = 50
        Top = 24
        Width = 96
        Height = 13
        Caption = 'Master Table Name:'
      end
      object Label9: TLabel
        Left = 55
        Top = 48
        Width = 91
        Height = 13
        Caption = 'Master Field Name:'
      end
      object Label10: TLabel
        Left = 306
        Top = 24
        Width = 86
        Height = 13
        Caption = 'Slave Field Name:'
      end
      object Label12: TLabel
        Left = 16
        Top = 72
        Width = 130
        Height = 13
        Caption = 'Master Lookup Field Name:'
      end
      object edtRelationMasterTableName: TEdit
        Left = 154
        Top = 24
        Width = 121
        Height = 21
        TabOrder = 0
      end
      object edtRelationMasterFieldName: TEdit
        Left = 154
        Top = 48
        Width = 121
        Height = 21
        TabOrder = 1
      end
      object edtRelationSlaveFieldName: TEdit
        Left = 402
        Top = 24
        Width = 121
        Height = 21
        TabOrder = 2
      end
      object edtRelationMasterLookupFieldName: TEdit
        Left = 154
        Top = 72
        Width = 121
        Height = 21
        TabOrder = 3
      end
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 289
    Width = 1246
    Height = 212
    Align = alTop
    TabOrder = 1
    object GroupBox3: TGroupBox
      Left = 8
      Top = 8
      Width = 409
      Height = 161
      Caption = 'Excel Headers'
      TabOrder = 0
      object lbExcelHeaders: TListBox
        Left = 16
        Top = 24
        Width = 377
        Height = 121
        ItemHeight = 13
        TabOrder = 0
      end
    end
    object GroupBox4: TGroupBox
      Left = 424
      Top = 8
      Width = 409
      Height = 161
      Caption = 'MySQL Table Fields'
      TabOrder = 1
      object lbMySQLFields: TListBox
        Left = 16
        Top = 24
        Width = 377
        Height = 121
        ItemHeight = 13
        TabOrder = 0
      end
    end
    object GroupBox5: TGroupBox
      Left = 840
      Top = 8
      Width = 393
      Height = 161
      Caption = 'Matchings'
      TabOrder = 2
      object lbMatchings: TListBox
        Left = 16
        Top = 24
        Width = 361
        Height = 121
        ItemHeight = 13
        TabOrder = 0
      end
    end
    object btnAddMatching: TButton
      Left = 725
      Top = 176
      Width = 107
      Height = 25
      Caption = 'Add Matching'
      TabOrder = 3
      OnClick = btnAddMatchingClick
    end
    object btnRemoveMatching: TButton
      Left = 1117
      Top = 176
      Width = 115
      Height = 25
      Caption = 'Remove Matching'
      TabOrder = 4
      OnClick = btnRemoveMatchingClick
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 501
    Width = 1246
    Height = 46
    Align = alTop
    TabOrder = 2
    object btnStartTransfer: TButton
      Left = 112
      Top = 8
      Width = 97
      Height = 25
      Caption = 'Start Transfer'
      TabOrder = 0
      OnClick = btnStartTransferClick
    end
    object cbUnique: TCheckBox
      Left = 8
      Top = 12
      Width = 89
      Height = 17
      Caption = 'Add Unique'
      TabOrder = 1
    end
  end
  object Panel4: TPanel
    Left = 0
    Top = 547
    Width = 1246
    Height = 272
    Align = alClient
    TabOrder = 3
    object memLogs: TMemo
      Left = 1
      Top = 1
      Width = 1244
      Height = 270
      Align = alClient
      Font.Charset = TURKISH_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Consolas'
      Font.Style = []
      ParentFont = False
      ScrollBars = ssBoth
      TabOrder = 0
    end
  end
  object odSelectExcelFile: TOpenDialog
    DefaultExt = '*.xlsx'
    Filter = 
      'Excel Files (*.xlsx)|*.xlsx|Excel Files (*.xls)|*.xls|All Files ' +
      '(*.*)|*.*'
    Left = 72
    Top = 120
  end
  object MySQLConnection: TZConnection
    Protocol = 'mysql'
    Left = 8
    Top = 120
  end
  object MySQLQuery: TZQuery
    Connection = MySQLConnection
    Params = <>
    Left = 40
    Top = 120
  end
  object MySQLQuery2: TZQuery
    Connection = MySQLConnection
    Params = <>
    Left = 120
    Top = 128
  end
end
