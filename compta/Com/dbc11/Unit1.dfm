object Form1: TForm1
  Left = 444
  Top = 177
  Width = 248
  Height = 219
  Caption = 'Paradox to Access'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object cboBDETblNames: TComboBox
    Left = 8
    Top = 8
    Width = 185
    Height = 21
    ItemHeight = 13
    TabOrder = 0
    Text = 'cboBDETblNames'
  end
  object Memo1: TMemo
    Left = 8
    Top = 32
    Width = 185
    Height = 89
    Lines.Strings = (
      'Memo1')
    TabOrder = 1
  end
  object Button1: TButton
    Left = 8
    Top = 128
    Width = 185
    Height = 25
    Caption = 'Construct Create command'
    TabOrder = 2
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 8
    Top = 160
    Width = 185
    Height = 25
    Caption = 'Create Table and copy data'
    TabOrder = 3
    OnClick = Button2Click
  end
  object ADOConnection1: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=C:\de' +
      'vpgi\CISX\PGZIMPACCESS.mdb;Mode=Share Deny None;Extended Propert' +
      'ies="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";J' +
      'et OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:' +
      'Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet ' +
      'OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password' +
      '="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Dat' +
      'abase=False;Jet OLEDB:Don'#39't Copy Locale on Compact=False;Jet OLE' +
      'DB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False'
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 200
    Top = 8
  end
  object ADOTable: TADOTable
    Connection = ADOConnection1
    Left = 200
    Top = 72
  end
  object ADOCommand: TADOCommand
    Connection = ADOConnection1
    Parameters = <>
    Left = 200
    Top = 104
  end
  object BDETable: TTable
    DatabaseName = 'DBDEMOS'
    Left = 200
    Top = 40
  end
end
