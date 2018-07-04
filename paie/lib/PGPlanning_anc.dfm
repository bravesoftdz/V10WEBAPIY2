object Planning: TPlanning
  Left = 44
  Top = 164
  Width = 604
  Height = 449
  HelpContext = 42255005
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Planning des absences'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PPanel: TPanel
    Left = 0
    Top = 0
    Width = 596
    Height = 415
    HelpContext = 42255005
    Align = alClient
    Caption = 'PPanel'
    TabOrder = 0
    object PRECAP: TPanel
      Left = 1
      Top = 335
      Width = 594
      Height = 79
      Align = alBottom
      TabOrder = 0
      Visible = False
      object GBAbsence: TGroupBox
        Left = 2
        Top = 4
        Width = 212
        Height = 73
        Anchors = [akLeft, akTop, akBottom]
        Caption = 'Absence en cours'
        TabOrder = 0
      end
      object GBRecap: TGroupBox
        Left = 213
        Top = 4
        Width = 381
        Height = 73
        Anchors = [akLeft, akTop, akRight, akBottom]
        Caption = 'Récapitulatif absence'
        TabOrder = 1
      end
    end
  end
end
