{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 15/04/2003
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit NumChek;

interface

uses
  Forms,
  Classes,
  Controls,
  StdCtrls,
  HCtrls,
  Buttons,
  ExtCtrls,
  Mask,
  SysUtils,  // StrToInt
  HMsgBox,   // THMsgBox
  HSysMenu   // THSystemMenu
  ;

function SaisieNumCheque(var sRef,snChq : string ; Renumerote : Boolean) : Boolean ;

type
  TFNumCheque = class(TForm)
    sTitre: THLabel;
    Panel1: TPanel;
    bValideCheque: THBitBtn;
    bAnnuleCheque: THBitBtn;
    BAide: THBitBtn;
    nCheque: TMaskEdit;
    eRef: TEdit;
    Msg: THMsgBox;
    HMTrad: THSystemMenu;
    procedure bValideChequeClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

function SaisieNumCheque(var sRef,snChq : string ; Renumerote : Boolean) : Boolean ;
var
  FNumCheque : TFNumCheque ;
begin
  result:=False ;
  FNumCheque := TFNumCheque.Create(Application) ;
  try
    if Renumerote then FNumCheque.sTitre.Caption:=FNumCheque.Msg.Mess[2]  // Saisissez le numéro du premier chèque sélectionné.
                  else FNumCheque.sTitre.Caption:=FNumCheque.Msg.Mess[1]; // Saisissez le numéro du premier chèque à imprimer.

    FNumCheque.nCheque.Text:=snChq ;
    if sRef<>'' then FNumCheque.eRef.Text:=sRef ;

    if (FNumCheque.ShowModal=mrOK) then begin
      snChq:=FNumCheque.nCheque.Text ;
      sRef:=FNumCheque.eRef.Text ;
      result:=True ;
    end ;
  finally
    FNumCheque.Free ;
  end ;
end ;

procedure TFNumCheque.bValideChequeClick(Sender: TObject);
begin
  if StrToInt(nCheque.Text)=0 then begin
    Msg.Execute(0,caption,'') ; // Le numéro de chèque doit être non nul.
    Exit ;
  end ;
  ModalResult:=mrOK ;
end;

procedure TFNumCheque.FormCreate(Sender: TObject);
begin
  PopupMenu:=AddMenuPop(PopupMenu,'','') ;
end;

procedure TFNumCheque.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self) ;
end;

end.
