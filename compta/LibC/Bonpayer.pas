{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 04/03/2003
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit BonPayer;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls,
{$IFNDEF EAGLCLIENT}
      EdtREtat,
{$ELSE}
      UtileAGL,
{$ENDIF}
  Ent1, hmsgbox, Buttons, ExtCtrls, HSysMenu,
  Choix, HEnt1, Hcompte, HTB97, HPanel, UiUtil ;

procedure LanceBonAPayer ;

type
  TFBonAPayer = class(TForm)
    HLabel1: THLabel;
    HLabel2: THLabel;
    HLabel3: THLabel;
    cModeleBAP: THValComboBox;
    E_SUIVDEC: THValComboBox;
    E_NOMLOT: TEdit;
    bListeLots: TToolbarButton97;
    HMTrad: THSystemMenu;
    Msg: THMsgBox;
    HLabel4: THLabel;
    E_GENERAL: THCpteEdit;
    BValider: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    Baide: TToolbarButton97;
    HPB: TToolWindow97;
    Dock: TDock97;
    Bevel1: TBevel;
    procedure BValiderClick(Sender: TObject);
    procedure bListeLotsClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BaideClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
  private
  public
  end;


implementation

{$R *.DFM}


procedure LanceBonAPayer ;
var X : TFBonAPayer ;
    PP : THPanel ;
begin
X:=TFBonAPayer.Create(Application) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     X.ShowModal ;
    finally
     X.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
end ;


procedure TFBonAPayer.BValiderClick(Sender: TObject);
var sSQL : string ;
begin
if ((E_GENERAL.Text='') or (E_GENERAL.ExisteH<=0)) then BEGIN Msg.Execute(4,caption,'') ; Exit ; END ; 
if E_SUIVDEC.Value='' then begin Msg.Execute(0,caption,'') ; exit ; end ;
if Trim(E_NOMLOT.Text)='' then begin Msg.Execute(1,caption,'') ; exit ; end ;
if cModeleBAP.Value='' then begin Msg.Execute(2,caption,'') ; exit ; end ;
SourisSablier ;
// Marquage des écritures
//RR 02/12/2002 sSQL:='UPDATE ECRITURE SET E_FLAGECR="'+V_PGI.User+'"' ;
//RR 02/12/2002 sSQL:=sSQL+' WHERE E_SUIVDEC="'+E_SUIVDEC.Value+'"' ;
sSQL:=sSQL+' E_SUIVDEC="'+E_SUIVDEC.Value+'"' ;
sSQL:=sSQL+' AND E_NOMLOT="'+E_NOMLOT.Text+'"' ;
sSQL:=sSQL+' AND E_GENERAL="'+E_GENERAL.Text+'"' ;
//RR 02/12/2002 ExecuteSQL(sSQL) ;
// Impression
//RR 02/12/2002 sSQL:='E_FLAGECR="'+V_PGI.User+'" AND E_GENERAL="'+E_GENERAL.Text+'"' ;
LanceEtat('E','BAP',cModeleBAP.Value,True,False,False,nil,sSQL,'',False) ;
//RR 02/12/2002 ExecuteSQL('UPDATE ECRITURE SET E_FLAGECR="" WHERE E_GENERAL="'+E_GENERAL.Text+'" AND E_FLAGECR="'+V_PGI.User+'"') ;
SourisNormale ;
end;

procedure TFBonAPayer.bListeLotsClick(Sender: TObject);
var sWhere : String ;
begin
if E_SUIVDEC.Value='' then begin Msg.Execute(0,caption,'') ; Exit ; end ;
sWhere:='E_GENERAL="'+E_GENERAL.Text+'" AND E_SUIVDEC="'+E_SUIVDEC.Value+'"' ;
E_NOMLOT.Text:=Choisir(Msg.Mess[3],'ECRITURE','DISTINCT E_NOMLOT','',sWhere,'E_NOMLOT') ;
end;

procedure TFBonAPayer.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
end;

procedure TFBonAPayer.FormClose(Sender: TObject; var Action: TCloseAction);
begin
//SG6 22.02.05
//if Parent is THPanel then Action:=caFree ;
  if IsInside(Self) then CloseInsidePanel(Self);
end;

procedure TFBonAPayer.BaideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

//SG6 22.02.05
procedure TFBonAPayer.BAnnulerClick(Sender: TObject);
begin
  Close;
end;

end.
