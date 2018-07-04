{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 08/02/2001
Modifié le ... : 29/08/2001
Description .. : Unit qui lance la bacule ou l'annulation à partir de la fiche
Suite ........ : 
Suite ........ : Met à jour la date et  fait les controles necessaires
Mots clefs ... : PAIE;PGEURO
*****************************************************************}
{
PT-1 : 01/08/2001 : V547 : PH
       Rajout message fin de traitement de la bascule et du sens de traitement Euro eou franc
       pour faire la conversion en Euro ou un retour à l'état initial
PT-2 : 29/08/01 : V547 : PH
       Annulation de la bascule
}
Unit UTOFPG_BASCULEEURO ;

Interface

Uses windows,StdCtrls, Controls, Classes,  forms, sysutils, ComCtrls,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     HCtrls,HEnt1,HMsgBox,UTOF,UTOB,Vierge,HPanel,HTB97,EuroPaiePGI,EntPaie ;

Type
  TOF_PG_BASCULEEURO = Class (TOF)
    private
    LePanel : THPanel;
    TypBas  : String; // Bascule EURO ou Annulation de la bascule
    procedure LanceBascule (Sender: TObject);
    procedure LanceAnnul (Sender: TObject);
    public
    procedure OnArgument (S : String ) ; override ;
  end ;

Implementation

procedure TOF_PG_BASCULEEURO.LanceBascule(Sender: TObject);
var    SystemTime0 : TSystemTime;
       LaDate      : TDateTime;
       Q           : TQuery;
       St          : String;
       Okok        : Integer;
begin
GetLocalTime(SystemTime0);
LaDate := SystemTimeToDateTime (SystemTime0);
LaDate := DEBUTDEMOIS (LaDate);
OkOk := mrYes;
St := 'SELECT PPU_SALARIE FROM PAIEENCOURS WHERE PPU_DATEDEBUT>="'+UsDateTime(LaDate)+'"';
Q:=OpenSQL(St,TRUE) ;
if  Not Q.EOF then
    BEGIN
    st := 'Vous avez fait des paies depuis le '+DateToStr (LaDate)+' ? #13#10 Vous avez déjà fait des paies sur ce mois '+
          '#13#10 Voulez vous continuer ?';
    OkOk := PGIAsk (st,Ecran.Caption);
    END ;
Ferme (Q);
if (NOT VH_PAIE.PGTenueEuro) then
   begin
   if OkOk = mrYes then
    begin
// PT-1 : 01/08/2001 : V547  Rajout message et du sens du traitement Euro ou Franc
    BasculePaieEuro ('', LePanel, 'E');
    PGIBox ('Fin de traitement', Ecran.Caption);
    end;
   end
   else PGIBox ('Votre dosssier est déjà tenu en EURO ?', Ecran.Caption);
end;

procedure TOF_PG_BASCULEEURO.LanceAnnul (Sender: TObject);
var    St          : String;
       Okok        : Integer;
begin
st := 'Vous avez converti votre paie en EURO, et vous voulez annuler la convertion !'+
          '#13#10 Voulez vous continuer ?';
OkOk := PGIAsk (st,Ecran.Caption);
if OkOk = mrYes then
    begin
    BasculePaieEuro ('', LePanel, 'F');
    PGIBox ('Fin de traitement', Ecran.Caption);
    end;
end;

procedure TOF_PG_BASCULEEURO.OnArgument (S : String ) ;
var       BtnVal  : TToolbarButton97;
begin
  Inherited ;
TypBas := READTOKENST (S);
if  TypBas = 'F' then Ecran.Caption := 'Annulation basculement en EURO'
    else Ecran.Caption := 'Bascule de la paie en EURO';
UpdateCaption(TFVierge(Ecran));
BtnVal := TToolbarButton97 (GetControl ('BValider'));
LePanel := THPanel (GetControl ('PANELBASCULE'));
if (BtnVal <> NIL) AND (LePanel <> NIL) then
 begin
 if TypBas = 'E' Then BtnVal.OnClick := LanceBascule;
 if TypBas = 'F' Then BtnVal.OnClick := LanceAnnul;
 end;
end ;


Initialization
  registerclasses ( [ TOF_PG_BASCULEEURO ] ) ;
end.
