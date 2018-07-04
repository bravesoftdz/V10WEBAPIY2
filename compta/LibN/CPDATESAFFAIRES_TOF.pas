{***********UNITE*************************************************
Auteur  ...... : YMO
Créé le ...... : 22/02/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPDATESAFFAIRES ()
Mots clefs ... : TOF;CPDATESAFFAIRES
*****************************************************************}
Unit CPDATESAFFAIRES_TOF ;

Interface

Uses Windows,
     forms,
     sysutils,
     ComCtrls,
     ExtCtrls,  // pout le TPanel
     Classes,
     StdCtrls,
     Controls,
     HTB97,     // pour TToolBarButton97
     Grids,     // Pour le TGridDrawState
     Graphics,  // pour clRed
{$IFDEF EAGLCLIENT}
    eMul,
    maineagl,
{$ELSE}
    FE_MAIN,
    Mul,
    db,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    SaisBor,       // Pour lanceSaisieFolio
    HDB,
    dbGrids,
{$ENDIF}
     HQry,
     HCtrls,
     HEnt1,
     Ent1,
     HMsgBox,
     UTOF,
     UTOB,
     UtilPGI,       // pour Resolution
     Vierge,
     ImgList;

function ModifZoneSerie: string;

Type
  TOF_CPDATESAFFAIRES = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure BValiderClick( sender: TObject ) ;
    procedure OnClickAffaire(Sender: TObject);
   private
    DebChantier : THEdit;
    FinChantier : THEdit;
    AffEnCours : TCheckBox;    
  end ;


Implementation

procedure TOF_CPDATESAFFAIRES.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPDATESAFFAIRES.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPDATESAFFAIRES.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPDATESAFFAIRES.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPDATESAFFAIRES.OnArgument (S : String ) ;
begin
  Inherited ;

  Ecran.Caption:='Mise à jour des dates';
  UpdateCaption(Ecran);

  AffEnCours    := TCheckBox(GetControl('AFFAIREENCOURS', True));
  DebChantier   := THEdit(GetControl('DATEDEBCHANTIER', True));
  FinChantier   := THEdit(GetControl('DATEFINCHANTIER', True));
  TToolBarButton97(GetControl ('BValider')).OnClick := BValiderClick ;
  DebChantier.Text:=StDate1900;
  FinChantier.Text:=StDate2099;
  AffEnCours.Checked:=True;
  AffEnCours.OnClick:=OnClickAffaire;
end ;

procedure TOF_CPDATESAFFAIRES.OnClose ;
var IsChecked, Str : String;
begin
 //YMO 04/07/2006 FQ18258 On ne met à jour que les dates renseignées 
 If Not AFFENCOURS.Checked Then
 begin
    SetControlText('DATEDEBCHANTIER', DateTimeToStr(iDate1900));
    SetControlText('DATEFINCHANTIER', DateTimeToStr(iDate2099));
 end;

 If AffEnCours.Checked then
    IsChecked := 'X'
 else
    IsChecked := '-';

 Str:=' S_AFFAIREENCOURS = "'+IsChecked;

 If (GetControlText('DATEDEBCHANTIER') <> DateTimeToStr(iDate1900)) Or Not (AFFENCOURS.Checked) then
   Str:=Str+'",S_DEBCHANTIER="'+USDateTime(StrToDate(DebChantier.Text));

 If (GetControlText('DATEFINCHANTIER') <> DateTimeToStr(iDate2099)) Or Not (AFFENCOURS.Checked) then
   Str:=Str+'",S_FINCHANTIER="'+USDateTime(StrToDate(FinChantier.Text));

 Str:=Str+'"';

 TFVierge(Ecran).Retour:=Str;

 Inherited ;
end ;

procedure TOF_CPDATESAFFAIRES.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPDATESAFFAIRES.OnCancel () ;
begin
  Inherited ;
end ;

function ModifZoneSerie: string;
BEGIN
Result:=AGLLanceFiche('CP', 'CPDATESAFFAIRES', '', '', '') ;
END;

procedure TOF_CPDATESAFFAIRES.BValiderClick(sender: TObject);
begin

end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 23/03/2006
Modifié le ... :   /  /
Description .. : Affaire en cours : faire apparaître les dates
Mots clefs ... :
*****************************************************************}
procedure TOF_CPDATESAFFAIRES.OnClickAffaire(Sender: TObject);
begin
    SetControlEnabled('DATEDEBCHANTIER',AFFENCOURS.Checked);
    SetControlEnabled('DATEFINCHANTIER',AFFENCOURS.Checked);
end;

Initialization
  registerclasses ( [ TOF_CPDATESAFFAIRES ] ) ;
end.

