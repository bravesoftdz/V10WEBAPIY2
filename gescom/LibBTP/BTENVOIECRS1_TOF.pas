{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 19/02/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTENVOIECRS1 ()
Mots clefs ... : TOF;BTENVOIECRS1
*****************************************************************}
Unit BTENVOIECRS1_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     uTob,
{$ENDIF}
     forms,
     ParamSoc,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     Ent1,
     HMsgBox,
     HTB97,
     UtilActionComSx,
     UTOF ;

Type
  TOF_BTENVOIECRS1 = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	UnEchangeComSx : TEchangeComSx;
    CBENVOYE : TCheckBox;
  	BLANCE : TToolBarButton97;
    REPERTDESTINATION : THCritMaskEdit;
    DATETRF,DATETRF1 : THedit;
    procedure BLanceClick (Sender : TObject);
  end ;

Implementation

procedure TOF_BTENVOIECRS1.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTENVOIECRS1.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTENVOIECRS1.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTENVOIECRS1.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTENVOIECRS1.OnArgument (S : String ) ;
var Frequence : string;
    MM,JJ,AA  : word;
    DateDep,DateFin : TdateTime;
begin
  Inherited ;
  BLANCE := TToolBarButton97 (GetControl('BLANCE'));
  Blance.OnClick := BLanceClick;
  REPERTDESTINATION := THCritMaskEdit(GetControl('REPERTDESTINATION'));
  REPERTDESTINATION.Text := GetParamSocSecur ('SO_CPRDREPERTOIRE','');
  DATETRF := THedit(GetControl('DATETRF'));
  DATETRF1 := THedit(GetControl('DATETRF1'));
  //
  Frequence := GetParamSocSecur ('SO_CPFREQUENCESX','HEB');
  if Frequence = 'HEB' Then
  begin
    DecodeDate (V_PGI.dateEntree,AA,MM,JJ);
    DateDep := PremierJourSemaine (NumSemaine(V_PGI.DateEntree),AA);
    DateFin := PlusDate(Datedep,7,'J');
    DATETRF.Text := DateToStr (DateDep);
    DATETRF.Text := DateToStr (DateFin);
  end
  else if Pos(Frequence ,'JOU;PON;')>0then
  begin
    DATETRF.Text := DateToStr (V_PGI.dateEntree);
    DATETRF1.Text := DateToStr (V_PGI.dateEntree);
  end
  else if Frequence ='MEN' then
  begin
    DATETRF.Text := DateToStr(DEBUTDEMOIS (V_PGI.DateEntree));
    DATETRF1.Text := DateToStr(FINDEMOIS(V_PGI.DateEntree));
  end;
  //
  CBENVOYE := TCheckBox (getControl('CBENVOYE'));
  UnEchangeComSx := TEchangeComSx.create;
end ;

procedure TOF_BTENVOIECRS1.OnClose ;
begin
  Inherited ;
  UnEchangeComSx.free;
end ;

procedure TOF_BTENVOIECRS1.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTENVOIECRS1.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTENVOIECRS1.BLanceClick(Sender: TObject);
var Fichier : string;
    EtatEnvoye : string;
begin
//
  if not IsValidDate (DATETRF.text) then
  begin
  	TForm(Ecran).ModalResult:=0;
    exit;
  end;
  If PgiAsk(TraduireMemoire ('Validez-vous le traitement d''envoi des écritures ?'),ecran.caption) = MrYes then
  begin
  	UnEchangeComSx.Init;
    Fichier := REPERTDESTINATION.text +'\'+'PG'+ V_PGI.NoDossier+'.TRA';
    if CBENVOYE.State = cbchecked then EtatEnvoye := 'X'
    else if CBENVOYE.State = cbUnchecked then EtatEnvoye := '-'
    else EtatEnvoye := '';
  	UnEchangeComSx.ExecuteAction (TmaSxExport,Fichier,DATETRF.Text,DATETRF.text,DATETRF1.text,EtatEnvoye);
    PgiBox('Traitement Terminé',ecran.caption);
  end;
end;

Initialization
  registerclasses ( [ TOF_BTENVOIECRS1 ] ) ;
end.

