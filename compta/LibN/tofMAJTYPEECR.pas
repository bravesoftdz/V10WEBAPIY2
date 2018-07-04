{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 08/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MAJTYPEECR ()
Mots clefs ... : TOF;MAJTYPEECR
*****************************************************************}
Unit tofMAJTypeEcr ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     dbtables,
     FE_Main,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF ; 

Type
  TOF_MAJTYPEECR = Class (TOF)
  private
    Sortie    : Boolean ;
    CExercic  : THValCOmboBox ;
    BStop     : TButton ;
    procedure BStopOnClick(Sender : TObject) ;
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

    procedure CLancheFiche_FTypeEcr ;

  const TMsg: array[01..3] of string 	= (
      {01}  '1;Mise à jour des types d''écritures ;Voulez-vous arrêter le traitement ?;Q;YN;N;N'
      {02} ,'1;Mise à jour des types d''écritures ;Traitement en cours sortie impossible;W;O;O;O'
      {03} ,''
      ) ;


Implementation

procedure CLancheFiche_FTypeEcr ;
begin
	AGLLanceFiche('CP','MAJTYPEECR','','','');
end ;

procedure TOF_MAJTYPEECR.BStopOnClick(Sender : TObject) ;
begin
if Not Sortie then if HShowMessage(TMsg[1],'','')=mrYes then Sortie:=TRUE ;
end ;


procedure TOF_MAJTYPEECR.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_MAJTYPEECR.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_MAJTYPEECR.OnUpdate ;
Var Q : TQuery;
  SQL : string ;
  i : integer ;
begin
  Inherited ;
  Sortie:=FALSE ;
  SQL := 'Update ecriture set e_qualifpiece= "N" where e_qualifpiece in (';
  SQL := SQL + 'Select e_qualifpiece from ecriture where not (e_qualifpiece in ("S","N","R","U","P","C")) ' ;
  if CExercic.itemindex<>0 then
    SQL := SQL + 'and e_exercice="' + CExercic.Values[CExercic.ItemIndex]+'"';
  SQL := SQL + ')';
  i:=ExecuteSQL(Sql);
  if i>0 then sortie:=true;
end ;

procedure TOF_MAJTYPEECR.OnLoad ;
begin
  inherited;
CExercic.ItemIndex:=0 ;
if (BStop<>nil) and (not Assigned(BStop.OnClick)) then BStop.OnClick:=BStopOnClick ;
Sortie:=TRUE ;

end;

procedure TOF_MAJTYPEECR.OnArgument (S : String ) ;
begin
  Inherited ;
  CExercic:=THValComboBox(GetControl('CEXO'));
  BStop:=TButton(GetControl('BSTOP'));

end ;

procedure TOF_MAJTYPEECR.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_MAJTYPEECR.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_MAJTYPEECR.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_MAJTYPEECR ] ) ; 
end.

