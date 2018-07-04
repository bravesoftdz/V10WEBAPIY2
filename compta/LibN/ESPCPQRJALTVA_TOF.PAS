{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 22/10/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : ESPCPQRJALTVA ()
Mots clefs ... : TOF;ESPCPQRJALTVA
*****************************************************************}
Unit ESPCPQRJALTVA_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     dbtables, 
     mul, 
{$else}
     eMul, 
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HMsgBox,
     UTOF,
     TofMeth ;

Type
  TOF_ESPCPQRJALTVA = Class (TOF_Meth)
   private
    E_EXERCICE        : THValComboBox ;
    E_DATECOMPTABLE,
    E_DATECOMPTABLE_  : THedit ;
    DateD,DateF       : TDatetime ;
    procedure ExoOnChange(Sender : TObject);
    procedure DateOnExit(Sender : TObject) ;
    procedure DoClickInitiales(Sender : TObject) ;
    procedure DoRenumeroter(Sender : TObject) ;
   public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    {Constructor} Procedure  TRUC_create ; //override ;
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  ULibExercice, Ent1, UtilDiv,
{$IFDEF eAGLClient}
    eQRS1
{$ELSE}
    QRS1
{$ENDIF}
    ;

Procedure TOF_ESPCPQRJALTVA.TRUC_create ;
Begin
  // inherited ;
  E_DATECOMPTABLE:=THEdit(GetControl('E_DATECOMPTABLE')) ; E_DATECOMPTABLE.OnExit:=DateOnExit ;

  E_DATECOMPTABLE_:=THEdit(GetControl('E_DATECOMPTABLE_')) ; E_DATECOMPTABLE_.OnExit:= DateOnExit ;

  E_EXERCICE:= THValComboBox(GetControl('E_EXERCICE')) ;
  CInitComboExercice(E_EXERCICE);
  E_EXERCICE.OnChange:=ExoOnChange ;
  if (CtxPCl in V_PGI.PgiContexte) and  (VH^.CPExoRef.Code<>'') then
     E_EXERCICE.Value:=CExerciceVersRelatif(VH^.CPExoRef.Code)
  else
     E_EXERCICE.Value:=CExerciceVersRelatif(VH^.Entree.Code) ;

  with THEdit(GetControl('BASEHTINI')) do
     DisplayFormat:=strfmask(V_PGI.OkDecV,'',TRUE,TRUE) ;

  with THEdit(GetControl('MONTANTTVAINI')) do
     DisplayFormat:=strfmask(V_PGI.OkDecV,'',TRUE,TRUE) ;

  with THEdit(GetControl('MONTANTTPFINI')) do
     DisplayFormat:=strfmask(V_PGI.OkDecV,'',TRUE,TRUE) ;

  with TCheckBox(GetControl('CBINITIALES')) do
     onClick:=DoClickinitiales ;
  DoClickInitiales(nil) ;

  with TCheckBox(GetControl('CBRenumeroter')) do
     onClick:=DoRenumeroter ;
End ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRJALTVA.OnNew ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRJALTVA.OnDelete ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRJALTVA.OnUpdate ;
var
   Exo : TExoDate ;
begin
  Exo.Code:=CRelatifVersExercice(GetControlText('E_EXERCICE')) ;
  RempliExoDate(Exo) ;
  SetControlText('FINEXERCICE',copy('-X',1+ord(Exo.Fin=strtodate(GetControlText('E_DATECOMPTABLE_'))),1)) ; //XMG 27/10/03
  Inherited ;
  TFQRS1(Ecran).WhereSQL:=CMajRequeteExercice(E_EXERCICE.Value,TFQRS1(Ecran).WhereSQL);
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRJALTVA.OnLoad ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRJALTVA.OnArgument (S : String ) ;
Var
  Titre,Etat,JAL : String ;
  OkVisible      : Boolean ;
begin
  Truc_Create ;
  Inherited ;
  Titre:='ventes' ;
  Etat:='JTV' ;
  JAL:='(J_NATUREJAL="VTE" or (J_NATUREJAL="ACH" and E_REGIMETVA="CEE"))' ;
  okVisible:=TRUE ;
  if TrouveArgument(S,'TYPEETAT=')='ACHAT' then begin
     Titre:='achats' ;
     Etat:='JTA' ;
     JAL:='J_NATUREJAL="ACH"' ;
     OkVisible:=FALSE ;
  end ;
  Ecran.Caption:=Ecran.Caption+' '+TraduireMemoire(Titre) ;
  updateCaption(Ecran) ;
  //SetControltext('J_NATUREJAL',JAL) ;
  SetControltext('XX_WHERE',JAL) ;
  SetControlVisible('MONTANTTPFINI',OkVisible) ;
  SetControlVisible('LBLMONTANTTPFINI',OkVisible) ;
  TFQRS1(Ecran).CodeEtat:=Etat ;
end ;
//////////////////////////////////////////////////////////////////////////////////

procedure TOF_ESPCPQRJALTVA.OnClose ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRJALTVA.OnDisplay () ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRJALTVA.OnCancel () ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRJALTVA.ExoOnChange(Sender: TObject);
Begin
  if (Assigned(E_EXERCICE) and (Assigned(E_DATECOMPTABLE)) and (Assigned(E_DATECOMPTABLE_))) then
    CExoRelatifToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_);
  DateD:=StrToDate(E_DATECOMPTABLE.Text) ;
  DateF:=StrToDate(E_DATECOMPTABLE_.Text) ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRJALTVA.DateOnExit(Sender: TObject);
begin
 DoDateOnExit(THEdit(Sender), E_DATECOMPTABLE, E_DATECOMPTABLE_, DateD, DateF);
end;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRJALTVA.DoClickInitiales(Sender : TObject) ;
Var
   ok : Boolean ;
Begin
  ok:=(GetControltext('CBInitiales')='X') ;
  SetControlEnabled('CBRenumeroter',ok) ;
  SetControlEnabled('NUMFACINI',ok and (GetControlText('CBRenumeroter')='X')) ;

  SetControlEnabled('NUMPAGEINI',ok) ;

  SetControlEnabled('BaseHtIni',ok) ;
  SetControlEnabled('MontantTVAIni',ok) ;
  SetControlEnabled('MontantTPFIni',ok) ;
End ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRJALTVA.DoRenumeroter(Sender : TObject) ;
Begin
  SetControlEnabled('NUMFACINI',(GetControlText('CBRenumeroter')='X')) ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Initialization
  registerclasses ( [ TOF_ESPCPQRJALTVA ] ) ;
end.
