{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 20/10/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : ESPCPQRMOD347 ()
Mots clefs ... : TOF;ESPCPQRMOD347
*****************************************************************}
Unit ESPCPQRMOD347_TOF ;

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
  TOF_ESPCPQRMOD347 = Class (TOF_Meth)
   private
    E_EXERCICE        : THValComboBox ;
    E_DATECOMPTABLE,
    E_DATECOMPTABLE_  : THedit ;
    DateD,DateF       : TDatetime ;
    procedure ExoOnChange(Sender : TObject);
    procedure DateOnExit(Sender : TObject) ;
    procedure T_NATUREAUXIExit(Sender : TOBject) ;
    Function GetNatureAuxi : THradioGroup ;

    Property T_NATUREAUXI : THRadioGroup read GetNatureAuxi ;

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

uses ULibExercice, Ent1,
{$IFDEF eAGLClient}
 eQRS1
{$ELSE}
 QRS1
{$ENDIF}
 ;

Procedure TOF_ESPCPQRMOD347.TRUC_create ;
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

  with THEdit(GetControl('MONTANTMINI')) do begin
     DisplayFormat:=strfmask(V_PGI.OkDecV,'',TRUE,TRUE) ;
     Text:='3000' ;
  End ;

  With THRadioGroup(getControl('T_NATUREAUXI')) do begin
     OnExit:=T_NATUREAUXIExit ;
  End ;
  T_NATUREAUXIExit(nil) ;

End ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRMOD347.OnNew ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRMOD347.OnDelete ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRMOD347.T_NATUREAUXIExit(Sender : TOBject) ;
var
  NatureJal    : String ;
begin
  NatureJal:='VTE' ;
  if T_NATUREAUXI.Values[T_NATUREAUxi.ItemIndex]='FOU' then
     NatureJal:='ACH' ;
  SetControlText('lblT_NATUREAUXI',T_NATUREAUXI.Items[T_NATUREAUxi.ItemIndex]) ;
  SetControltext('J_NATUREJAL',NatureJal) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////
Function TOF_ESPCPQRMOD347.GetNatureAuxi : THradioGroup ;
Begin
  Result:=THRadioGroup(getControl('T_NATUREAUXI')) ;
End ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRMOD347.OnUpdate ;
var
  SQL,GroupBy,Where,Calcul,NaturePieces : String ;
begin
  Inherited ;
  TFQRS1(Ecran).WhereSQL:=CMajRequeteExercice(E_EXERCICE.Value,TFQRS1(Ecran).WhereSQL);
  GroupBy:='T_SIRET, T_AUXILIAIRE,T_LIBELLE, T_ADRESSE1,  T_ADRESSE2, T_ADRESSE3, T_CODEPOSTAL, T_VILLE' ;
  if T_NATUREAUXI.Values[T_NATUREAUxi.ItemIndex]='CLI' then begin
     Calcul:='(sum(E_DEBITDEV)-SUM(E_CREDITDEV))' ;
     NaturePieces:='"FC","AC"' ;
  end
  Else begin
     Calcul:='(sum(E_CREDITDEV)-SUM(E_DEBITDEV))' ;
     NaturePieces:='"FF","AF"' ;
  End ;
  SQL:='Select '+Groupby+', '+Calcul+' as MONTANT '
      +       'from TIERS left outer join ECRITURE on E_AUXILIAIRE=T_AUXILIAIRE left outer join JOURNAL on J_JOURNAL=E_JOURNAL' ;
  Where:=TFQRS1(Ecran).WhereSQL ;
  if trim(Where)<>'' then
     Where:=where+' and ' ;
  where:=where+'E_NATUREPIECE in('+NaturePieces+') ' ;
  SQL:=SQL+' where '+Where
          +' group by '+GroupBy
          +' having '+Calcul+'>='+strfpoint(Valeur(getcontroltext('MONTANTMINI'))) ; //XMG 22/10/03

  TFQRS1(Ecran).WhereSQL:=SQL ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRMOD347.OnLoad ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRMOD347.OnArgument (S : String ) ;
begin
 Truc_Create ;
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRMOD347.OnClose ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRMOD347.OnDisplay () ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRMOD347.OnCancel () ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRMOD347.ExoOnChange(Sender: TObject);
Begin
  if E_EXERCICE.Value <> '' then //SDA le 20/03/2008
  begin
    if (Assigned(E_EXERCICE) and (Assigned(E_DATECOMPTABLE)) and (Assigned(E_DATECOMPTABLE_))) then
      CExoRelatifToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_);
    DateD:=StrToDate(E_DATECOMPTABLE.Text) ;
    DateF:=StrToDate(E_DATECOMPTABLE_.Text) ;
  end;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRMOD347.DateOnExit(Sender: TObject);
begin
  if E_EXERCICE.Value <> '' then //SDA le 20/03/2008
  begin
    DoDateOnExit(THEdit(Sender), E_DATECOMPTABLE, E_DATECOMPTABLE_, DateD, DateF);
  end;
end;
//////////////////////////////////////////////////////////////////////////////////
Initialization
  registerclasses ( [ TOF_ESPCPQRMOD347 ] ) ;
end.


