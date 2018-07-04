{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 25/06/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTFOURNMULTI_MUL ()
Mots clefs ... : TOF;BTFOURNMULTI_MUL
*****************************************************************}
Unit BTFOURNMULTI_MUL_TOF ;

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
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     Ent1,
     HMsgBox,
     UtilGc,
     UTOB,
     UtilRt,
     UtilSelection,
     ParamSoc,
     AglInit,
     HTB97,
     UTOF ;

Type
  TOF_BTFOURNMULTI_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    F : TForm;
    xx_where,StFiltre : string ;
    TOBFOurnSel : TOB;
    procedure BouvrirClick (Sender : TObject);
  end ;

Implementation

procedure TOF_BTFOURNMULTI_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTFOURNMULTI_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTFOURNMULTI_MUL.OnUpdate ;
begin
end ;

procedure TOF_BTFOURNMULTI_MUL.OnLoad ;
begin
  Inherited ;
  SetControlText('XX_WHERE',xx_where) ;
end ;

procedure TOF_BTFOURNMULTI_MUL.OnArgument (S : String ) ;
var CC  : THValComboBox;
    ChampMul,ValMul,Critere : String;
    x : integer;
begin
  Inherited ;
  TOBFOurnSel := LaTOB;
  //
  F := TForm (Ecran);
  if GetParamSocSecur('SO_RTGESTINFOS003',False) = True then
      MulCreerPagesCL(F,'NOMFIC=GCFOURNISSEURS');
  {$ifdef GRC} // mcd 17/08/04 sinon plante si exe GI sans GRC
  xx_where:=RTXXWhereConfident('CONF');
  {$endif}
  GCMAJChampLibre (F, False, 'COMBO', 'YTC_TABLELIBREFOU', 3, '');
  GCMAJChampLibre (F, False, 'EDIT', 'YTC_VALLIBREFOU', 3, '');
  GCMAJChampLibre (F, False, 'EDIT', 'YTC_VALLIBREFOU', 3, '_');
  GCMAJChampLibre (F, False, 'EDIT', 'YTC_DATELIBREFOU', 3, '');
  GCMAJChampLibre (F, False, 'EDIT', 'YTC_DATELIBREFOU', 3, '_');

  // CCMX - DM - 21/09/2005 - DEBUT
  SetControlVisible('TYTC_TYPEFOURNI'  , GetParamSocSecur('SO_GCTRANSPORTEURS', ''));
  SetControlVisible('YTC_TYPEFOURNI'  , GetParamSocSecur('SO_GCTRANSPORTEURS', ''));
  // CCMX - DM - 21/09/2005 - FIN
  TToolBarButton97(GetControl('BOuvrir')).onClick := BouvrirClick;
  //
  //Gestion Restriction Domaine et Etablissements
  CC:=THValComboBox(GetControl('T_DOMAINE')) ;
  if CC<>Nil then PositionneDomaineUser(CC) ;

end ;

procedure TOF_BTFOURNMULTI_MUL.OnClose ;
begin
  Inherited ;
  if TOBFournSel.detail.count > 0 then TheTOB := TOBFOurnSel;
end ;

procedure TOF_BTFOURNMULTI_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTFOURNMULTI_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTFOURNMULTI_MUL.BouvrirClick(Sender: TObject);
var OneTOB : TOB;
    Q : TQuery ;
    i : integer;
begin
  Inherited ;
  Q := TFmul(F).Q;
  //
  if TFMul(F).Fliste.AllSelected then
  BEGIN
    Q.First;
    while Not Q.EOF do
    BEGIN
      OneTOB := TOB.Create ('TIERS',TOBFOurnSel,-1);
      OneTOB.putValue('T_TIERS',Q.FindField('T_TIERS').AsString);
      Q.NEXT;
    END;
    TFMul(F).Fliste.AllSelected:=False;
  END else
  begin
    for i:=0 to TFMul(F).Fliste.nbSelected-1 do
    begin
      TFMul(F).Fliste.GotoLeBookmark(i);
      OneTOB := TOB.Create ('TIERS',TOBFOurnSel,-1);
      OneTOB.putValue('T_TIERS',TFMul(F).Fliste.datasource.dataset.FindField('T_TIERS').AsString);
    end;
  end;
  //
  TFMul(F).Close;
end;

Initialization
  registerclasses ( [ TOF_BTFOURNMULTI_MUL ] ) ;
end.

