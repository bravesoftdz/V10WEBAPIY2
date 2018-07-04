{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 06/02/2002
Modifié le ... : 04/03/2003
Description .. : Source TOF de la FICHE : CPECRPRORATISEES ()
Suite ........ : Passage en eAGL
Mots clefs ... : TOF;CPECRPRORATISEES
*****************************************************************}
Unit CPECRPRORATISEES_TOF ;

Interface

Uses StdCtrls,    // CbChecked, CbUnchecked, cbGrayed, TCheckBox
     Classes,     // registerclasses
{$IFDEF EAGLCLIENT}
     MaineAGL,    // AGLLanceFiche
     eMul,        // TFMul
     UTOB,        // TQuery
{$ELSE}
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     Fe_main,     // AGLLanceFiche
     Mul,         // TFMul
     Hdb,         // THDBGrid
     SaisBor,     //  LanceSaisieFolio
{$ENDIF}
     Saisie,      // TrouveEtLanceSaisie
     Forms,       // TForm
     sysutils,    // DateToStr, StrToDate, IntToStr, Date
     ComCtrls,    // TTabSheet
     HCtrls,      // THGrid, OpenSQL, Ferme, THLabel, THEdit, ReadTokenSt, THValComboBox
     HEnt1,       // EstSerie, TaConsult
     Ent1,        // VH
     SaisUtil,    // QuelExoDt
     UTOF          // TOF
     ;

procedure CCLanceFiche_EcrTVAProratisees;

Type
  TOF_CPECRPRORATISEES = Class (TOF)
    Fiche : TFMul ;
{$IFDEF EAGLCLIENT}
    ListeEcr : THGrid ;
    FListe   : THGrid ;
{$ELSE}
    ListeEcr : THDBGrid ;
    FListe   : THDbGrid ;
{$ENDIF}
    QlisteEcr : TQuery ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    //SG6 12/01/05 FQ 15155 Gestion des ecritures d'immo
    procedure qualifpiecechange(Sender : TObject);
    procedure ExerciceOnChange(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure TxInitClick(Sender: TObject);
    procedure TxDefinClick(Sender: TObject);
    procedure TxInterClick(Sender: TObject);
    procedure InitTablesLibres(TT: TTabSheet);
    procedure InitValid;
  public
    ComboEtab: THValComboBox ;

    {JP 28/06/06 : FQ 16149 : gestion des réstrictions Etablissements et à défaut des ParamSoc}
    procedure GereEtablissement;
    {JP 28/06/06 : FQ 16149 : on s'assure que le filtre coincide avec les restrictions utilisateurs sur l'établissement}
    procedure ControlEtab;
    {JP 28/06/06 : FQ 16149 : pour appeler ControlEtab à la fin du chargement des filtres}
    procedure AfterSelectFiltre;
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  ULibWindows; //TraductionTHMultiValComboBox //SG6 12/01/05 FQ 15155


procedure CCLanceFiche_EcrTVAProratisees;
begin
  AGLLanceFiche ('CP' , 'CPECRPRORATISEE', '', '' , '' );
end;

procedure TOF_CPECRPRORATISEES.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPECRPRORATISEES.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPECRPRORATISEES.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPECRPRORATISEES.OnLoad ;
begin
  Inherited ;
end ;

Procedure TOF_CPECRPRORATISEES.ExerciceOnChange (Sender : TObject);
var Q : TQuery ;
    FinDate : TDateTime ;
BEGIN
Q := OpenSql('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="'+GetControlText('E_Exercice')+'"', True) ;
SetControlText('E_DATECOMPTABLE',DateToStr(Q.Fields[0].AsDateTime)) ;
SetControlText('E_DATECOMPTABLE_',DateToStr(Q.Fields[1].AsDateTime)) ;
SetControlText('DATEGEN',DateToStr(Q.Fields[1].AsDateTime)) ;
FinDate := Q.Fields[1].AsDateTime ;
if FinDate >= StrToDate('01/01/2002') then
  Begin
  SetControlEnabled('Money',False) ; SetControlText('Money','E') ;
  end else
  begin
  SetControlEnabled('Money',True) ; if VH^.TenueEuro then SetCOntrolText('Money','E') else SetControlText('Money','F') ;
  end ;
Ferme(Q) ;
END ;

procedure TOF_CPECRPRORATISEES.InitTablesLibres ( TT : TTabSheet ) ;
Var LesLib : HTStringList ;
    i : Integer ;
    St,Titre,Ena : String ;
    Trouver : Boolean ;
    LL      : THLabel ;
    CC      : THEdit ;
BEGIN
if TT=Nil then Exit ; Trouver:=False ;
LesLib:=HTStringList.Create ; GetLibelleTableLibre('G',LesLib) ;
for i:=0 to LesLib.Count-1 do
  BEGIN
  St:=LesLib.Strings[i] ; Titre:=ReadTokenSt(St) ; Ena:=St ;
  LL:=THLabel(TForm(TT.Owner).FindComponent('TG_TABLE'+IntToStr(i))) ;
  CC:=THEdit(TForm(TT.Owner).FindComponent('G_TABLE'+IntToStr(i))) ;
  if LL<>Nil then
    BEGIN
    LL.Caption:=Titre ; LL.Enabled:=(Ena='X') ;
    if ((EstSerie(S3)) and (i>2)) then LL.Visible:=False ;
    if CC<>Nil then
      BEGIN
      CC.Enabled:=LL.Enabled ; if CC.Enabled then Trouver:=True ;
      if ((EstSerie(S3)) and (i>2)) then CC.Visible:=False ;
      END ;
    END ;
  END ;
TT.TabVisible:=Trouver ;
LesLib.Clear ; LesLib.Free ;
END ;

Procedure TOF_CPECRPRORATISEES.FListeDblClick (Sender: TObject);
Var sMode : String ;
    Fiche : TFMul ;
    sSimul : String ;
begin
inherited;
Fiche := TFMul(Ecran) ;
ListeEcr := Fiche.FListe ;
{$IFDEF EAGLCLIENT}
QListeEcr := Fiche.Q.TQ ;
Fiche.Q.TQ.Seek(Fiche.FListe.Row-1);
{$ELSE}
QListeEcr := Fiche.Q ;
{$ENDIF}
if ((QListeEcr.EOF) and (QListeEcr.BOF)) then Exit ;
sMode:=QListeEcr.FindField('E_MODESAISIE').AsString ;
sSimul:=QListeEcr.FindField('E_QUALIFPIECE').AsString ;
if ((sMode<>'') and (sMode<>'-')) then
{$IFDEF EAGLCLIENT}
  // A FAIRE Voir SAISBOR.PAS
{$ELSE}
  LanceSaisieFolio(QListeEcr,TaConsult)
{$ENDIF}
// YMO FQ 16686 On envoyait "N" en dur au lieu du type d'écriture
  else TrouveEtLanceSaisie(QListeEcr,TaConsult, sSimul) ;
end;

Procedure TOF_CPECRPRORATISEES.TxInitClick (Sender : TObject) ;
begin
If GetCheckBoxState('TxInit')=CbChecked then SetControlText('XX_WHERE1','E_QUALIFORIGINE="TV1"') ;
If GetCheckBoxState('TxInit')=CbUnChecked then SetControlText('XX_WHERE1','((E_QUALIFORIGINE<>"TV1") OR (E_QUALIFORIGINE IS NULL))') ;
If GetCheckBoxState('TxInit')=CbGrayed then SetControlText('XX_WHERE1','') ;
end ;

Procedure TOF_CPECRPRORATISEES.TxInterClick (Sender : TObject) ;
begin
If GetCheckBoxState('TxInter')=CbChecked then SetControlText('XX_WHERE3','E_QUALIFORIGINE="TVI"') ;
If GetCheckBoxState('TxInter')=CbUnChecked then SetControlText('XX_WHERE3','((E_QUALIFORIGINE<>"TVI") OR (E_QUALIFORIGINE IS NULL))') ;
If GetCheckBoxState('TxInter')=CbGrayed then SetControlText('XX_WHERE3','') ;
end ;

Procedure TOF_CPECRPRORATISEES.TxDefinClick (Sender : TObject) ;
begin
If GetCheckBoxState('TxDefin')=CbChecked then SetControlText('XX_WHERE2','E_QUALIFORIGINE="TV2"') ;
If GetCheckBoxState('TxDefin')=CbUnChecked then SetControlText('XX_WHERE2','((E_QUALIFORIGINE<>"TV2") OR (E_QUALIFORIGINE IS NULL))') ;
If GetCheckBoxState('TxDefin')=CbGrayed then SetControlText('XX_WHERE2','') ;
end ;

procedure TOF_CPECRPRORATISEES.OnArgument (S : String ) ;
var
  Devise      : THValComboBox;            {FP 08/11/2005 FQ15457}
begin
SetControlVisible('BSELECTALL', False);
Ecran.HelpContext := 1553;
ThValComboBox(GetControl('E_Exercice')).OnChange := ExerciceOnChange ;
SetControlText('E_EXERCICE',QuelExoDt(Date)) ;
{$IFDEF EAGLCLIENT}
THGrid(GetControl('FLISTE')).OnDblClick := FListeDblClick;
{$ELSE}
THDBGrid(GetControl('FLISTE')).OnDblClick := FListeDblClick;
{$ENDIF}
TcheckBox(GetControl('TXINIT')).OnCLick := TxInitClick ;
TcheckBox(GetControl('TXINTER')).OnCLick := TxInterClick ;
TcheckBox(GetControl('TXDEFIN')).OnCLick := TxDefinClick ;
//SG6 12/01/05 Gestion des ecritures de simul
THMultiValComboBox(getcontrol('QUALIFPIECE',true)).OnChange := qualifpiecechange;
THMultiValComboBox(getcontrol('QUALIFPIECE',true)).Plus := 'AND (CO_CODE="N" OR CO_CODE="S")';

InitTablesLibres (TTabSheet(GetControl('TabLib'))) ;
  Inherited ;
  {JP 06/06/2003 : pour mettre à jour le titre de la fenêtre}
  Ecran.Caption:='Liste des écritures' ;
  UpdateCaption(Ecran);

  InitValid;

  {b FP 08/11/2005 FQ15457}
  Devise := THValComboBox(getcontrol('E_DEVISE',true));
  Devise.Enabled := False;               {FP 01/02/2006 FQ15457}
  Devise.Value   := V_PGI.DevisePivot;
  {e FP 08/11/2005 FQ15457}

  {JP 28/06/06 : FQ 16149 : refonte de la gestion des établissements}
  ComboEtab := THValComboBox(GetControl('E_ETABLISSEMENT'));
  GereEtablissement;
  TFMul(Ecran).OnAfterSelectFiltre := AfterSelectFiltre;;
end ;

procedure TOF_CPECRPRORATISEES.OnClose ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 12/01/2005
Modifié le ... :   /  /
Description .. : Procedure qui permet de gérer la condition générée par le
Suite ........ : controle QUALIFPIECE
Mots clefs ... : QUALIFPIECE
*****************************************************************}
procedure TOF_CPECRPRORATISEES.qualifpiecechange(Sender : TObject);
var
  sSql,tmp1,tmp2 : string;
  qualifp : THMultiValComboBox;
begin
  qualifp := THMultiValComboBox(Sender);
  tmp2 := 'E_QUALIFPIECE';
  if qualifp.Text = '<<Tous>>' then sSql := '(E_QUALIFPIECE = "S" or E_QUALIFPIECE = "N")'
  else TraductionTHMultiValComboBox(qualifp,sSql,tmp1,tmp2,false);
  SetControlText('XX_WHERE5',sSql);

  InitValid;
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 25/08/2005
Modifié le ... :   /  /
Description .. : Case 'écriture validée O/N' uniquement si écritures normales
Suite ........ :
Mots clefs ... : QUALIFPIECE VALID
*****************************************************************}
procedure TOF_CPECRPRORATISEES.InitValid;
begin
  if (getcontroltext('QUALIFPIECE') = 'N')
  or (getcontroltext('QUALIFPIECE') = 'N;') then
  begin
      SetControlProperty('E_VALIDE', 'VISIBLE', True);
      SetControlProperty('FE__HLabel4', 'VISIBLE', True);
  end
  else
  begin
      SetControlProperty('E_VALIDE', 'VISIBLE', False);
      SetControlProperty('FE__HLabel4', 'VISIBLE', False);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPECRPRORATISEES.GereEtablissement;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(ComboEtab) then begin
    {Si l'on ne gère pas les établissement ...}
    if not VH^.EtablisCpta  then begin
      {... on affiche l'établissement par défaut}
      ComboEtab.Value := VH^.EtablisDefaut;
      {... on désactive la zone}
      ComboEtab.Enabled := False;
    end

    {On gère l'établisement, donc ...}
    else begin
      {... On commence par regarder les restrictions utilisateur}
      PositionneEtabUser(ComboEtab);
      {... s'il n'y a pas de restrictions, on reprend le paramSoc
       JP 25/10/07 : FQ 19970 : Finalement on oublie l'option de l'établissement par défaut
      if ComboEtab.Value = '' then begin
        {... on affiche l'établissement par défaut
        ComboEtab.Value := VH^.EtablisDefaut;
        {... on active la zone
        ComboEtab.Enabled := True;
      end;}
    end;
  end;

end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPECRPRORATISEES.ControlEtab;
{---------------------------------------------------------------------------------------}
var
  Eta : string;
begin
  if not Assigned(ComboEtab) then Exit;
  {S'il n'y a pas de gestion des établissement, logiquement, on ne force pas l'établissement !!!}
  if not VH^.EtablisCpta then Exit;

  Eta := EtabForce;
  {S'il y a une restriction utilisateur et qu'elle ne correspond pas au contenu de la combo ...}
  if (Eta <> '') and (Eta <> ComboEtab.Value) then begin
    {... on affiche l'établissement des restrictions}
    ComboEtab.Value := Eta;
    {... on désactive la zone}
    ComboEtab.Enabled := False;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPECRPRORATISEES.AfterSelectFiltre;
{---------------------------------------------------------------------------------------}
begin
  ControlEtab;
end;


Initialization
  registerclasses ( [ TOF_CPECRPRORATISEES ] ) ;
end.
