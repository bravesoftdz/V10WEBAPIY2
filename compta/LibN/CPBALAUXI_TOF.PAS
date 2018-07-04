{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 08/07/2002
Modifié le ... : 31/07/2003
Description .. : Source TOF de la FICHE : CPBALAUXI
Suite ........ :
Suite ........ : Gestion du QRS1 des balances auxiliaires
Suite ........ :
Suite ........ : Depuis 25/07/2003 : les traitements communs aux
Suite ........ : balances ont été remontés dans une tof intermédiaire :
Suite ........ : TOF_METH --> TOF_CPBALANCE --> TOF_CPBALAUXI
Suite ........ :
Suite ........ : JP 01/07/05 : Gestion des caractères Joker : fonctions de
Suite ........ : base définies dans TofMeth
Suite ........ :
Mots clefs ... : TOF;CPBALAUXI
*****************************************************************}
Unit CPBALAUXI_TOF ;

Interface

Uses StdCtrls, Classes,
{$IFDEF EAGLCLIENT}
     MainEAgl, utob, eQRS1,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     Fe_Main,
{$ENDIF}
     sysutils, Spin, Ent1,
     HCtrls, HEnt1, UTOF,
     TofMeth,
     hMsgBox,
     AGLInit,           // TheData
     CritEdt,           // ClassCritEdt
     uLibExercice,      // CInitComboExercice(E_EXERCICE);
     CPBALANCE_TOF,
     Paramsoc ;

procedure CPLanceFiche_BalanceAuxiliaire ( Args : String = '' ) ;

Type
  TOF_CPBALAUXI = Class (TOF_CPBALANCE)

    ChoixTypeTaLi   : THRadioGroup;
    NatCptMulti     : ThMultiValComboBox ;
    AuxiFermes             : THEdit;

    // EVT TOF
    procedure OnNew                    ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure ChargementCritEdt        ; override ;
    procedure SetTypeBalance           ; override ;

    // EVT FICHE
    procedure TableLibreChanged        (Sender: TObject); override ;
    procedure CompteOnExit             (Sender: TObject); override ;
    procedure AuxiElipsisClick         (Sender: TObject);
    procedure OnClickChoixTypeTaLi     (Sender: TObject);

    // TRAITEMENT
    // -> Construit la requête de l'état
    function  GenererRequeteBAL        : String ; override ;

    // Gestion des ruptures sur tables libres TIERSCOMPL
    procedure InitChoixTableLibre;
    Function  GetNumTableLibre         : String ; override ;

    function  GetWhereNature           : String ;  override ;
    procedure InitZones ;                          override ;
    procedure paramRuptures                     ;  override ;

    // Nouvelles fonctions de maj des comptes
    function  GetJointureCpt           : String ; override ;
    procedure RemplirEDTBALANCE            ; override ;

  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  uLibWindows, UTofMulParamGen; // TestJoker


//==============================================================================
procedure CPLanceFiche_BalanceAuxiliaire ( Args : String ) ;
begin
	AGLLanceFiche('CP','CPBALAUXI','','',Args);
end;

{TOF_CPBALAUXI}

//==============================================================================
//====================         EVT TOF        ==================================
//==============================================================================

procedure TOF_CPBALAUXI.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPBALAUXI.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPBALAUXI.OnLoad ;
var lStSQL : String ;
    lStLib : String ;
begin

  Inherited ;

  // Traduction des natures de pièces :
  TraductionTHMultiValComboBox( NatCptMulti, lStSQL, lStLib, 'T_NATUREAUXI' ) ;
  SetControlText( 'NATURECPTLIB', lStLib ) ;

end ;

procedure TOF_CPBALAUXI.OnArgument (S : String ) ;
begin

  // CONTROLES SPECIFIQUES
  // Spécifique tables libres TIERSCOMPL
  ChoixTypeTaLi    := THRadioGroup(GetControl('CHOIXTYPETALI',True));
  ChoixTypeTaLi.OnClick  := OnClickChoixTypeTaLi;

  // Nature de compte en multi
  NatCptMulti	        := THMultiValComboBox(Getcontrol('NATURECPT', true));
  NatCptMulti.OnChange  := NatureCptChanged ;

  AuxiFermes := THEdit(GetControl('AUXIFERMES', true));
  // Controles communs
  Inherited ;

  // Rubrique d'aide
  Ecran.HelpContext:=7448000;

  // Type de plan comptable :
  Case V_PGI.LaSerie of
    S7 :  begin
          Corresp.plus := 'AND (CO_CODE = "AU1" OR CO_CODE = "AU2")';
          end;
    else  begin
          Corresp.plus := 'AND CO_CODE = "AU1"';
          end;
    end;

  // GCO 08/08/2003
  CompteDe.MaxLength := VH^.CPta[fbAux].Lg;
  CompteA.MaxLength  := VH^.CPta[fbAux].Lg;

  // Type de balance de situation
  BalSit.Plus := ' BSI_TYPECUM = "TIE" ' ;

  // SBO - 28/06/2004 - Initialisation du type de table libre pour les ruptures
  ChoixTypeTaLi.Visible := TL_TIERSCOMPL_Actif;
  ChoixTypeTaLi.ItemIndex := 0;
  InitChoixTableLibre;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
  begin
    THEdit(GetControl('COMPTE', true)).OnElipsisClick:=AuxiElipsisClick;
    THEdit(GetControl('COMPTE_', true)).OnElipsisClick:=AuxiElipsisClick;
  end;

end ;

//==============================================================================
//====================        EVT QRS1        ==================================
//==============================================================================


procedure TOF_CPBALAUXI.TableLibreChanged(Sender: TObject);
begin
  if TableLibre.ItemIndex < 0 then
    begin
    LibreDe.DataType := '';
    LibreA.DataType  := '';
    LibreDe.Text     := '';
    LibreA.Text      := '';
    end
  else
    begin
    if ChoixTypeTaLi.Value = 'AUX' then
      begin
      LibreDe.DataType := 'TZNATTIERS' + GetNumTableLibre ;
      LibreA.DataType  := 'TZNATTIERS' + GetNumTableLibre ;
      end
    else if ChoixTypeTaLi.Value = 'CLI' then
      begin
      LibreDe.DataType := 'GCLIBRETIERS' + GetNumTableLibre ;
      LibreA.DataType  := 'GCLIBRETIERS' + GetNumTableLibre ;
      end
    else if ChoixTypeTaLi.Value = 'SCL' then
      begin
      LibreDe.DataType := 'AFLRESSOURCE';
      LibreA.DataType  := 'AFLRESSOURCE';
      end;
    end;
end;

//==============================================================================
//====================   TRAITEMENTS DES DONNEES  ==============================
//==============================================================================
procedure TOF_CPBALAUXI.CompteOnExit(Sender: TObject);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;
  if Trim(THEdit(Sender).Text) = '' then Exit;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if HasJoker(Sender) then Exit;

  // Complétion auto du numéro de compte si possible
  if not CompleteAuto( Sender, fbAux ) then
    THEdit(Sender).ElipsisClick(nil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 12/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPBALAUXI.AuxiElipsisClick(Sender: TObject);
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

procedure TOF_CPBALAUXI.InitChoixTableLibre;
begin

  if ChoixTypeTaLi.Value = 'AUX' then
    begin
    SetControlProperty('TABLELIBRE','DataType','TTTABLESLIBRESAUX');
    SetControlProperty('TABLELIBRE','Plus','');
    end
  else if ChoixTypeTaLi.Value = 'CLI' then
    begin
    SetControlProperty('TABLELIBRE','DataType','GCZONELIBRETIE');
    SetControlProperty('TABLELIBRE','Plus','AND CC_CODE LIKE "CT%"');
    end
  else if ChoixTypeTaLi.Value ='SCL' then
    begin
    SetControlProperty('TABLELIBRE','DataType','GCZONELIBRETIE');
    SetControlProperty('TABLELIBRE','Plus','AND CC_CODE LIKE "CR%"');
    end;

  TableLibre.ReLoad;

  //SG6 14.03.05 Gestion des filtres FQ 14996
  if cLoadFiltre then Exit;

  TableLibre.ItemIndex := 0 ;
  SetControlText ('LIBREDE','');
  SetControlText ('LIBREA','');

end;

procedure TOF_CPBALAUXI.OnClickChoixTypeTaLi(Sender: TObject);
begin
  InitChoixTableLibre;
  TableLibreChanged ( nil );
end;

function TOF_CPBALAUXI.GenererRequeteBAL: String;
begin
  // Champs
{  if FBoRegroupement and ( GetControlText('MULTIDOSSIER') <> '' )
    then result := 'SELECT SYSDOSSIER, CED_COMPTE CED_AUXILIAIRE, CED_NATURE, CED_LIBELLE'
    else }

  result := 'SELECT CED_COMPTE CED_AUXILIAIRE, CED_NATURE, CED_LIBELLE' ;
  result := result
               + ', CED_DEBIT1, CED_CREDIT1'
               + ', CED_DEBIT2, CED_CREDIT2'
               + ', CED_RUPTURE, CED_COLLECTIF ' ;

  if Rupture.Value <> 'SANS' then
    begin
    // Rupture sur compte de correspondance
    if (RuptureType.Value = 'RUPCORRESP') and (Corresp.ItemIndex >= 0) then
      result := result + ', CR_LIBELLE RUPTURELIB '
    // Rupture sur tables libres
    else if (RuptureType.Value = 'RUPLIBRES') and (TableLibre.ItemIndex >= 0) then
      begin
      if ChoixTypeTaLi.Value = 'AUX' then
        result := result + ', NT_LIBELLE RUPTURELIB '
      else if ChoixTypeTaLi.Value = 'CLI' then
        result := result + ', YX_LIBELLE RUPTURELIB '
      else if ChoixTypeTaLi.Value = 'SCL' then
        result := result + ', ARS_LIBELLE RUPTURELIB ' ;
      end
    // Nothing
    else result := result + ', "&#@" RUPTURELIB ' ;
    end
  // Pas de rupture
  else result := result + ', "&#@" RUPTURELIB ' ;

  // Tables
  result := result + ' FROM CEDTBALANCE ' ;
  if Rupture.Value <> 'SANS' then
    begin
    if (RuptureType.Value = 'RUPCORRESP') and (Corresp.ItemIndex >= 0) then
      result := result + ' LEFT JOIN CORRESP ON CR_CORRESP=CED_RUPTURE AND CR_TYPE="' + Corresp.Value + '"'
    else if (RuptureType.Value = 'RUPLIBRES') and (TableLibre.ItemIndex >= 0) then
      begin
      if ChoixTypeTaLi.Value = 'AUX' then
        result := result + ' LEFT JOIN NATCPTE ON NT_NATURE=CED_RUPTURE AND NT_TYPECPTE="' + TableLibre.Value + '"'
      else if ChoixTypeTaLi.Value = 'CLI' then
        result := result + ' LEFT JOIN CHOIXEXT ON YX_CODE=CED_RUPTURE AND YX_TYPE="LT' + GetNumTableLibre + '"'
      else if ChoixTypeTaLi.Value = 'SCL' then
        result := result + ' LEFT JOIN RESSOURCE ON ARS_RESSOURCE=CED_RUPTURE '
      end ;
    end ;

  // Conditions
  result := result + ' WHERE CED_USER = "' + V_PGI.User + '"' ;
  // Tri
  result := result + ' ORDER BY CED_RUPTURE, CED_AUXILIAIRE' ;

end;

function TOF_CPBALAUXI.GetNumTableLibre: String;
begin
  result := '' ;
  if TableLibre.ItemIndex < 0 then Exit ;
  result := Copy( TableLibre.Value, 3, 1 )  ;

end;


function TOF_CPBALAUXI.GetWhereNature: String;
var lStTous : String ;
    lStVal  : String ;
    lStText : String ;
    lTest   : String ;
begin
  result := '' ;
  lStTous := '<<' + TraduireMemoire('Tous') + '>>';
  if NatCptMulti.Tous or ( NatCptMulti.Text = lStTous ) or ( NatCptMulti.Text = '' ) then Exit ;

  lStText := NatCptMulti.Text ;
  ltest := NatCptMulti.getSQLValue ;

  while ( lStText <> '' ) do
    begin
    lStVal := readTokenSt(  lStText ) ;
    if Trim(lStVal) <> '' then
      if result = ''
        then result := '"' + lStVal + '"'
        else result := result + ',"' + lStVal + '"' ;
    end ;

  if result <> '' then
    result := ' T_NATUREAUXI IN (' + result + ') ' ;

end;

procedure TOF_CPBALAUXI.InitZones;
begin
  inherited;
  NatCptMulti.text := '<<' + TraduireMemoire( 'Tous' ) + '>>' ;
end;

procedure TOF_CPBALAUXI.ChargementCritEdt;
begin
  // FQ 15844 SBO 27/09/2005 : Application des critères sur nature auxi
  if (TheData <> nil) and (TheData is ClassCritEdt) then
    begin
    // Nature auxiliaire en multi :
    NatCptMulti.Text := ClassCritEdt(TheData).CritEdt.NatureCpt ;
    end ;

  inherited;

end;



function TOF_CPBALAUXI.GetJointureCpt: String;
begin
  if ( Rupture.value <> 'SANS' )  and ( RuptureType.Value = 'RUPLIBRES' ) and
     ( ( ChoixTypeTaLi.Value = 'CLI' ) or ( ChoixTypeTaLi.Value = 'SCL' ) )
    then result := result + ' LEFT JOIN ' + GetTablePourBase('TIERSCOMPL') + ' ON YTC_AUXILIAIRE = T_AUXILIAIRE'
    else result := '' ;
end;

procedure TOF_CPBALAUXI.SetTypeBalance;
begin
  TypeBal     := balAuxi ;
  TableEcr    := 'ECRITURE' ;
  PfEcr       := 'E' ;
  TableCpt    := 'TIERS' ;
  PfCpt       := 'T' ;
  ChampCpt    := 'AUXILIAIRE' ;
  ChampNatCpt := 'T_NATUREAUXI' ;
end;

procedure TOF_CPBALAUXI.paramRuptures;
begin
  inherited;

  // Les ruptures sur champ libre tiers
  if ( Rupture.Value <> 'SANS')  and (RuptureType.Value = 'RUPLIBRES') and (TableLibre.ItemIndex >= 0) then
    begin
      // Tables libres tiers
      if ChoixTypeTaLi.Value = 'AUX'
        then XXRupture.Text := 'T_TABLE' + GetNumTableLibre
      // Table libre clients
      else if ChoixTypeTaLi.Value = 'CLI'
        then XXRupture.Text := 'YTC_TABLELIBRETIERS' + GetNumTableLibre
      // Tables libres ressource
      else if ChoixTypeTaLi.Value = 'SCL'
        then XXRupture.Text := 'YTC_RESSOURCE' + IntToStr(TableLibre.ItemIndex + 1) ;
    end ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 31/01/2006
Modifié le ... : 12/09/2006
Description .. : Ne pas prendre les comptes auxi fermés FQ 17207
Mots clefs ... :
*****************************************************************}
procedure TOF_CPBALAUXI.RemplirEDTBALANCE ;
begin
  inherited;
  //YMO 22/09/2006 prise en compte de l'utilisateur
  IF ExecuteSQL('DELETE FROM CEDTBALANCE WHERE CED_USER = "' + V_PGI.User + '"'
  +' AND (CED_DEBIT1=0 AND CED_CREDIT1=0)'
  +' AND CED_COMPTE ='
  +'(SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE = CED_COMPTE AND T_FERME = "X")')>0 then
      AuxiFermes.Text := 'X'
  else
      AuxiFermes.Text := '-';
end;

Initialization
  registerclasses ( [ TOF_CPBALAUXI ] ) ;
end.
