{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 20/02/2003
Modifié le ... :   /  /
Description .. : Passage en eAGL
Mots clefs ... :
*****************************************************************}
unit UTOFMULPARAMGEN;

interface              
uses

  Controls, Windows, Messages, SysUtils, Classes, Forms, Dialogs, Menus,graphics,
  HTB97,  // TToolBarButton97
  HCtrls, // ReadTokenST
  ComCtrls, // TTabSheet
  StdCtrls, // TCheckbox
{$IFDEF EAGLCLIENT}
  MaineAGL,
  eMul,
  uTob,
{$ELSE}
  FE_Main,
  Mul,
  AGLInit,
  HDB,       // THDBGrid
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  HEnt1,
  Ent1,     // LibellesTableLibre
  HPop97,
  Buttons,  // blGlyphBottom
  HMsgBox,
  Filtre,
{$IFNDEF COMPTAPAIE}
{$IFNDEF GCGC}
//$IFNDEF CCADM}
  AGlStdcompta,
  OuvrFermcpte,
  AGlFctSuppr,
//$ENDIF CCADM}
{$ENDIF COMPTAPAIE}
{$ENDIF GCGC}
  utilPGI, // EstTableCommune
  UTOF  ,UentCommun;

procedure CCLanceFiche_MULGeneraux(pszArg : String);
function  CPLanceFiche_MULTiers(pszArg : String) : String;  {YMO 12/04/07}
procedure CPLanceFiche_MULSection(pszArg : String);
procedure CPLanceFiche_MULJournal(pszArg : String);

type
  TOF_MULPARAMGEN = Class (TOF)
  {$IFDEF EAGLCLIENT}
    FListe : THGrid;
  {$ELSE}
    FListe : THDBGrid;
  {$ENDIF EAGLCLIENT}
    procedure OnArgument(stArgument: String); override ;
    procedure OnLoad ; override ;
    procedure OnClose ; override ;
    procedure OnDisplay ; override ;
    procedure FormKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
    procedure OnclickParam(Sender: TObject);
    procedure OnClickVentilable(Sender : TObject);
    procedure OnAfterShow;
  private
    szExtended    : String ;
    gszFonction : String ;
    FBoToucheCtrl : boolean ;
    FEnChargement : boolean ;
    fOnSaveKeyDownEcran   : procedure (Sender : TObject; var Key: Word; Shift: TShiftState) of Object;
    FOnBChercheClick      : TNotifyEvent ;
    procedure InitXX_Where ( vStWhere : string = '' ) ;
    function  getTableName : String ;
    {$IFNDEF GCGC}
    function  getFBName    : String ;
    function  getAxe       : String ;
    function  getFieldName : String ;
    function  getTypeAction : String ;
    {$ENDIF !GCGC}
    function  getNomEcran   : String ;
    procedure InitPresentation ;
    procedure InitEvenement ;
    procedure DetermineListe ;
    procedure BChercheClick   ( Sender : TObject ) ;
    {$IFNDEF GCGC}
    procedure BSelectAllClick ( Sender : TObject ) ;
    {$ENDIF !GCGC}

{$IFDEF COMPTAPAIE}
    procedure AfficheRibCompte ( Sender : TObject );
{$ENDIF}

{$IFNDEF COMPTAPAIE}
{$IFNDEF GCGC}
//$IFNDEF CCADM}
    procedure FermeCpteOnClick(Sender : TObject);
    procedure OuvreCpteOnClick(Sender : TObject);
    procedure SupprCpteOnClick(Sender : TObject);
    procedure ImprCpteOnClick(Sender : TObject);
    procedure InsertCpteOnClick(Sender : TObject);
    procedure ModifSerieOnClick(Sender : TObject);
    procedure ModifCpteOnClick (Sender : TObject);
    procedure AxeChange        (Sender : TObject);
//$ENDIF CCADM}
{$ENDIF COMPTAPAIE}
{$ENDIF GCGC}
    procedure SetVisibleBSelectAll( condition : boolean ); // BVE 15.05.07 FQ 20183
end;

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  {$IFDEF COMPTAPAIE}
  BanqueCP_TOM, // FicheBanqueCP
  {$ENDIF}
  ParamSoc,
  UFonctionsCBP;

// Paramètres :
// 1°- C : Consultation, création, modification ;
//     S : Suppression
//     I : Impression    A FAIRE
//     O : Ouverture
//     F : Fermeture
// 2°- N° d'HelpContext
procedure CCLanceFiche_MULGeneraux(pszArg : String);
begin
  V_PGI.ExtendedFieldSelection := '';
  AGLLanceFiche('CP','MULGENERAUX','','',pszArg);
end;

function CPLanceFiche_MULTiers(pszArg : String) : String;
begin
  V_PGI.ExtendedFieldSelection := '';
  result:=AGLLanceFiche('CP','MULTIERS','','',pszArg);   {YMO 12/04/07 fonction pour F5 sur les THEdit Auxiliaire}
end;

procedure CPLanceFiche_MULSection(pszArg : String);
begin
  V_PGI.ExtendedFieldSelection := '';
  AGLLanceFiche('CP','MULSECTION','','',pszArg);
end;

procedure CPLanceFiche_MULJournal(pszArg : String);
begin
  V_PGI.ExtendedFieldSelection := '';
  AGLLanceFiche('CP','MULJOURNAL','','',pszArg);
end;

// Paramètres :
// 1°- C : Consultation, création, modification ;
//     S : Suppression
//     I : Impression    A FAIRE
//     O : Ouverture
//     F : Fermeture
// 2°- N° d'HelpContext
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/09/2004
Modifié le ... : 10/09/2007
Description .. : - LG - 10/09/2004 - FB 14527 - on affichait plus les compte
Suite ........ : en S5
Suite ........ : - FB 15172 - LG - 04/01/2004 - concept applicatif "Création
Suite ........ : des comptes généraux" sa seule utilisation masque
Suite ........ : automatiquement les icônes "Nouveau" dans la liste des
Suite ........ : comptes auxiliaires, des journaux ainsi que dans les sections
Suite ........ : analytiques.
Suite ........ : - LG - 21/12/2006 - on ne regarde plus si une saisie et 
Suite ........ : ouverte qd on ouvre le mul pour selectionner un auxiliaire 
Suite ........ : depuis la saisie
Suite ........ : - LG - 10/09/2007 - FB 19217 - correction du titre d'un hint
Mots clefs ... : 
*****************************************************************}
procedure TOF_MULPARAMGEN.OnArgument(stArgument: String);
var szHelpContext : String ;
begin
{$IFDEF EAGLCLIENT}
  FListe := THGrid(GetControl('FListe', True));
{$ELSE}
  FListe := THDBGrid( GetControl('FListe', True));
{$ENDIF EAGLCLIENT}
 FEnChargement := true ;
 FBoToucheCtrl := GetKeyState(VK_CONTROL) < 0 ;
  // Récupère les arguments
  gszFonction := ReadTokenSt(stArgument);
  if gszFonction <> 'M' then
   begin
    szHelpContext := ReadTokenSt(stArgument);
    szExtended    := ReadTokenSt(stArgument);
   end ;

  if szHelpContext<> '' then Ecran.HelpContext := StrToInt(szHelpContext);

  //SG6 28.02.05
  DetermineListe ;

  if GetNomEcran = 'MULGENERAUX' then
  begin
    // Ajout GCO - 20/02/2007
    THEdit(GetControl('G_GENERAL', True)).MaxLength := VH^.Cpta[fbGene].lg;

    TCheckbox(GetControl('G_VENTILABLE1', True)).State := cbGrayed;
    {$IFNDEF COMPTAPAIE}
    SetControlVisible ('BINSERT', ( ExJaiLeDroitConcept(TConcept(ccGenCreat),False) ) and ( gszFonction='C' ) and ( not SaisieLancer ) ) ;
    {$ELSE}
    {VG 30/09/2004 - FQ N°11633
    IFDEF COMPTAPAIE, BINSERT et BOUVRIR doivent être invisibles car pas de prise de
    contrôle dans le InitEvenement. BINSERT OK dans procédure InitPresentation.
    }
    SetControlVisible ('BOUVRIR', False);
    {$ENDIF}

    // YM 24/08/2005 : FQ 15618 : On n'affiche que les comptes ouverts si menu fermer
    //                            et inversement, cette case est inutile
    if (gszFonction = 'O') or (gszFonction = 'F') then
    begin
          SetControlProperty('G_FERME', 'VISIBLE', False);

          {b FP FQ16012}
          if gszFonction = 'O' then
            TFMul(Ecran).FNomFiltre:='MULFERGENERAUX'
          else if gszFonction = 'F' then
            TFMul(Ecran).FNomFiltre:='MULOUVGENERAUX' ;
          {e FP FQ16012}
    end;
  end

  else if GetNomEcran = 'MULJOURNAL' then
  begin
    SetControlVisible ('BINSERT', ( ExJaiLeDroitConcept(TConcept(ccJalCreat),False) ) and ( not SaisieLancer ) );

    {b FP FQ16012}
    if gszFonction = 'O' then
      TFMul(Ecran).FNomFiltre:='MULFERJOURNAL'
    else if gszFonction = 'F' then
      TFMul(Ecran).FNomFiltre:='MULOUVJOURNAL' ;
    {e FP FQ16012}
  end

  else if GetNomEcran = 'MULSECTION' then
  begin
    // YM 24/08/2005 : FQ 15618 : On n'affiche que les comptes ouverts si menu fermer
    //                            et inversement, cette case est inutile
    if (gszFonction = 'O') or (gszFonction = 'F') then
    begin
      SetControlProperty('S_FERME', 'VISIBLE', False);
          {b FP FQ16012}
      if gszFonction = 'O' then
        TFMul(Ecran).FNomFiltre:='MULFERSECTION'
      else if gszFonction = 'F' then
        TFMul(Ecran).FNomFiltre:='MULOUVSECTION' ;
          {e FP FQ16012}
    end;
  end

  else
  if GetNomEcran = 'MULTIERS' then
  begin
    // Ajout GCO - 20/02/2007
    THEdit(GetControl('T_AUXILIAIRE', True)).MaxLength := VH^.Cpta[fbAux].lg;

    if ( gszFonction = 'M' ) then
     begin
      SetControlVisible ('BINSERT', ( ExJaiLeDroitConcept(TConcept(ccAuxCreat),False) ) ) ;
      TToolBarButton97( GetControl('BOUVRCPTE1') ).Hint := TraduireMemoire('Valider') ;
     end
    else
      SetControlVisible ('BINSERT', ( ExJaiLeDroitConcept(TConcept(ccAuxCreat),False) ) and ( not SaisieLancer ) ) ;

    // SBO 18/02/2005 En attente de maj de SocRef...
    SetControlProperty('T_TABLE0', 'DATATYPE', 'TZNATTIERS0') ;
    SetControlProperty('T_TABLE1', 'DATATYPE', 'TZNATTIERS1') ;
    SetControlProperty('T_TABLE2', 'DATATYPE', 'TZNATTIERS2') ;
    SetControlProperty('T_TABLE3', 'DATATYPE', 'TZNATTIERS3') ;
    SetControlProperty('T_TABLE4', 'DATATYPE', 'TZNATTIERS4') ;
    SetControlProperty('T_TABLE5', 'DATATYPE', 'TZNATTIERS5') ;
    SetControlProperty('T_TABLE6', 'DATATYPE', 'TZNATTIERS6') ;
    SetControlProperty('T_TABLE7', 'DATATYPE', 'TZNATTIERS7') ;
    SetControlProperty('T_TABLE8', 'DATATYPE', 'TZNATTIERS8') ;
    SetControlProperty('T_TABLE9', 'DATATYPE', 'TZNATTIERS9') ;

    {b FP FQ16012}
    if gszFonction = 'O' then
      TFMul(Ecran).FNomFiltre:='MULFERTIERS'
    else if gszFonction = 'F' then
      TFMul(Ecran).FNomFiltre:='MULOUVTIERS' ;
    {e FP FQ16012}
  end ;

{$IFDEF CCS3}
  if GetNomEcran = 'MULGENERAUX' then begin
    SetControlVisible('TVENTILABLE', False);
    SetControlVisible('VENTILABLE', False);
    if not EstComptaSansAna then SetControlVisible('G_VENTILABLE1', True)
                            else TCheckbox(GetControl('G_VENTILABLE1', True)).State := cbGrayed;

    end
  else if (GetNomEcran = 'MULJOURNAL') and EstComptaSansAna then begin
    SetControlProperty('J_NATUREJAL', 'PLUS', 'AND CO_CODE<>"ODA" AND CO_CODE<>"ANA"');
  end;
{$ENDIF}


  {FP FQ15996 Masque le bouton 'Nouveau' si O: Ouverture ou F: Fermeture}
  if (gszFonction = 'O') or (gszFonction = 'F') then
    SetControlVisible ('BINSERT', False);
  {e FP FQ15996}

  // Mise en place des évènements contextuels
  InitPresentation ;
  InitEvenement ;

  // Maj du titre
  Ecran.Caption:=TraduireMemoire(Ecran.Caption) ;
  UpdateCaption(Ecran);

 // Conditions supplémentaires
 if gszFonction = 'M' then
  begin
   if FBoToucheCtrl then
    begin
     SetControlText('T_LIBELLE',ReadTokenSt(stArgument) ) ;
     SetControlText('XX_ORDERBY', '' ) ;
     if Ecran <> nil then
      SetControlText('XX_ORDERBY', ' T_LIBELLE ' ) ;
    end
     else
      SetControlText('T_AUXILIAIRE',ReadTokenSt(stArgument) ) ;
   InitXX_Where(ReadTokenSt(stArgument))
  end 
   else
    InitXX_Where ;
end;

procedure TOF_MULPARAMGEN.OnLoad;
begin
  { YMO 30/12/2005 FQ 16434 Ce code était dans le OnArgument, au mauvais endroit car
  V_PGI.ExtendedFieldSelection est réinitialisée dans le Formshow de mul.pas
  Pour bien faire, on pouvait affecter la variable dans le AGLLanceFiche,
  mais celà remettait en question des modifs d'autres FQ, en particulier la 13689}
  if (szExtended <> '') then V_PGI.ExtendedFieldSelection := szExtended;
  Ecran.OnKeyDown :=  FormKeyDown;
  if V_PGI.ExtendedFieldSelection = '1' then begin
    SetControlVisible ('BSUPPRIMER', False);
    SetControlVisible ('BOUVRE', False);
    SetControlVisible ('BINSERT', False);
  end;
end;

procedure TOF_MULPARAMGEN.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var POPMENUF11 : TPopUpMenu;
begin
  case key of
    VK_F5  : begin ;
              if  FEnChargement then
               Key := 0
                else
                 fOnSaveKeyDownEcran(Sender, Key, Shift ) ;
             end ;
    VK_F10 : if ( gszFonction = 'M' ) then
              begin
               Key := 0 ;
               TFMul(Ecran).Retour := GetField('T_AUXILIAIRE') ;
               SendMessage(Ecran.Handle, WM_CLOSE, 0, 0) ;
              end ;
    VK_RETURN :
      begin
        if ( gszFonction = 'M' ) then
         begin
          TFMul(Ecran).Retour := GetField('T_AUXILIAIRE') ;
          SendMessage(Ecran.Handle, WM_CLOSE, 0, 0)
         end
          else
           if not (ctxPCL in V_PGI.PGIContexte) then
           begin
             if (gszFonction = 'S') then TToolBarButton97(GetControl('BSUPPRIMER')).OnClick(Sender);  // Suppression
             if (gszFonction = 'F') then TToolBarButton97(GetControl('BFERMECPTE1')).OnClick(Sender); // Fermeture
             if (gszFonction = 'O') then TToolBarButton97(GetControl('BOUVRCPTE1')).OnClick(Sender);  // Ouverture
             if (gszFonction = 'SERIE') then TToolBarButton97(GetControl('BMODIF')).OnClick(Sender);  // Ouverture
           end;
      end;
    VK_F11, VK_RIGHT :
      begin
        if (ctxPCL in V_PGI.PGIContexte) and (gszFonction='C') then
        begin
          if ( GetNomEcran = 'MULSECTION') and GetParamSocSecur('SO_CPPCLSAISIETVA',false) and ( GetField('S_AXE') = GetParamSocSecur('SO_CPPCLAXETVA', '') ) then exit ;
          if V_PGI.ExtendedFieldSelection = '1' then
            POPMENUF11 := TPopUpMenu(getcontrol('POPMENUF11AXE'))
          else
           POPMENUF11 := TPopUpMenu(getcontrol('POPMENUF11'));
          if  (POPMENUF11 <> nil) then
            if POPMENUF11.items.Count >  0 then
              // POPMENUF11.Popup (+100,+100);
              POPMENUF11.Popup (Mouse.CursorPos.X,Mouse.CursorPos.Y);
        end;
      end;
    VK_DELETE :
      begin
        key := 0;
        if ((Shift = [ssCtrl]) and (TToolbarButton97(GetControl('BSUPPRIMER')).Enabled)
              and (TToolbarButton97(GetControl('BSUPPRIMER')).Visible)) then
          TToolbarButton97(GetControl('BSUPPRIMER')).Click;
      end;
    else fOnSaveKeyDownEcran(Sender, Key, Shift );
  end;
end;

procedure TOF_MULPARAMGEN.OnclickParam(Sender: TObject);
var
  sz : String;
  sTable,sCrit,sTri,sTriOld,sChamp,sLarg,sAlign,sParams,sPerso : String;
  wTitre,wLeTitre,wNomCol : WideString;
  bOkTri,bOkNumCol : boolean;
  order : string;
begin
inherited;
  { FQ 19681 BVE 26.06.07 }
  if GetControl('XX_ORDERBY') <> nil then
  begin
     ChargeHListeUser(TFMul(Ecran).DBListe,
                      V_PGI.User,
                      sTable,
                      sCrit,
                      sTriOld,
                      sChamp,
                      wTitre,
                      sLarg,
                      sAlign,
                      sParams,
                      wLeTitre,
                      wNomCol,
                      sPerso,
                      bOkTri,
                      bOkNumCol);
  end;
  sz := V_PGI.ExtendedFieldSelection;
  V_PGI.ExtendedFieldSelection := '';
  TFMul(Ecran).BParamListeClick(Self);
  V_PGI.ExtendedFieldSelection := sz;
  if GetControl('XX_ORDERBY') <> nil then
  begin
     ChargeHListeUser(TFMul(Ecran).DBListe,
                      V_PGI.User,
                      sTable,
                      sCrit,
                      sTri,
                      sChamp,
                      wTitre,
                      sLarg,
                      sAlign,
                      sParams,
                      wLeTitre,
                      wNomCol,
                      sPerso,
                      bOkTri,
                      bOkNumCol);
     order := GetControlText('XX_ORDERBY');
     if order = '' then
        // Il n'y a aucun données ecrasement sans probleme.
        SetControlText('XX_ORDERBY',sTri)
     else if order = sTriOld then
        // L'ancienne valeur est l'ancien Order BY on remplace.
        SetControlText('XX_ORDERBY',sTri)
     else if Pos(sTri,order) < 1 then
     begin
        // Il y a des données dans le order à conserver
        if Pos(sTriOld,order) > 0 then
           // Il y a l'ancien order à supprimer
           System.Delete(order,Pos(sTriOld,order),Length(sTriOld));
        SetControlText('XX_ORDERBY',order + ',' + sTri);
     end;
  end;
  { END FQ 19681 }
  if EstTablePartagee( getTableName ) then
    BChercheClick( nil ) ;


end;

procedure TOF_MULPARAMGEN.OnClickVentilable(Sender: TObject);
Var ComboVentil : THValComboBox ;
    lStWhere    : String ;
begin
  ComboVentil := THValComboBox( GetControl('VENTILABLE',True) ) ;
  InitXX_Where ;
  if ComboVentil.ItemIndex <= 0 then Exit ;

  lStWhere := GetControlText('XX_WHERE') ;
  if lStWhere <> '' then
    lStWhere := lStwhere + ' AND ' ;

  if ComboVentil.ItemIndex < 6 then
    lStWhere := lStwhere + 'G_VENTILABLE' + IntToStr(ComboVentil.ItemIndex) + '="X"'
  else
    if ComboVentil.ItemIndex = 6 then
      lStWhere := lStwhere + 'G_VENTILABLE1="-"'
                  + ' AND G_VENTILABLE2="-"'
                  + ' AND G_VENTILABLE3="-"'
                  + ' AND G_VENTILABLE4="-"'
                  + ' AND G_VENTILABLE5="-"'
    else
      if ComboVentil.ItemIndex = 7 then
        lStWhere:=lStWhere + ' (G_VENTILABLE1="X"'
                           + ' OR G_VENTILABLE2="X"'
                           + ' OR G_VENTILABLE3="X"'
                           + ' OR G_VENTILABLE4="X"'
                           + ' OR G_VENTILABLE5="X")' ;

  SetControltext( 'XX_WHERE' , lStWhere ) ;
end;

procedure TOF_MULPARAMGEN.InitXX_Where ( vStWhere : string = '' ) ;
Var lStWhere : String ;
begin
  SetControlText('XX_WHERE', '' ) ;

  if Ecran=nil then Exit ;

  if vStWhere <> '' then
   begin
    SetControlText('XX_WHERE', vStWhere ) ;
    exit ;
   end ; // if

  // Condition sur le champs xx_FERME
  if (gszFonction = 'F') or (gszFonction = 'O') then
    begin
    // Table ?
    if getNomEcran = 'MULGENERAUX'
      then lStWhere := ' G_FERME ='
      else if getNomEcran = 'MULSECTION'
        then lStWhere := ' S_FERME ='
        else if getNomEcran = 'MULTIERS'
          then lStWhere := ' T_FERME ='
          else if getNomEcran = 'MULJOURNAL'
            then lStWhere := ' J_FERME =' ;
    // Ouvert / Ferme ?
    if (gszFonction = 'O')
      then lStWhere := lStWhere + '"X" '
      else lStWhere := lStWhere + '"-" ';
    end ;

    //Modifie en serie des comptes (ana)    //SG6 05/01/05 FQ 15213
    if gszFonction='SERIEANA' then
    begin
      //On prend pas les comptes utilisés pour des écritures multi échéance
      // lStWhere := 'NOT (EXISTS(SELECT E_GENERAL FROM ECRITURE WHERE E_GENERAL=G_GENERAL AND E_NUMECHE>1))'
    end;

  // FQ 13717
  if getNomEcran = 'MULTIERS' then
   begin
    if not (lStWhere='') then lStWhere := lStWhere + ' AND ';
    lStWhere := lStWhere+'(T_NATUREAUXI="CLI" OR T_NATUREAUXI="FOU" OR T_NATUREAUXI="DIV" OR T_NATUREAUXI="AUC" OR T_NATUREAUXI="AUD" OR T_NATUREAUXI="SAL")';
   end;
  SetControltext( 'XX_WHERE' , lStWhere ) ;
end;


procedure TOF_MULPARAMGEN.OnClose;
begin
  V_PGI.ExtendedFieldSelection := '' ;
  V_PGI.enableTableToView      := False ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/07/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_MULPARAMGEN.OnAfterShow;
begin
{$IFDEF CCSTD}
  TFMul(Ecran).ListeFiltre.ForceAccessibilite := FaPublic;
{$ENDIF}
  FEnChargement := false ;
end;
////////////////////////////////////////////////////////////////////////////////



{$IFNDEF COMPTAPAIE}
{$IFNDEF GCGC}
//$IFNDEF CCADM}

procedure TOF_MULPARAMGEN.FermeCpteOnClick(Sender: TObject);
var lStFB : String ;
    lInForm  : Longint ;
begin
  if not TToolbarButton97(GetControl('BFermeCPTE1')).Enabled then exit ;
  lInForm  := Longint( Ecran ) ;
  lStFB := GetFBName ;
  AglOuvrFermListCpte( [ lInForm, lStFB, 'FALSE' ] , 0 );
  BChercheClick(nil) ;
end;

procedure TOF_MULPARAMGEN.ImprCpteOnClick(Sender: TObject);
var lStTable : String ;
    lInForm  : Longint ;
begin
  lInForm  := Longint( Ecran ) ;
  lStTable := GetTableName ;
  AglEditCompte ( [ lInForm, lStTable ] , 0 );
end;

procedure TOF_MULPARAMGEN.OuvreCpteOnClick(Sender: TObject);
var lStFB : String ;
    lInForm  : Longint ;
begin
  if not TToolbarButton97(GetControl('BOuvrCPTE1')).Enabled then exit ;

  if ( gszFonction = 'M' ) then
   begin
    TFMul(Ecran).Retour := GetField('T_AUXILIAIRE') ;
    SendMessage(Ecran.Handle, WM_CLOSE, 0, 0) ;
    exit ;
   end ;

  lInForm  := Longint( Ecran ) ;
  lStFB := GetFBName ;
  AglOuvrFermListCpte( [ lInForm, lStFB, 'TRUE' ] , 0 );
  BChercheClick(nil) ;
end;

procedure TOF_MULPARAMGEN.SupprCpteOnClick(Sender: TObject);
var lStTable : String ;
    lInForm  : Longint ;
begin
  if not TToolbarButton97(GetControl('BSUPPRIMER')).Enabled then exit ;
  lInForm  := Longint( Ecran ) ;
  lStTable := GetTableName ;
  AGLSupprimeListCpte( [ lInForm, lStTable ] , 0 );
  BChercheClick(nil) ;
end;

procedure TOF_MULPARAMGEN.InsertCpteOnClick(Sender: TObject);
var lStTable : String ;
begin
  lStTable := GetTableName ;
  AglCreationCompte ( [ lStTable ] , 0 );
  BChercheClick(nil) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_MULPARAMGEN.AxeChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  {JP 19/08/05 : FQ 15617 : Affectation de la tablette afin de permettre à l'Agl de prendre en
                 charge les caractères joker : Pour les sections il faut filtrer sur l'axe}
  if GetNomEcran = 'MULSECTION' then begin
    if GetControlText('S_AXE') <> '' then
      SetControlProperty('S_SECTION', 'PLUS', 'S_AXE = "' + GetControlText('S_AXE') + '"')
    else
      SetControlProperty('S_SECTION', 'PLUS', '');
  end;
end;

procedure TOF_MULPARAMGEN.ModifCpteOnClick(Sender: TObject);
var lStAxe    : String ;
    lStCpt    : String ;
    lStTable  : String ;
    lStAction : String ;
    lInForm   : Longint ;
begin
  lInForm   := Longint( Ecran ) ;
  lStAxe    := GetAxe ;
  lStCpt    := GetField(GetFieldName) ;
  lStTable  := GetTableName ;
  lStAction := GetTypeAction ;
  if lStAxe <> ''
    then AglModifCompte( [ lInForm, lStCpt, 'Modif', lStTable, lStAxe, lStAction ] , 0 )
    else AglModifCompte( [ lInForm, lStCpt, 'Modif', lStTable, lStAction ] , 0 ) ;
  BChercheClick(nil) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 04/01/2005
Modifié le ... : 04/01/2005
Description .. : - FB 15167 - LG - 04/01/2004 - Les concepts applicatifs de 
Suite ........ : la comptabilité m'interdisent de faire des modifications sur les
Suite ........ : comptes généraux mais il est possible d'utiliser la fonction de 
Suite ........ : modification en série.
Mots clefs ... : 
*****************************************************************}
procedure TOF_MULPARAMGEN.ModifSerieOnClick(Sender: TObject);
var lStAxe   : String ;
    lStCpt   : String ;
    lStTable : String ;
    lInForm  : Longint ;
begin
  if (GetNomEcran = 'MULGENERAUX') and Not ExJaiLeDroitConcept(TConcept(ccGenModif),True) then Exit ;
  if (GetNomEcran = 'MULJOURNAL') and Not ExJaiLeDroitConcept(TConcept(ccJalModif),True) then Exit ;
  if (GetNomEcran = 'MULTIERS') and Not ExJaiLeDroitConcept(TConcept(ccAuxModif),True) then Exit ;
  lInForm  := Longint( Ecran ) ;
  lStAxe   := GetAxe ;
  lStCpt   := GetField(GetFieldName) ;
  lStTable := GetTableName ;
  if lStAxe <> ''
    then AglModifCompte( [ lInForm, lStCpt, 'Serie', lStTable, lStAxe ] ,  0 )
    else AglModifCompte( [ lInForm, lStCpt, 'Serie', lStTable ]        ,  0 ) ;
end;

//$ENDIF CCADM}
{$ENDIF COMPTAPAIE}
{$ENDIF GCGC}


procedure TOF_MULPARAMGEN.InitEvenement;
  {$IFNDEF GCGC}
var MenuPopUp : TPopUpMenu ;
    i         : Integer ;
    CbAxe     : THValComboBox; {JP 19/08/05 : FQ 15617}
  {$ENDIF !GCGC}
begin

  // PARAM LISTE
  TToolBarButton97(GetControl ('BParamListe')).Onclick := OnclickParam;

  FOnBChercheClick := TToolBarButton97(GetControl('BCherche')).OnClick  ;
  TToolBarButton97(GetControl ('BCherche')).OnClick := BChercheClick ;

  // Evt sur combo ventilable
  if GetControl('VENTILABLE') <> nil then
    begin
    THValComboBox( GetControl('VENTILABLE') ).OnChange  := OnClickVentilable ;
    THValComboBox( GetControl('VENTILABLE') ).ItemIndex := 0 ;
    end ;

  // FORM KEY DOWN
  fOnSaveKeyDownEcran   := Ecran.OnKeyDown;
  Ecran.OnKeyDown       := FormKeyDown;

  // AFTER FORM SHOW
  TFmul(Ecran).OnAfterFormShow := OnAfterShow;

{$IFDEF COMPTAPAIE}
  FListe.OnDblClick := AfficheRibCompte;
{$ENDIF}

{$IFNDEF COMPTAPAIE}
{$IFNDEF GCGC}
//$IFNDEF CCADM}

  // Boutons communs
  if GetControl('BInsert')<>nil then
    TToolBarButton97( GetControl('BInsert') ).OnClick    := InsertCpteOnClick ;
  if GetControl('BSupprimer')<>nil then
    TToolBarButton97( GetControl('BSupprimer') ).OnClick := SupprCpteOnClick ;

  TToolBarButton97(GetCOntrol('BSELECTALL', True)).OnClick := BSelectAllClick;


  // Double click sur la liste
  if gszFonction = 'M' then
   FListe.OnDblClick := OuvreCpteOnClick
   else
    FListe.OnDblClick := ModifCpteOnClick ;



  if GetControl('BMODIF')<>nil then
    TToolBarButton97( GetControl('BMODIF') ).OnClick      := ModifSerieOnClick ;
  if GetControl('BOuvrCPTE1')<>nil then
    TToolBarButton97( GetControl('BOuvrCPTE1') ).OnClick  := OuvreCpteOnClick ;
  if GetControl('BFermeCPTE1')<>nil then
    TToolBarButton97( GetControl('BFermeCPTE1') ).OnClick := FermeCpteOnClick ;
  if GetControl('BImprimer')<>nil then
    TToolBarButton97( GetControl('BImprimer') ).OnClick   := ImprCpteOnClick ;
  if GetControl('BOuvrir')<>nil then
    TToolBarButton97( GetControl('BOuvrir') ).OnClick     := ModifCpteOnClick ;

  //==========================
  //===      MODE PCL      ===
  //==========================
  if (ctxPCL in V_PGI.PGIContexte) then
    begin
    // mode PCL popup Fermer / ouvrir
    MenuPopUp      := TPopUpMenu(GetControl('POPUPMENU1', True));
    for i := 0 to MenuPopUp.Items.Count -1 do
      begin
      if MenuPopUp.Items[i].Name = 'BFermeCpte'
        then MenuPopUp.Items[i].OnClick := FermeCpteOnClick
      else if MenuPopUp.Items[i].Name = 'BOuvrCpte'
        then MenuPopUp.Items[i].OnClick := OuvreCpteOnClick;
      end;

    // mode PCL F11
    MenuPopUp      := TPopUpMenu(GetControl('POPMENUF11', True));
    for i := 0 to MenuPopUp.Items.Count -1 do
      begin
      if UpperCase(MenuPopUp.Items[i].Name) = 'NOU'
        then MenuPopUp.Items[i].OnClick := InsertCpteOnClick
      else if UpperCase(MenuPopUp.Items[i].Name) = 'SUP'
        then MenuPopUp.Items[i].OnClick := SupprCpteOnClick
      else if UpperCase(MenuPopUp.Items[i].Name) = 'MOD'
        then MenuPopUp.Items[i].OnClick := ModifSerieOnClick
      else if UpperCase(MenuPopUp.Items[i].Name) = 'BFERM' then
        begin
          if GetNomEcran = 'MULJOURNAL' then
            MenuPopUp.Items[i].Caption := TraduireMemoire('&Fermer journal');
          MenuPopUp.Items[i].OnClick := FermeCpteOnClick;
        end
      else if UpperCase(MenuPopUp.Items[i].Name) = 'BOUV' then
      begin
        if GetNomEcran = 'MULJOURNAL' then
          MenuPopUp.Items[i].Caption := TraduireMemoire('&Ouvrir journal');
        MenuPopUp.Items[i].OnClick := OuvreCpteOnClick ;
      end;
      end;
  end
  else
  //==========================
  //===      MODE PGE      ===
  //==========================
  begin
    if (gszFonction = 'F') then
    begin
    if GetControl('BACTION')<>nil then
      TToolBarButton97( GetControl('BACTION') ).OnClick := FermeCpteOnClick ;
    end
    // Ouverture
    else if (gszFonction = 'O') then
      begin
      if GetControl('BACTION')<>nil then
        TToolBarButton97( GetControl('BACTION') ).OnClick  := OuvreCpteOnClick ;
      end ;
  end ;

  {JP 19/08/05 : FQ 15617 : Affectation de la tablette afin de permettre à l'Agl de prendre en
                 charge les caractères joker : Pour les sections il faut filtrer sur l'axe}
  if GetNomEcran = 'MULSECTION' then begin
    cbAxe := THValComboBox(GetControl('S_AXE'));
    if Assigned(cbAxe) then cbAxe.OnChange := AxeChange;
  end;

//$ENDIF CCADM}
{$ENDIF COMPTAPAIE}
{$ENDIF GCGC}
end;


{$IFNDEF GCGC}
function TOF_MULPARAMGEN.getFBName: String;
begin
  result := '' ;
  if getNomEcran = 'MULGENERAUX'
    then result := 'fbGene'
    else if getNomEcran = 'MULSECTION'
      then result := 'fbSect'
      else if getNomEcran = 'MULTIERS'
        then result := 'fbAux'
        else if getNomEcran = 'MULJOURNAL'
          then result := 'fbJal' ;
end;
{$ENDIF GCGC}

function TOF_MULPARAMGEN.getTableName: String;
begin
  result := '' ;
  if getNomEcran = 'MULGENERAUX'
    then result := 'GENERAUX'
    else if getNomEcran = 'MULSECTION'
      then result := 'SECTION'
      else if getNomEcran = 'MULTIERS'
        then result := 'TIERS'
        else if getNomEcran = 'MULJOURNAL'
          then result := 'JOURNAL' ;
end;

{$IFNDEF GCGC}
function TOF_MULPARAMGEN.getFieldName: String;
begin
  result := '' ;
  if getNomEcran = 'MULGENERAUX'
    then result := 'G_GENERAL'
    else if getNomEcran = 'MULSECTION'
      then result := 'S_SECTION'
      else if getNomEcran = 'MULTIERS'
        then result := 'T_AUXILIAIRE'
        else if getNomEcran = 'MULJOURNAL'
          then result := 'J_JOURNAL' ;
end;

function TOF_MULPARAMGEN.getAxe: String;
begin
  result := '' ;
  if getNomEcran = 'MULSECTION' then
    result := GetField('S_AXE') ;
end;
{$ENDIF GCGC}

procedure TOF_MULPARAMGEN.InitPresentation;
begin
  {$IFDEF EAGLCLIENT}
  TFMul(Ecran).FListe.SortEnabled := False;
  {$ELSE}
  TFMul(Ecran).FListe.SortEnabled := True;
  {$ENDIF}

  // =============================
  // ===  ONGLET TABLES LIBRES ===
  // =============================
  // MAJ de l'onglet table libre pour généraux
  if GetControl('PLIBREG') <> nil then
    LibellesTableLibre( TTabSheet(GetControl('PLIBREG')) ,'TG_TABLE','G_TABLE','G') ;
  // MAJ de l'onglet table libre pour auxiliaires
  if GetControl('PLIBRET') <> nil then
    LibellesTableLibre( TTabSheet(GetControl('PLIBRET')) ,'TT_TABLE','T_TABLE','T') ;
  // MAJ de l'onglet table libre pour sections
  if GetControl('PLIBRES') <> nil then
    LibellesTableLibre( TTabSheet(GetControl('PLIBRES')) ,'TS_TABLE','S_TABLE','S') ;


  // ==========================
  // ===  PRESENTATION PCL  ===
  // ==========================
  if (ctxPCL in V_PGI.PGIContexte) then
  begin
    // Bouton tout sélectionner visible en mode PCL
    // FQ 20183 SetControlVisible ('BSELECTALL',True);
    SetVisibleBSelectAll(true);
    // Modif En Série des paramètres analytiques des comptes généraux.
    if (gszFonction = 'SERIEANA') then
    begin
      SetControlVisible ( 'BMODIF',     True );
      // FQ 20183 SetControlVisible ( 'BSELECTALL', True );
      SetVisibleBSelectAll(true);
      Ecran.Caption := TraduireMemoire('Modification en série des paramètres analytiques des ')+LowerCase(Ecran.Caption) ;
    end
    else if (gszFonction <> 'C' ) then
    begin
    SetControlVisible ('BSUPPRIMER',  False);
    SetControlVisible ('BOUVRE',      False);
    SetControlVisible ('BMODIF',      False);
    end;
  end
  else
  // ==========================
  // ===  PRESENTATION PGE  ===
  // ==========================
  begin
    // on est pas en mode PCL ... cacher les boutons inutiles !
    SetControlVisible ('BSUPPRIMER',  False);
    SetControlVisible ('BOUVRE',      False);
    SetControlVisible ('BMODIF',      False);
    //SG6 13/01/05 Bouton nouveau dans le mul FQ 15225
    //SetControlVisible ('BINSERT',     False);
    SetControlVisible ('BFERMECPTE1', False);
    SetControlVisible ('BOUVRCPTE1',  False);

    {$IFDEF COMPTA}
    {JP 04/08/05 : FQ 15900 (et 12151) : c'est deux zones n'ont, a priori, rien à voir avec la comptabilité}
    if getNomEcran = 'MULSECTION' then begin
      SetControlVisible ('S_CHANTIER', False);
      SetControlVisible ('S_MAITREOEUVRE', False);
      SetControlVisible ('TS_CHANTIER', False);
      SetControlVisible ('TS_MAITREOEUVRE', False);
    end;
    {$ENDIF COMPTA}

    // Suppression
    if (gszFonction = 'S') then
    begin
      SetControlVisible ('BSUPPRIMER', True);
      SetControlProperty('BSUPPRIMER', 'HINT', TraduireMemoire('Lancer la suppression'));
      Ecran.Caption := TraduireMemoire('Suppression des ')+LowerCase(Ecran.Caption);
      (GetControl('BSUPPRIMER') as TToolbarButton97).GlobalIndexImage := 'Z0003_S16G2';
    end
    // Fermeture
    else if (gszFonction = 'F') then
    begin
      SetControlVisible ('BACTION', True );
      SetControlProperty('BACTION', 'HINT', TraduireMemoire('Lancer le traitement'));
      Ecran.Caption := TraduireMemoire('Fermeture des ')+LowerCase(Ecran.Caption);
    end
    // Ouverture
    else if (gszFonction = 'O') then
    begin
      SetControlVisible ('BACTION', True );
      SetControlProperty('BACTION', 'HINT', TraduireMemoire('Lancer le traitement'));
      Ecran.Caption := TraduireMemoire('Ouverture des ')+LowerCase(Ecran.Caption);
    end
    // Modif En Série
    else if (gszFonction = 'SERIE') then
    begin
      SetControlVisible ( 'BMODIF',     True );
      // FQ 20183 SetControlVisible ( 'BSELECTALL', True );
      SetVisibleBSelectAll(true);
      Ecran.Caption := TraduireMemoire('Modification en série des ')+LowerCase(Ecran.Caption);
    end
    // Modif En Série des paramètres analytiques des comptes généraux.
    else if (gszFonction = 'SERIEANA') then
    begin
      SetControlVisible ( 'BMODIF',     True );
      // FQ 20183 SetControlVisible ( 'BSELECTALL', True );
      SetVisibleBSelectAll(true);
      Ecran.Caption := TraduireMemoire('Modification en série des paramètres analytiques des ')+LowerCase(Ecran.Caption) ;
    end;

  // Gestion du mode multiselection de la grille
  if not ( ( gszFonction = 'S' )    or
           ( gszFonction = 'F' )    or
           ( gszFonction = 'O' )    or
           ( gszFonction = 'SERIE') or
           ( gszFonction = 'SERIEANA')
          ) then
    begin
   {$IFDEF EAGLCLIENT}
    TFMul(Ecran).FListe.MultiSelect := False;
   {$ELSE}
    TFMul(Ecran).FListe.MultiSelection := False;
   {$ENDIF}
    end;

  end ;

 // FQ 20183 SetControlVisible ( 'BSELECTALL', ( gszFonction <> 'M' ) ) ;
 SetVisibleBSelectAll( gszFonction <> 'M' );
 SetControlVisible ( 'BOuvrCPTE1', ( gszFonction = 'M' ) ) ;

end;

{$IFNDEF GCGC}
function TOF_MULPARAMGEN.getTypeAction: String;
begin
  if ( gszFonction = 'C' )
    then result := 'ACTION=MODIFICATION'
    else result := 'ACTION=CONSULTATION' ;
end;
{$ENDIF !GCGC}

function TOF_MULPARAMGEN.getNomEcran: String;
begin
  result := '' ;
  if pos( 'MULGENERAUX', Ecran.Name ) > 0
    then result := 'MULGENERAUX'
    else if pos( 'MULSECTION', Ecran.Name ) > 0
      then result := 'MULSECTION'
      else if pos( 'MULTIERS', Ecran.Name ) > 0
        then result := 'MULTIERS'
        else if pos( 'MULJOURNAL', Ecran.Name ) > 0
          then result := 'MULJOURNAL' ;
end;

////////////////////////////////////////////////////////////////////////////////
{$IFDEF COMPTAPAIE}
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/10/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_MULPARAMGEN.AfficheRibCompte(Sender: TObject);
var lStCpt : string;
begin
  lStCpt := GetField(GetFieldName);
  if ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL = "' + lStCpt + '" AND ' +
               'G_NATUREGENE = "BQE" ORDER BY G_GENERAL') then
  begin
    if not ExisteSQL('SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_GENERAL="' + lStCpt
                    +'" AND BQ_NODOSSIER = "'+V_PGI.NoDossier+'"') then // 19/10/2006 YMO Multisociétés
      FicheBanqueCP(lStCpt, TaCreat, 0) 
    else
      FicheBanqueCP(lStCpt, TaModif, 0) ;
  end;
end;
{$ENDIF}
////////////////////////////////////////////////////////////////////////////////

procedure TOF_MULPARAMGEN.DetermineListe;
begin

  // Spécif pour modif en série sur analytique des comptes généraux
  if gszFonction = 'SERIEANA'
    then TFMul(Ecran).Q.Liste := 'MULMGENEMODIFSANA'
  {b FP 29/12/2005}
  else if (gszFonction = 'SERIE') and (GetNomEcran = 'MULGENERAUX') then
    begin
    TFMul(Ecran).DBListe := 'CVGENERAUX';
    TFMul(Ecran).Q.Liste := 'CVGENERAUX';
    end;
  {e FP 29/12/2005}

end;

procedure TOF_MULPARAMGEN.BChercheClick(sender: TObject);
{$IFNDEF EAGLCLIENT}
var
  sTable,sCrit,sTri,sChamp,sLarg,sAlign,sParams,sPerso : String;
  wTitre,wLeTitre,wNomCol : WideString;
  bOkTri,bOkNumCol : boolean;
  order : string;
{$ENDIF}  
begin
inherited;
  { FQ 19681 BVE 26.06.07 }
{$IFNDEF EAGLCLIENT}
  if GetControl('XX_ORDERBY') <> nil then
  begin
     ChargeHListeUser(TFMul(Ecran).DBListe,
                      V_PGI.User,
                      sTable,
                      sCrit,
                      sTri,
                      sChamp,
                      wTitre,
                      sLarg,
                      sAlign,
                      sParams,
                      wLeTitre,
                      wNomCol,
                      sPerso,
                      bOkTri,
                      bOkNumCol);
     order := GetControlText('XX_ORDERBY');
     if order = '' then
        // Il n'y a aucun données ecrasement sans probleme.
        SetControlText('XX_ORDERBY',sTri)
     else if Pos(sTri,order) < 1 then
     begin
        // Il y a des données dans le order à conserver
        SetControlText('XX_ORDERBY',order + ',' + sTri);
     end;
  end;
{$ENDIF}  
  { END FQ 19681 }
  
  // Gestion multisociete : recherche sur la vue associée dans DESHARE si table partagée
  if EstTablePartagee( getTableName ) then
    v_pgi.enableTableToView := True ;

  if Assigned ( FOnBChercheClick ) then
    FOnBChercheClick( Sender ) ;

  // Gestion multisociete : désactivation de la vue
  v_pgi.enableTableToView := False ;

  {$IFNDEF EAGLCLIENT}
  if ( gszFonction = 'M' ) and FBoToucheCtrl then
   FListe.Columns[1].Title.Font.Style := FListe.Columns[1].Title.Font.Style + [fsUnderline] ;
  {$ENDIF} 

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/07/2007
Modifié le ... :   /  /    
Description .. : GCO - 17/07/2007 - FQ 20714
Mots clefs ... :
*****************************************************************}
{$IFNDEF GCGC}
procedure TOF_MULPARAMGEN.BSelectAllClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  if not FListe.AllSelected then
  begin
    if not TFMul(Ecran).FetchLesTous then
    begin
      PGIInfo('Sélection annulée. Impossible de récupérer tous les enregistrements.');
      Exit;
    end;
  end;
{$ENDIF}

  TFMUl(Ecran).bSelectAllClick( nil );
end;
{$ENDIF !GCGC}

////////////////////////////////////////////////////////////////////////////////

procedure TOF_MULPARAMGEN.OnDisplay ;
var
 lBoAxeTva : boolean ;
begin
 inherited ;

 if ( GetNomEcran = 'MULSECTION') then
 begin
   if GetParamSocSecur('SO_CPPCLSAISIETVA',false) then
    begin
     lBoAxeTva := GetField('S_AXE') = GetParamSocSecur('SO_CPPCLAXETVA', '') ;
     SetControlEnabled('BSUPPRIMER'    , not lBoAxeTva) ;
     SetControlEnabled('BOUVRCPTE1'    , not lBoAxeTva) ;
     SetControlEnabled('BFERMECPTE1'   , not lBoAxeTva) ;
     SetControlEnabled('BMODIF'        , not lBoAxeTva) ;
     SetControlEnabled('BOUVRE'        , not lBoAxeTva) ;
    end ; // if
 end;

 if ( gszFonction = 'M' ) and FListe.CanFocus then
  FListe.SetFocus ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 15/05/2007
Modifié le ... : 15/05/2007
Description .. : On grise le bouton Tout selectionner si les bouton BMODIF
Suite ........ : et BOUVRE le sont aussi.
Suite ........ : FQ 20183
Mots clefs ... : 
*****************************************************************}
procedure TOF_MULPARAMGEN.SetVisibleBSelectAll( condition : boolean ) ;
begin
{$IFDEF EAGLCLIENT}
  condition := condition and ( GetControlVisible('BOUVRE') or GetControlVisible('BMODIF') );
{$ENDIF}
  SetControlVisible('bSelectAll',condition);
end;

Initialization
RegisterClasses([TOF_MULPARAMGEN]);

end.



