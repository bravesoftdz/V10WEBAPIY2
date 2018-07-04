{ Unité : Source TOM de la TABLE : RUBRIQUE (RUBRIQUE)
--------------------------------------------------------------------------------------
    Version   |   Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
 4.3.1.101.001  19/06/03   JP   Création de l'unité
 FUSION         30/01/04   JP   Xavier MG a utilisé RUBRIQUE.PAS pour la version espagnole
                                J'ai essayer d'intégrer les modifications de XMG :
                                rechercher "XMG-FUSION" (Fiche Decla modifiée en fonction)
 1.0.5.001.002  19/03/04   JP   Suppression de edRubriqueExit et mise du code dans
                                OnChangeField, car edRubriqueExit ne fonctionne pas
                                correctement en eAgl
 1.90.001.001   10/06/04   JP   TRESO : Avant de supprimer une rubrique on s'assure qu'elle
                                n'est pas utilisée dans la table TREcriture.
 6.0x.xxx.xxx   19/07/04   JP   Gestion des commissions de mouvements pour la tréso
 6.01.016.001   27/09/04   JP   Gestion des mode de règlements dans les rubriques de Tréso FQ 10119
 6.20.001.002   02/02/05   VL   Suppression Conseils et avertissement + Modif pour les budgets
 6.50.001.021   03/10/05   JP   Gestion du type de rubrique en Treso pour la version CWAS
 7.01.001.003   10/08/06   JP   On cache les contrôles RB_DATEVALIDITE en TRESO
 7.05.001.004   21/11/06   CA   FQ 19152 : pas de contrôle du premier caractère en enregistrement dans un dossier
 8.00.001.015   15/05/07   JP   FQ 20340 : Problème de focus sur la grille
--------------------------------------------------------------------------------------}

unit RUBRIQUE_TOM ;

Interface

Uses
  StdCtrls,
  Controls,
  Classes,
  Messages,
  {$IFDEF EAGLCLIENT}
  MaineAGL, eFiche, eFichList, eMul, UtileAGL,
  {$ELSE}
  FE_Main, Fiche, FichList, Mul, db, HDB, dbCtrls,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  EdtREtat,
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOM, AGLInit, HTB97, Math,
  UTob, {$IFNDEF TRESO}CpteUtil, {$ENDIF} Ent1, Grids, Windows,
  Graphics,
  ParamSoc, // GetParamSocSecur
  {$IFDEF MODENT1}
  CPTypeCons, 
  {$ENDIF MODENT1}
  choix;

Type

  TRubInterval = record
                   Borne1Deb : string;
                   Borne1Fin : string;
                   Borne2Deb : string;
                   Borne2Fin : string;
                   AvecExclusion : Boolean;
                 end ;

  Type TCtxSurRubrique = ( CtxRubrique, CtxBudget, CtxCycle, CtxLiasse) ;

  TOM_RUBRIQUE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  private
    ListFam  : TListBox;
    Msg      : THMsgBox;
//    TobRub   : TOB;
    G        : THGrid;
    {$IFDEF EAGLCLIENT}
    cbType   : THValComboBox;
    cbClasse : THValComboBox;
    cbSigne  : THValComboBox;
    cbAxe    : THValComboBox;
    cbBudget : THValComboBox;
    {$ELSE}
    cbType   : THDBValComboBox;
    cbClasse : THDBValComboBox;
    cbSigne  : THDBValComboBox;
    cbAxe    : THDBValComboBox;
    cbBudget : THDBValComboBox;
    {$ENDIF}

    FCtxSurRubrique : TCtxSurRubrique;

    Depasse,
    InsertOk   : Boolean;
    CarBour,
    CarBour1,
    CarBourBud : Char ;
    MemoTyp,
    MemoSig,
    MemoFam,
    MemoAx,
    MemoTl,
    MemoJal    : String ;
    MemoNbCol  : Byte ;
    {$IFDEF TRESO}
    tCodeFlux  : TStringList;
    cbModePaie : THMultiValComboBox;
    {$ENDIF}
    Formulerubrique : TMemo;
//    LaRubrique      : THEdit;

    FStFamille : string; // Sauvegarde de la famille au chargement du record
    fbPersoStandard : boolean;

    procedure GCpte1DblClick    (Sender: TObject);
    procedure BFamClick         (Sender: TObject);
    procedure BCopierClick      (Sender: TObject);
    procedure BImprimerClick    (Sender: TObject);
    procedure RB_TYPERUBChange  (Sender: TObject);
    procedure RB_CLASSERUBChange(Sender: TObject);
    procedure FCalculChange     (Sender: TObject);
    procedure FamBudChange      (Sender: TObject);
    procedure GCpte1CellExit    (Sender: TObject; var ACol,ARow: Longint; var Cancel: Boolean);
    procedure GCpte1KeyPress    (Sender: TObject; var Key: Char);
    procedure GCpte1KeyDown     (Sender: TObject; var Key: Word; Shift: TShiftState);
    {$IFDEF TRESO}
    procedure RB_SIGNERUBChange (Sender: TObject);
    procedure MODEPAIEChange    (Sender: TObject);
    procedure GetModePaiement;
    procedure SetModePaiement;
    procedure OnArgumentTreso;
    {$ELSE}
    procedure RempliComboFamRub(C : THValComboBox);
    procedure ChargeListeFam;
    {$ENDIF}
//    procedure BRubriquesClick   (Sender : TObject); {XMG-FUSION}

    procedure LitCpteExclu     (St : String ; Col : Byte);
    procedure ComboAxe         (Avec : Boolean) ;
    procedure TronqueLeCompte  (var Chaine : string);
    procedure TransFormeCode   (var Chaine : string ; PVirgule : Boolean);

    procedure NewEnreg;
    procedure EffaceLeGrid;
    procedure RetailLeGrid;
    procedure ChargeEnreg;
    procedure QuelCarBour;
    procedure AjouteLigne;
    procedure EffaceLigne;
    procedure AfficheLibelle;
    procedure PositionneFamille;
    procedure InitMsg ;
    procedure FermeMsg ;
    procedure MemoChampParDefaut;
    procedure ChampsParDefaut;
    procedure EcritCpteExclu;
    procedure MetRequeteEnEdit;
    procedure OnAfterFormShow; {JP 15/05/07 : FQ 20340}

    function  AfficheMsg(num : Integer; Av, Ap : string ) : Word ;
    function  Grid2St   (ACol : Integer )      : string ;
    function  QuelEstLefb                      : TFichierBase;
    function  ChampsParDefautPourCreatenSerie  : Boolean;
    function  CpteVide                         : Boolean;
    function  EnregOK                          : Boolean;
    function  ControleCpteAbrOk                : Boolean;
    function  ControleCodeRubriqueOk           : Boolean;
    function  AnalyseCompteOk                  : Boolean;
    function  BudJalOk                         : Boolean;
  end ;

function ParametrageRubrique(LaRub : String ; Action : TActionFiche ; vCtxSurRubrique : TCtxSurRubrique; InsertOk : Boolean =False; bPersoStandard : boolean = False ): string ;
procedure FicheRubrique(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
//procedure ParametrageRubrique2(LaRub : String ; Action : TActionFiche; Budget : Boolean ) ;

function  ControlIntervalle( vStRubrique, vStCompte1, vStExclu1, vStFamilles : string ) : Boolean;

implementation

uses
  {$IFDEF MODENT1}
  CPVersion,
  {$ENDIF MODENT1}
  {$IFDEF eAGLCLIENT}
    MenuOLX,
  {$ELSE}
    MenuOLG,
  {$ENDIF eAGLCLIENT}
  {$IFNDEF CCS3}
    HZoomSp,
  {$ENDIF}

  {$IFNDEF TRESO}
    CPQRTRUBRIQUE_TOF,
    FAMRUB_TOF ,
    CPMULRUB_TOF,
  {$ENDIF}

  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}

  Lookup,
  UtilPGI,
  TabLiRub,
  GalOutil, // NoDossierBaseCommune
  utilSoc;  // AffecteCycleRevisionAuxGene

var
  Resultat : string;

{---------------------------------------------------------------------------------------}
function ParametrageRubrique(LaRub : String ; Action : TActionFiche ; vCtxSurRubrique : TCtxSurRubrique; InsertOk : Boolean =False; bPersoStandard : boolean = False ): string ;
{---------------------------------------------------------------------------------------}
var
  ch,
  lStCtx : string;
  lStEnSerie : string;
  lStInsert  : string;
  lStPersoStandard : string;
{$IFNDEF TRESO}
  QLoc   : TQuery ;
{$ENDIF TRESO}
begin
  {$IFDEF TRESO}
  {$ELSE}
  if (vCtxSurRubrique = CtxRubrique) then
  begin
    QLoc := OpenSql('Select YDS_CODE, YDS_LIBELLE From CHOIXDPSTD Where YDS_TYPE = "RBF" Order by YDS_CODE ',True);
    if QLoc.Eof then
    begin
      Ferme(QLoc) ;
      HShowMessage('0;Rubriques;ATTENTION ! Vous devez d''abord paramétrer des familles.;E;O;O;O;','','');
      Exit ;
    end ;
    Ferme(QLoc) ;
  end ;
  {$ENDIF TRESO}

  // GCO - 30/05/2005 - TaModifEnSerie non géré par la fonction ActionToString
  // donc géré en deuxième paramètre du OnArgument dans CPMULRUB_TOF
  lStEnSerie := 'NON';

{$IFDEF EAGLCLIENT}
{$ELSE}
  if Action = TaModifEnSerie then
  begin
    Action := TaModif;
    lStEnSerie := 'OUI';
  end;
{$ENDIF}

  // Contexte de fonctionnement
  if vCtxSurRubrique = CtxBudget then
    lStCtx := 'BUDGET'
  else
    lStCTx := 'RUBRIQUE';

  // Autorisation de création ?? ( voir logique VL )
  if InsertOk then
    lStInsert := 'OUI'
  else
    lStInsert := 'NON';

  // Personnalisation d'un standard
  if bPersoStandard then
    lStPersoStandard := 'OUI'
  else
    lStPersoStandard := 'NON';

  ch := ActionToString(Action) + ';' + lStEnSerie + ';' + lStCtx + ';' + lStInsert + ';'+ lStPersoStandard + ';';

  AGLLanceFiche('CP','CPFICRUBRIQUE','', LaRub, ch);

  Result := Resultat;
end;

{---------------------------------------------------------------------------------------}
procedure FicheRubrique(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
{---------------------------------------------------------------------------------------}
var
  ch, Ok : string;
begin
  if QuellePage = 0 then Ok := 'OUI'
                    else Ok := 'NON';

  ch := ActionToString(Comment) + ';NON;RUBRIQUE;NON;';
  AGLLanceFiche('CP','CPFICRUBRIQUE','', Compte, ch);
end;

(*
{---------------------------------------------------------------------------------------}
procedure ParametrageRubrique2(LaRub : String ; Action : TActionFiche ; Budget : Boolean ) ;
{---------------------------------------------------------------------------------------}
var
  QLoc : TQuery ;
  ch,
  B2   : string;
begin
  if not Budget then begin
    QLoc := OpenSql('Select YDS_CODE,YDS_LIBELLE From CHOIXDPSTD Where YDS_TYPE="RBF" Order by YDS_CODE ',True) ;
    if QLoc.Eof then begin
      Ferme(QLoc) ;
      HShowMessage('0;Rubriques;ATTENTION ! Vous devez d''abord paramétrer des familles;E;O;O;O;','','') ;
      Exit ;
    end ;
    Ferme(QLoc) ;
  end ;

  if Budget then B2 := 'OUI'
            else B2 := 'NON';

  ch := ActionToString(Action) + ';OUI;' + B2 + ';NON;';
  AGLLanceFiche('CP','CPFICRUBRIQUE','', LaRub, ch);

  LaRub := Resultat;
end ;
*)

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/04/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function  ControlIntervalle( vStRubrique, vStCompte1, vStExclu1, vStFamilles : string ) : Boolean;
var lRubInterval1     : TRubInterval;
    lRubinterval2     : TRubInterval;
    lStListeCompte    : string;
    lStListeExclusion : string;
    lStCompte         : string;
    lStExclusion      : string;
    lQuery            : TQuery;
    lTobListeLigne1   : Tob;
    lTobListeLigne2   : Tob;
    lTob              : Tob;
    i,j               : integer;

    function _FaitWhereFamille ( stListeFamille : string ) : string;
    var St : string;
    begin
      Result := '';
      St := ReadTokenSt ( stListeFamille );
      if St <> '' then Result := ' AND (';
      while St <> '' do
      begin
        Result := Result + ' RB_FAMILLES LIKE "%'+St+';%"';
        St := ReadTokenSt ( stListeFamille );
        if St <> '' then Result := Result + ' OR ';
      end;
      if Result <> '' then Result := Result + ')';
    end;

    procedure _RemplitRubInterval( var vInterval : TRubInterval; vStCompte, vStExclusion : string);
    var lBoIntervalEntier : Boolean;
        lSt : string;
        lInPos1 : integer;
        lInPos2 : integer;
    begin
      FillChar( vInterval, sizeOf( vInterval), #0);

      // Remplissage des bornes extérieures
      lBoIntervalEntier := Pos(':', vStCompte) > 0;
      if lBoIntervalEntier then
      begin
        lInPos1 := Pos(':', vStCompte);
        lInPos2 := Pos('(SM)', vStCompte);
        lSt     := Copy( vStCompte, 1, lInPos1 - 1 );
        vInterval.Borne1Deb := BourreEtLess( lSt, fbGene); // OK

        lSt := Copy( vStCompte, lInPos1 + 1, (lInPos2) - (lInPos1 + 1));
        vInterval.Borne2Fin := BourreLaDoncSurLesComptes( lSt, '9');
      end
      else
      begin
        lInPos1 := Pos('(SM)', vStCompte);
        lSt     := Copy( vStCompte, 1, lInPos1 - 1);
        vInterval.Borne1Deb := BourreEtLess( lSt, fbGene);
        vInterval.Borne2Fin := BourreLaDoncSurLesComptes( lSt, '9');
      end;

      // Remplissage des bornes intérieures
      if vStExclusion = '' then
      begin
        vInterval.AvecExclusion := False;
        vInterval.Borne1Fin := vInterval.Borne2Fin;
        vInterval.Borne2Deb := vInterval.Borne1Deb;
      end
      else
      begin
        // Rajout d'un ; pour marquer la fin de la chaine
        vStExclusion := vStExclusion + ';';
        vInterval.AvecExclusion := True;
        lBoIntervalEntier := Pos(':', vStExclusion) > 0;
        if lBoIntervalEntier then
        begin
          lInPos1 := Pos(':', vStExclusion);
          lInPos2 := Pos(';', vStExclusion);

          lSt := Copy( vStExclusion, 1, lInPos1 - 1);
          vInterval.Borne1Fin := IntToStr( StrToInt( lst ) - 1);
          vInterval.Borne1Fin := BourreLaDoncSurLesComptes( vInterval.Borne1Fin, '9');

          lSt := Copy( vStExclusion, lInPos1 + 1, (lInPos2) - (lInPos1 + 1));
          if Length(lSt) = VH^.Cpta[fbGene].lg then
            vInterval.Borne2Deb := lSt
          else
          begin
            vInterval.Borne2Deb := IntToStr( StrToInt ( lSt ) + 1);
            vInterval.Borne2Deb := BourreEtLess( vInterval.Borne2Deb, fbGene);
          end;
        end
        else
        begin
          lSt := Copy( vStExclusion, 1, Pos(';', vStExclusion)-1);
          vInterval.Borne1Fin := IntToStr( StrToInt( lSt )- 1);
          vInterval.Borne1Fin := BourreLaDoncSurLesComptes( vInterval.Borne1Fin, '9');

          vInterval.Borne2Deb := IntToStr( StrToInt( lSt ) + 1);
          vInterval.Borne2Deb := BourreEtLess( vInterval.Borne2Deb, fbGene);
        end;
      end;
    end;

    function _ChevauchementRubInterval : Boolean;
    begin
      Result := False;

      // GCO - 28/04/2005 - Pour les IF, y avait moyen d'en faire un seul mais
      // en cas de bug c'est plus simple à trouver qui renvoie TRUE
      if ((lRubInterval2.Borne1Deb >= lRubInterval1.Borne1Deb) and
          (lRubInterval2.Borne1Deb <= lRubInterval1.Borne1Fin)) or
          ((lRubInterval1.Borne1Deb >= lRubInterval2.Borne1Deb) and
          (lRubInterval1.Borne1Deb <= lRubInterval2.Borne1Fin)) then
      begin
        Result := True;
        Exit;
      end;

      if ((lRubInterval2.Borne1Fin >= lRubInterval1.Borne1Deb) and
          (lRubInterval2.Borne1Fin <= lRubInterval1.Borne1Fin)) or
          ((lRubInterval1.Borne1Fin >= lRubInterval2.Borne1Deb) and
          (lRubInterval1.Borne1Fin <= lRubInterval2.Borne1Fin)) then
      begin
        Result := True;
        Exit;
      end;

      if ((lRubInterval2.Borne1Deb >= lRubInterval1.Borne2Deb) and
          (lRubInterval2.Borne1Deb <= lRubInterval1.Borne2Fin)) or
          ((lRubInterval1.Borne1Deb >= lRubInterval2.Borne2Deb) and
          (lRubInterval1.Borne1Deb <= lRubInterval2.Borne2Fin)) then
      begin
        Result := True;
        Exit;
      end;

      if ((lRubInterval2.Borne1Fin >= lRubInterval1.Borne2Deb) and
          (lRubInterval2.Borne1Fin <= lRubInterval1.Borne2Fin)) or
          ((lRubInterval1.Borne1Fin >= lRubInterval2.Borne2Deb) and
          (lRubInterval1.Borne1Fin <= lRubInterval2.Borne2Fin)) then
      begin
        Result := True;
        Exit;
      end;

      if ((lRubInterval2.Borne2Deb >= lRubInterval1.Borne1Deb) and
         (lRubInterval2.Borne2Deb <= lRubInterval1.Borne1Fin)) or
         ((lRubInterval1.Borne2Deb >= lRubInterval2.Borne1Deb) and
         (lRubInterval1.Borne2Deb <= lRubInterval2.Borne1Fin)) then
      begin
        Result := True;
        Exit;
      end;

      if ((lRubInterval2.Borne2Fin >= lRubInterval1.Borne1Deb) and
          (lRubInterval2.Borne2Fin <= lRubInterval1.Borne1Fin)) or
          ((lRubInterval1.Borne2Fin >= lRubInterval2.Borne1Deb) and
          (lRubInterval1.Borne2Fin <= lRubInterval2.Borne1Fin)) then
      begin
        Result := True;
        Exit;
      end;

      if ((lRubInterval2.Borne2Deb >= lRubInterval1.Borne2Deb) and
          (lRubInterval2.Borne2Deb <= lRubInterval1.Borne2Fin)) or
          ((lRubInterval1.Borne2Deb >= lRubInterval2.Borne2Deb) and
          (lRubInterval1.Borne2Deb <= lRubInterval2.Borne2Fin)) then
      begin
        Result := True;
        Exit;
      end;

      if ((lRubInterval2.Borne2Fin >= lRubInterval1.Borne2Deb) and
         (lRubInterval2.Borne2Fin <= lRubInterval1.Borne2Fin)) or
         ((lRubInterval1.Borne2Fin >= lRubInterval2.Borne2Deb) and
         (lRubInterval1.Borne2Fin <= lRubInterval2.Borne2Fin)) then
      begin
        Result := True;
        Exit;
      end;
    end;

begin
  Result := True;
  lQuery := nil;
  try
    try
      // Contient les lignes du Cycle révision que l'on sauve
      lTobListeLigne1   := Tob.Create('', nil, -1);
      lStListeCompte    := vStCompte1;
      lStListeExclusion := vStExclu1;

      while lStListeCompte <> '' do
      begin
        lStCompte    := ReadTokenSt( lStlisteCompte );
        lStExclusion := ReadTokenSt( lStlisteExclusion );

        if lStCompte <> '' then
        begin
          lTob := Tob.Create( '', lTobListeLigne1, -1);
          lTob.AddChampSupValeur('COMPTE1', lStCompte, False );
          lTob.AddChampSupValeur('EXCLUSION1', lStExclusion, False );
        end;
      end;

      // Contient les lignes des Cycles de révision dans la base
      lTobListeLigne2 := Tob.Create('', nil, -1);

      { CA - 04/10/2005 - Ajout critère famille dans le contrôle }
      lQuery := OpenSQL('SELECT RB_RUBRIQUE, RB_COMPTE1, RB_EXCLUSION1 FROM ' +
                        'RUBRIQUE WHERE RB_RUBRIQUE <> "' + vStRubrique + '" AND ' +
                        'RB_CLASSERUB = "CDR" '+_FaitWhereFamille ( vStFamilles )+' ORDER BY RB_RUBRIQUE', True);

      // Charge tout les intervalles de la table RUBRIQUE sauf celle concernée
      while not lQuery.Eof do
      begin
        lStListeCompte    := lQuery.FindField('RB_COMPTE1').AsString;
        lStListeExclusion := lQuery.FindField('RB_EXCLUSION1').AsString;

        while lStListeCompte <> '' do
        begin
          lStCompte    := ReadTokenSt( lStlisteCompte );
          lStExclusion := ReadTokenSt( lStlisteExclusion );

          if lStCompte <> '' then
          begin
            lTob := Tob.Create( '', lTobListeLigne2, -1);
            lTob.AddChampSupValeur('RUBRIQUE', lQuery.FindField('RB_RUBRIQUE').AsString);
            lTob.AddChampSupValeur('COMPTE1', lStCompte, False );
            lTob.AddChampSupValeur('EXCLUSION1', lStExclusion, False );
          end;
        end;
        lQuery.Next;
      end;

      for i := 0 to lTobListeLigne1.Detail.Count -1 do
      begin
        // On remplit la ligne[i] dy cycle de révision que l'on sauve
        _RemplitRubInterval( lRubInterval1,
                             lTobListeLigne1.Detail[i].GetString('COMPTE1'),
                             lTobListeLigne1.Detail[i].GetString('EXCLUSION1'));

        // Parcours des lignes de chaque ycle de révision que l'on doit controller
        for j := 0 to lTobListeLigne2.Detail.Count -1 do
        begin
          _RemplitRubInterval( lRubInterval2,
                               lTobListeLigne2.Detail[j].GetString('COMPTE1'),
                               lTobListeLigne2.Detail[j].GetString('EXCLUSION1'));

          if _ChevauchementRubInterval then
          begin
            PgiInfo('Erreur de fourchettes avec le cycle de révision ' +
                    lTobListeLigne2.Detail[j].GetString('RUBRIQUE') + '.',
                    'Contrôle des cycles de révision');
            Result := False;
            Break;
          end;
        end;
      end;

    except
      on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : ComprareInterval');
    end;

  finally
    Ferme( lQuery ) ;
    FreeAndNil( lTobListeLigne1 );
    FreeAndNil( lTobListeLigne2 );
  end;
end;

////////////////////////////////////////////////////////////////////////////////

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.OnNewRecord ;
{---------------------------------------------------------------------------------------}
var
  TRubrique : TOB;
  stSQL : string;
  lstRubrique : string;
  lstCode : string;
  lstPredefini : string;
  lstNoDossier : string;
  lDateValidite : TDateTime;
begin
  Inherited;
  {Initialisation des valeurs par défaut}
  NewEnreg;
  lstRubrique := TFFiche(Ecran).FLequel;
  {$IFDEF TRESO}
  SetControlVisible ('RB_DATEVALIDITE', False);
  SetControlVisible ('TRB_DATEVALIDITE', False);
  SetField('RB_CLASSERUB', 'TRE');
  {$ELSE}
  SetControlVisible ('RB_DATEVALIDITE',(FCtxSurRubrique = ctxRubrique));
  SetControlVisible ('TRB_DATEVALIDITE',(FCtxSurRubrique = ctxRubrique));
  {$ENDIF TRESO}
  { Création d'une rubrique à partir d'une rubrique CEGID ou STANDARD }
  if fbPersoStandard then
  begin
    lstCode := ReadTokenSt(lstRubrique);
    lstNoDossier := ReadTokenSt(lstRubrique);
    lstPredefini := ReadTokenSt(lstRubrique);
    lDateValidite := StrToDate(ReadTokenSt(lstRubrique));
    TFFiche(Ecran).FLequel := lstCode;

    SetControlEnabled ('RB_RUBRIQUE',False);
    SetControlEnabled ('RB_CODEABREGE',False);
    SetControlEnabled ('RB_CLASSERUB',False);
    SetControlEnabled ('RB_TYPERUB',False);
    SetControlEnabled ('BCOPIER',False);

    stSQL := 'SELECT * FROM RUBRIQUE WHERE RB_RUBRIQUE = "'+lstCode+
      '" AND RB_PREDEFINI="'+lstPredefini+'" AND RB_NODOSSIER="'+lstNodossier+
      '" AND RB_DATEVALIDITE="'+USDateTime(lDateValidite)+'"';
    TRubrique := TOB.Create ('', nil, -1);
    try
      // Chargement des informations la rubrique initiale
      TRubrique.LoadDetailFromSQL (stSQL);
      // Copie des informations de la rubrique initiale dans la nouvelle rubrique
      if TRubrique.Detail.Count >  0 then
      begin
        SetField('RB_RUBRIQUE', TRubrique.Detail[0].GetValue('RB_RUBRIQUE'));
        SetField('RB_CODEABREGE', TRubrique.Detail[0].GetValue('RB_CODEABREGE'));
        SetField('RB_LIBELLE', TRubrique.Detail[0].GetValue('RB_LIBELLE'));
        SetField('RB_FAMILLES', TRubrique.Detail[0].GetValue('RB_FAMILLES'));
        SetField('RB_SIGNERUB', TRubrique.Detail[0].GetValue('RB_SIGNERUB'));
        SetField('RB_TYPERUB', TRubrique.Detail[0].GetValue('RB_TYPERUB'));
        SetField('RB_COMPTE1', TRubrique.Detail[0].GetValue('RB_COMPTE1'));
        SetField('RB_EXCLUSION1', TRubrique.Detail[0].GetValue('RB_EXCLUSION1'));
        SetField('RB_COMPTE2', TRubrique.Detail[0].GetValue('RB_COMPTE2'));
        SetField('RB_EXCLUSION2', TRubrique.Detail[0].GetValue('RB_EXCLUSION2'));
        SetField('RB_AXE', TRubrique.Detail[0].GetValue('RB_AXE'));
        SetField('RB_NATRUB', TRubrique.Detail[0].GetValue('RB_NATRUB'));
        SetField('RB_TABLELIBRE', TRubrique.Detail[0].GetValue('RB_TABLELIBRE'));
        SetField('RB_BUDJAL', TRubrique.Detail[0].GetValue('RB_BUDJAL'));
        SetField('RB_CATEGORIE', TRubrique.Detail[0].GetValue('RB_CATEGORIE'));
        SetField('RB_RUBSYNTHESE', TRubrique.Detail[0].GetValue('RB_RUBSYNTHESE'));
        SetField('RB_CLASSERUB', TRubrique.Detail[0].GetValue('RB_CLASSERUB'));
      end;
    finally
      TRubrique.Free;
    end;
  end;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.OnDeleteRecord ;
{---------------------------------------------------------------------------------------}
begin
{$IFDEF TRESO}
   if ExisteSQL('SELECT TE_CODEFLUX FROM TRECRITURE WHERE TE_CODEFLUX = "' +GetControlText('RB_RUBRIQUE') + '"') then begin
     LastErrorMsg := TraduireMemoire('La rubrique est utilisée dans des écritures de trésorerie.'#13 +
                                     'La suppression est impossible.');
     LastError := 2;
   end;
{$ENDIF}

  ExecuteSql('DELETE FROM CUMULS WHERE CU_TYPE="RUB" AND CU_COMPTE1="' + GetControlText('RB_RUBRIQUE') + '"');
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.OnUpdateRecord ;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRESO}
  SetModePaiement; {29/09/04 : on stocke les modes de paiement dans RB_EXCLUSION2}
  {$ENDIF}
  inherited;
  {Si les contrôles et les mises à jour des champs ne sont pas corrects, on empêche le
  le update}
  if not EnregOK then LastError := -1;
end ;

{---------------------------------------------------------------------------------------}
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... : 08/09/2005
Description .. :
Suite ........ : GCO - 08/09/2005 - Mise à jour du champ YDS_LIBRE pour les
Suite ........ : plans de révision
Mots clefs ... :
*****************************************************************}
procedure TOM_RUBRIQUE.OnAfterUpdateRecord ;
{---------------------------------------------------------------------------------------}
var lStLiasseDossier : string;
begin
  Inherited ;

  if FCtxSurRubrique = ctxRubrique then
  begin
    lStLiasseDossier := GetParamSocSecur('SO_CPCONTROLELIASSE', '');
    if lStLiasseDossier <> '' then
    begin
      if (Pos(lStLiasseDossier, GetField('RB_FAMILLES')) > 0) or
         (Pos(lStLiasseDossier, FStFamille) = 0) then
      begin
        if VH^.BStatutLiasse <> nil then
          VH^.BStatutLiasse.ImageIndex := 3;
      end;
    end;
  end;

end ;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.OnLoadRecord ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  ChargeEnreg;
  {$IFDEF TRESO}
  GetModePaiement; {29/09/04 : on récupère les modes de paiement de RB_EXCLUSION2}
  SetControlVisible ('RB_DATEVALIDITE', False);
  SetControlVisible ('TRB_DATEVALIDITE', False);
  {$ELSE}
  SetControlVisible ('RB_DATEVALIDITE',(FCtxSurRubrique = ctxRubrique));
  SetControlVisible ('TRB_DATEVALIDITE',(FCtxSurRubrique = ctxRubrique));
  {$ENDIF}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.OnChangeField ( F: TField ) ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  {19/03/2004 : reprise du code de edRubriqueExit}
  if F.FieldName = 'RB_RUBRIQUE' then begin
    if GetControlText('RB_CODEABREGE') = '' then {JP 15/03/04}
      SetField('RB_CODEABREGE', GetField('RB_RUBRIQUE'))
  end;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.OnArgument ( S: String ) ;
{---------------------------------------------------------------------------------------}
var SQL : string;
    lStCtxSurRubrique : string;
begin
  inherited ;
  // Lecture d'ACTION=
  ReadTokenSt(S);

  // Lecture si ModifEnSerie
  ReadTokenSt(S);

  // Lecture du Contexte de fonctionnement
  lStCtxSurRubrique := ReadTokenSt(S);
  if lStCtxSurRubrique = 'BUDGET' then
    FCtxSurRubrique := CtxBudget
  else
    FCtxSurRubrique := CtxRubrique;

  // Autorisation insert ?? ( voir logique VL )
  InsertOk := ReadTokenSt(S) = 'OUI';

  // Personnalisation d'un standard
  fbPersoStandard := ReadTokenSt(S) = 'OUI';

  InitMsg;

  ListFam := TListBox(GetControl('LISTFAM'));
  TToolbarButton97(GetControl('BFAM')).OnClick :=BFamClick;

  // Invisible = TRUE pour les composants des BUDGET
  SetControlEnabled('TRB_BUDJAL', False);
  SetControlEnabled('RB_BUDJAL', False);

  if FCtxSurRubrique = CtxBudget then
  begin
    TFFiche(Ecran).HelpContext := 15417000 ;
    SetControlVisible('BFAM', False );
    SetControlEnabled('TRB_BUDJAL', True);
    SetControlEnabled('RB_BUDJAL', (TFFiche(Ecran).FTypeAction in [taCreat..taCreatOne]));
    {b FP 15/11/2005 FQ13268}
    if (TFFiche(Ecran).FTypeAction in [taCreat..taCreatOne]) then
      PGIInfo('Nous vous conseillons de passer par la fonction Automates de création des rubriques bugetaires.');
    {e FP 15/11/2005 FQ13268}
  end
  else
  begin
    if ctxTreso in V_PGI.PGIContexte then Ecran.HelpContext := 50000142
                                     else Ecran.HelpContext := 7784000;
  end;

  {CODE NON PORTÉ CAR "ParametrageRubrique2" NE SEMBLE JAMAIS UTILISÉ

  if OnlyGrid then  begin
    GbParam.Visible:=False ;     GroupBox1.Visible:=False ;
    RB_TABLELIBRE.Visible:=False ;
    BFirst.Visible:=False ; BPrev.Visible:=False ; BNext.Visible:=False ; BLast.Visible:=False ;
    BAnnuler.Visible:=False ; BInsert.Visible:=False ; BDelete.Visible:=false ;
    XX:=GCpte1.Left-3 ;
    BValider.Left:=BValider.Left-XX ;
    BAide.Left:=BAide.Left-XX ;
    BImprimer.Left:=BImprimer.Left-XX ;
    BFerme.Left:=BFerme.Left-XX ;
    GCpte1.Left:=0 ; GCpte1.Top:=0 ; Width:=GCpte1.Width+18 ; GCpte1.Align:=alClient ;
    GCpte1.SetFocus ;
  end ;}

{$IFDEF EAGLCLIENT}
  cbType   := THValComboBox(GetControl('RB_TYPERUB'));
  cbClasse := THValComboBox(GetControl('RB_CLASSERUB'));
  cbAxe    := THValComboBox(GetControl('RB_AXE'));
  cbSigne  := THValComboBox(GetControl('RB_SIGNERUB'));
  cbBudget := THValComboBox(GetControl('RB_BUDJAL'));
{$ELSE}
  cbType   := THDBValComboBox(GetControl('RB_TYPERUB'));
  cbClasse := THDBValComboBox(GetControl('RB_CLASSERUB'));
  cbAxe    := THDBValComboBox(GetControl('RB_AXE'));
  cbSigne  := THDBValComboBox(GetControl('RB_SIGNERUB'));
  cbBudget := THDBValComboBox(GetControl('RB_BUDJAL'));
{$ENDIF}
  FormuleRubrique:=nil ;
  if VH^.PaysLocalisation=CodeISOES then
     Formulerubrique:=TMemo(GetControl('Formulerubrique')) ; //XVI 24/02/2005

  cbType  .OnChange := RB_TYPERUBChange;
  cbClasse.OnChange := RB_CLASSERUBChange;
  THValComboBox(GetControl('FCALCUL')).OnChange := FCalculChange;
  THValComboBox(GetControl('FAMBUD' )).OnChange := FamBudChange;

  if ((EstSerie(S3)) or (EstSerie(S5))) then
  begin
    SetControlChecked('RB_TABLELIBRE', False);
    SetControlVisible('RB_TABLELIBRE', False);
  end ;

  CarBourBud := VH^.BourreLibre;
//  TobRub := TOB.Create('', nil, -1);

  SQL := 'SELECT * FROM RUBRIQUE';
{$IFDEF TRESO}
  {JP 03/10/05 : En eAgl, au moment d'écrire dans la table rubrique, l'agl regarde si la
                 valeur du champ correspond à une valeur qui existe dans la tablette : dans
                 le cas contraire, le champ est vidé => pour débrancher le test, je supprime
                 la tablette attachée au champ}
  cbType.DataType := '';
  cbType.Values.Clear;
  cbType.Values.Add('TRP');
  cbType.Values.Add('TRN');

  {RB_TYPERUB : TR pour Tréso, P (Positif), N (Négatif)}
  SQL := SQL + ' WHERE RB_NATRUB = "TRE" AND (RB_TYPERUB = "TRP" OR RB_TYPERUB = "TRN")';
{$ELSE}
  if FCtxSurRubrique = CtxBudget then
  begin
    SQL := SQL + ' WHERE (RB_NATRUB="BUD")';
    RempliComboFamRub(THValComboBox(GetControl('FAMBUD')));
  end
  else
  begin
    FactoriseComboTypeRub(cbType);
    SQL := SQL + ' WHERE (RB_NATRUB="CPT")';
  end ;
{$ENDIF}
  if VH^.PaysLocalisation=CodeISOES then
     SQL:=SQL+' and (not RB_RUBRIQUE like "@%")' ; //XVI 24/02/2005

(*  Qr := OpenSql(SQL, True);
  TobRub.LoadDetailDB('RUBRIQUE', '', '', Qr, True);
  Ferme(Qr);

  if TobRub.Detail.Count > 0 then begin
  if (TFFiche(Ecran).FLequel <> '') and not (TFFiche(Ecran).FTypeAction in [taCreat..taCreatOne]) then
    if (TobRub.FindFirst(['RB_RUBRIQUE'], [TFFiche(Ecran).FLequel], True) = nil) or
       (not TobRub.FindFirst(['RB_RUBRIQUE'], [TFFiche(Ecran).FLequel], True).Detail.Count = 0) then begin
      MessageAlerte(Msg.Mess[6]) ;
      PostQuitMessage(1);
      Exit;
    end;
    end;
*)
  MemoNbCol := 0;

  G := THGrid(GetControl('GCPTE1'));
  G.OnCellExit := GCpte1CellExit;
  G.OnKeyPress := GCpte1KeyPress;
  G.OnKeyDown  := GCpte1KeyDown;
  G.OnDblClick := GCpte1DblClick;

  {JP 17/02/04 : Avec les modifications de la comptabilité espagnole, la grille n'est plus à alClient}
  if VH^.PaysLocalisation<>CodeISOES then
     G.Align := alClient;

  TToolbarButton97(GetControl('BCOPIER')  ).OnClick := BCopierClick;
  TToolbarButton97(GetControl('BIMPRIMER')).OnClick := BImprimerClick;

  case TFFiche(Ecran).FTypeAction of
       taConsult           : begin
                               SetControlEnabled('GCPTE1', True);
                               SetControlEnabled('FCALCUL', False);
                               G.Options := G.Options - [goEditing] + [goRowSelect] ;
                               Exit ;
                             end ;
       taCreat..taCreatOne : begin
                               SetControlEnabled('BANNULER', False);
                              // SetControlText('RB_RUBRIQUE', TFFiche(Ecran).FLequel);
                             end ;
  end ;

  SetControlVisible ('RB_DATEVALIDITE',False);
  SetControlVisible ('TRB_DATEVALIDITE',False);  

  {$IFDEF CCS3}
  SetControlVisible('RB_AXE', False);
  SetControlVisible('TRB_AXE', False);
  SetControlVisible('RB_BUDJAL', False);
  SetControlVisible('TRB_BUDJAL', False);
  {$ENDIF}

  {$IFDEF TRESO}
  OnArgumentTreso;
  {$ENDIF}
  TFFiche(Ecran).OnAfterFormShow := OnAfterFormShow; {JP 15/05/07 : FQ 20340}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.OnClose ;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRESO}
  FreeAndNil(tCodeFlux);
  {$ENDIF}
//  FreeAndNil(TobRub);
  Resultat := GetControlText('RB_RUBRIQUE');
  G.VidePile(False);
  FermeMsg;
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.OnCancelRecord ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.BFamClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Famille, LeLibFam : String ;
begin
  if ListFam.CanFocus then
    ListFam.SetFocus ;
  if Trim(GetControlText('RB_RUBRIQUE')) = '' then begin
    AfficheMsg(2, '', '');
    Exit ;
  end ;

  Famille := GetField('RB_FAMILLES');

{$IFNDEF TRESO}
  ParametrageFamilleRubrique(GetControlText('RB_RUBRIQUE'), GetControlText('RB_LIBELLE'), Famille, LeLibFam, False) ;
{$ENDIF}

  if Famille <> GetField('RB_FAMILLES') then
  begin
    if DS.State = dsBrowse then ForceUpdate;

    if Famille <> '' then SetField('RB_FAMILLES', Famille);

    ListFam.Items.Clear ;
    while LeLibFam <> '' do ListFam.Items.Add(ReadTokenSt(LeLibFam)) ;
  end ;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.GCpte1CellExit(Sender: TObject; var ACol,ARow: Longint; var Cancel: Boolean);
{---------------------------------------------------------------------------------------}
{$IFNDEF TRESO}
var
  FCalcul : THValComboBox;
{$ENDIF}
begin
  if TFFiche(Ecran).FTypeAction = taConsult then Exit ;
  {$IFDEF TRESO}
  if (G.Cells[0,ARow] <> '') then
    AjouteLigne ;
  {$ELSE}
  FCalcul := THValComboBox(GetControl('FCALCUL'));
  if ((G.Cells[2,ARow] = '') and (G.Cells[ACol,ARow] <> '')) then begin
    FCalcul.Value := 'SM' ;
    G.Cells[2,ARow] := FCalcul.Text ;
    AjouteLigne ;
  end ;
  {$ENDIF}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.MetRequeteEnEdit ;
{---------------------------------------------------------------------------------------}
begin
  if not (DS.State in [dsEdit,dsInsert]) then begin
    DS.Edit ;
    G.SetFocus ;
  end ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.GCpte1KeyPress(Sender: TObject; var Key: Char);
{---------------------------------------------------------------------------------------}
begin
  if TFFiche(Ecran).FTypeAction = taConsult then Exit ;

  MetRequeteEnEdit ;
  if (Key = ';') or (GetField('RB_TABLELIBRE') = 'X') or (cbClasse.Value = 'RDR') then begin
    Key:=#0 ;
    Exit ;
  end ;
  Key := UpCase(Key) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.GCpte1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
{---------------------------------------------------------------------------------------}
Var
  i  : Byte ;
  St : string ;
begin
  if TFFiche(Ecran).FTypeAction = taConsult then Exit ;

  if (Key = VK_DELETE) and not (DS.Modified) then
    if not (DS.State in [dsEdit,dsInsert]) then DS.Edit ;

  if not(ssCtrl in Shift) then Exit ;

  if Key = VK_DELETE then begin
    for i := 0 to G.ColCount - 1 do G.Cells[i,G.Row] := '' ;
    EffaceLigne ;
    if G.Cells[0, G.RowCount - 1] <> '' then G.RowCount := G.RowCount + 1;
  end ;

  {$IFNDEF CCS3}
  if (Key = VK_F5) and (Pos('SECT',UpperCase(G.Cells[G.Col, 0])) > 0) then
  begin
    St := G.Cells[G.Col,G.Row] ;
    // GCO - 29/11/2006 - FQ 19175
    if ExisteSQL('SELECT SS_AXE FROM STRUCRSE WHERE SS_AXE = "' + cbAxe.Value + '"') then
    begin
      if ChoisirSousPlan(AxeToFb(cbAxe.Value), St, True, taModif) then
      begin
        if G.Cells[G.Col,G.Row] = '' then
          G.Cells[G.Col, G.Row] := St
        else
          G.Cells[G.Col, G.Row] := G.Cells[G.Col,G.Row] + St;
      end;
    end;  
  end;
  {$ENDIF}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.FCalculChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if DS.State = dsBrowse then DS.Edit ;
end;

{Copie la rubrique courante et place les données en mode nouveau
{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.BCopierClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var St : String ;
    Q : TQuery ;
    lStWhere : string;
begin
  {On s'assure que le code de la rubrique n'est pas vide}
  if Trim(GetControlText('RB_RUBRIQUE')) = '' then
  begin
    AfficheMsg(18,'','') ;
    SetFocusControl('RB_RUBRIQUE');
    Exit ;
  end;

  {Choix de la rubrique à dupliquer}
  lStwhere := 'RB_RUBRIQUE <> "' + GetControlText('RB_RUBRIQUE') + '"';

  St := Choisir(Msg.Mess[19], 'RUBRIQUE', 'RB_RUBRIQUE || " / " || RB_LIBELLE', 'RB_RUBRIQUE', lStwhere, 'RB_RUBRIQUE');
  if St = '' then Exit;

  {Confirmation de la duplication}
  if AfficheMsg(20, '', '') <> mrYes then Exit ;

  {Changement l'état si nécessaire}
  if not (TFFiche(Ecran).FTypeAction in [taCreat, taCreatOne, taCreatEnSerie]) then
    TFFiche(Ecran).FTypeAction := taCreat;

  //NewRecord;
  //DS.Insert;
  Q := OpenSQL('SELECT * FROM RUBRIQUE WHERE RB_RUBRIQUE="' + St + '"', TRUE) ;

  if not Q.Eof then
  begin
    //SetField('RB_RUBRIQUE'  , '');
    //SetField('RB_CODEABREGE', '');

    SetField('RB_FAMILLES',    Q.FindField('RB_FAMILLES'   ).AsString);
    SetField('RB_LIBELLE',     Q.FindField('RB_LIBELLE'    ).AsString);
    SetField('RB_SIGNERUB',    Q.FindField('RB_SIGNERUB'   ).AsString);
    SetField('RB_TYPERUB',     Q.FindField('RB_TYPERUB'    ).AsString);
    SetField('RB_AXE',         Q.FindField('RB_AXE'        ).AsString);
    SetField('RB_COMPTE1',     Q.FindField('RB_COMPTE1'    ).AsString);
    SetField('RB_COMPTE2',     Q.FindField('RB_COMPTE2'    ).AsString);
    SetField('RB_NATRUB',      Q.FindField('RB_NATRUB'     ).AsString);
    SetField('RB_TABLELIBRE',  Q.FindField('RB_TABLELIBRE' ).AsString);
    SetField('RB_BUDJAL',      Q.FindField('RB_BUDJAL'     ).AsString);
    SetField('RB_CATEGORIE',   Q.FindField('RB_CATEGORIE'  ).AsString);
    SetField('RB_EXCLUSION1',  Q.FindField('RB_EXCLUSION1' ).AsString);
    SetField('RB_EXCLUSION2',  Q.FindField('RB_EXCLUSION2' ).AsString);
    SetField('RB_RUBSYNTHESE', Q.FindField('RB_RUBSYNTHESE').AsString);
    SetField('RB_CLASSERUB',   Q.FindField('RB_CLASSERUB'  ).AsString);
  end ;
  Ferme(Q) ;
  ChargeEnreg;
  SetControlEnabled('RB_RUBRIQUE', True);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.BImprimerClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
{$IFNDEF TRESO}
var
  MaTob : TOB;
{$ENDIF}
begin
{$IFDEF TRESO}
{$ELSE}
  TOF_CPQRTRUBRIQUE.RemplirTob(MaTob, 'SELECT * FROM RUBRIQUE WHERE RB_RUBRIQUE = "' + GetControlText('RB_RUBRIQUE') + '"', True, FCtxSurRubrique = CtxBudget);
  LanceEtatTob('E', 'RUB', 'FRU', MaTob, True, False, False, nil, '', 'Rubrique : ' + GetControlText('RB_RUBRIQUE'), False, 0);
  FreeAndNil(MaTob);
{$ENDIF}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.RB_TYPERUBChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  CB   : THValComboBox;
begin
  {$IFDEF TRESO}
  Exit;
  {$ENDIF TRESO}

  CB := THValComboBox(GetControl('CBCAPTION'));
  G.Cells[2,0] := CB.Items[8] ;
  RetailLeGrid ;
  AfficheLibelle;
  if (cbType.Value = 'ANA') or (cbType.Value = 'A/B') or (cbType.Value = 'A/G')or
     (cbType.Value = 'B/A') or (cbType.Value = 'G/A') then begin
    ComboAxe(True);
    SetField('RB_AXE', 'A1');
  end
  else begin
    ComboAxe(False);
    SetField('RB_AXE', '');
  end ;

  QuelCarBour ;
  PositionneFamille ;
  SetControlEnabled('RB_TABLELIBRE', Pos('/', cbType.Value) <= 0) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.RB_CLASSERUBChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
//XMG 05/04/04 début
var lBoRubdeRub : Boolean ;
begin
  if cbClasse.Value = 'TLI' then SetField('RB_TABLELIBRE', 'X')
                            else SetField('RB_TABLELIBRE', '-');
  if VH^.PaysLocalisation=CodeISOES then
  Begin
     lBoRubdeRub:=GetField('RB_CLASSERUB')='RDR' ;
     SetControlVisible('FormuleRubrique', lBoRubDeRub) ;
     SetControlVisible('BRubriques',      lBoRubDeRub) ;
     SetControlVisible('GCpte1',          not lBoRubDeRub) ;
     SetControlEnabled('RB_TYPERUB',      not lBoRubDeRub) ;
     SetControlEnabled('RB_AXE',          Not lBoRubDeRub) ;
     SetControlEnabled('RB_BUDJAL',       (FCtxSurRubrique = CtxBudget) and (not lBoRubDeRub)) ;
  End ; //XVI 24/02/2005
end;
//XMG 05/04/04 fin
{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.FamBudChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  {$IFNDEF TRESO}
  ChargeListeFam;
  {$ENDIF}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.GCpte1DblClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
Var
  Lefb   : TFichierBase ;
  CodDe,
  CodA,
  Stock,
  St    : String ;
  i, j  : Integer ;
begin
  if TFFiche(Ecran).FTypeAction = taConsult then Exit;

  if (G.Col <> 0) and ((G.Col <> 1) or (cbClasse.value<>'RDR')) then Exit ;

  if GetField('RB_TABLELIBRE') = 'X' then begin
    MetRequeteEnEdit ;
    Lefb := QuelEstLefb ;
    if G.Cells[0,G.Row] <> '' then begin
      i := Pos(':',G.Cells[0,G.Row]);

      if i > 0 then begin
        CodDe := Copy(G.Cells[0, G.Row], 1, i - 1);
        CodA  := Copy(G.Cells[0, G.Row], i + 1, Length(CodDe)) ;
        Stock := Copy(G.Cells[0, G.Row], 2 * Length(CodDe) + 2 , Length(G.Cells[0, G.Row])) ;
        TransformeCode(CodDe, True);
        TransformeCode(CodA , True);
      end

      else begin
        i := Pos('&', G.Cells[0,G.Row]);
        j := Pos('|', G.Cells[0,G.Row]);
        if (i = 0) and (j = 0) then CodDe := G.Cells[0, G.Row] ;

        if (i<> 0) and (j = 0) then begin
          CodDe := Copy(G.Cells[0, G.Row], 1, i - 2) ;
          Stock := Copy(G.Cells[0, G.Row], i, Length(G.Cells[0, G.Row]));
        end;

        if (j<> 0) and (i=0) then begin
          CodDe := Copy(G.Cells[0, G.Row], 1, j - 2) ;
          Stock := Copy(G.Cells[0, G.Row], j, Length(G.Cells[0,G.Row]));
        end;

        if (i <> 0) and (j <> 0)then begin
          CodDe := Copy(G.Cells[0,G.Row], 1, Min(i,j)-2) ;
          Stock := Copy(G.Cells[0,G.Row], Min(i,j), Length(G.Cells[0, G.Row]));
        end;
        TransformeCode(CodDe,True) ;
        CodA := CodDe ;
       end;
    end;
    if (Stock <> '') and (Stock[1] = ',') then
      Stock := Copy(Stock, 2, Length(Stock));

    ChoixTableLibrePourRub(Lefb, cbAxe.Value, Stock, CodDe, CodA);

    TransformeCode(CodDe,False) ;
    TransformeCode(CodA, False) ;

    if CodA <> CodDe then St := CodDe + ':' + CodA
                     else St := CodDe;

    if Stock <> '' then begin
      if Stock[1] <> ',' then Stock := ',' + Stock;
      St := St + Stock;
    end;
    G.Cells[0, G.Row] := St ;
  end ;

  {$IFNDEF TRESO}
  if (cbClasse.Value = 'RDR') then begin
    MetRequeteEnEdit ;
    CodDe := G.Cells[G.Col, G.Row] ;
    MultiCritereRubriqueV3(cbType.Value, cbAxe.Value, '', cbBudget.Text, CodDe, False);
    G.Cells[G.Col,G.Row]:=CodDe;
  end;
  {$ENDIF}

  {$IFNDEF CCS3}
  if cbClasse.Value = 'CPT' then
  begin
    if (Pos('SECT',UpperCase(G.Cells[G.Col, 0])) > 0) and
    // GCO - 29/11/2006 - FQ 19175
    ExisteSQL('SELECT SS_AXE FROM STRUCRSE WHERE SS_AXE = "' + cbAxe.Value + '"') then
    begin
      St := G.Cells[G.Col,G.Row] ;
      if ChoisirSousPlan(AxeToFb(cbAxe.Value), St, True, taModif) then
      begin
        if G.Cells[G.Col,G.Row] = '' then
          G.Cells[G.Col, G.Row] := St
        else
          G.Cells[G.Col, G.Row] := G.Cells[G.Col,G.Row] + St;
      end;
    end;
  end;
  {$ENDIF}
end;

{---------------------------------------------------------------------------------------}
function TOM_RUBRIQUE.QuelEstLefb : TFichierBase;
{---------------------------------------------------------------------------------------}
begin
  Result := fbGene ;
  if FCtxSurRubrique = CtxBudget then
  begin
    case cbType.Value[1] of
      'A' : case cbAxe.Value[2] of
              '1' : Result := fbBudsec1;
              '2' : Result := fbBudsec2;
              '3' : Result := fbBudsec3;
              '4' : Result := fbBudsec4;
              '5' : Result := fbBudsec5;
            end;

       'G' : Result := fbBudgen;
    end ;
  end
  else
  begin
    case cbType.Value[1] of
      'A' : case cbAxe.Value[2] of
              '1' : Result := fbAxe1;
              '2' : Result := fbAxe2;
              '3' : Result := fbAxe3;
              '4' : Result := fbAxe4;
              '5' : Result := fbAxe5;
             end;

       'G' : Result := fbGene;
       'T' : Result := fbAux ;
    end ;
  end ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.NewEnreg;
{---------------------------------------------------------------------------------------}
var
  bStdCegid : Boolean;
begin
  {JP 09/09/2003 : Lorsque le New est appelé depuis la fiche, on est en taModif}
  if TFFiche(Ecran).FTypeAction = taModif then
    TFFiche(Ecran).FTypeAction := taCreat;

{$IFDEF TRESO}
  SetField('RB_NATRUB', 'TRE');
{$ELSE}
  if FCtxSurRubrique = CtxBudget then
    SetField('RB_NATRUB', 'BUD')
  else
    SetField('RB_NATRUB', 'CPT');
{$ENDIF}

  if ctxStandard in V_PGI.PGIContexte then
  begin
    bStdCegid := EstSpecif('51502');
    if ((bStdCegid) and (ctxStandard in V_PGI.PGIContexte)) then
      SetField('RB_PREDEFINI', 'CEG')
    else
      SetField('RB_PREDEFINI', 'STD');

    SetField('RB_NODOSSIER', NoDossierBaseCommune);
  end
  else
  begin
    SetField('RB_PREDEFINI', 'DOS');
    SetField('RB_NODOSSIER', V_PGI.NoDossier);
  end;

  { Date de validité }
  SetField('RB_DATEVALIDITE',iDate2099);

  {Autoriser la copie en mode création}
  SetControlEnabled('RB_RUBRIQUE', True);
  SetFocusControl('RB_RUBRIQUE');

  {$IFNDEF TRESO}
  // Pas d'accès à la duplication si personnalisation de standard
  SetControlEnabled('BCOPIER', not fbPersoStandard);
  {$ENDIF}

  if not ChampsParDefautPourCreatenSerie then ChampsPardefaut;
end ;

{---------------------------------------------------------------------------------------}
function TOM_RUBRIQUE.ChampsParDefautPourCreatenSerie : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result:=False ;

  if TFFiche(Ecran).FTypeAction <> taCreatenSerie then Exit ;
  if MemoTyp = '' then Exit ;
  SetField('RB_TABLELIBRE', MemoTl );
  SetField('RB_TYPERUB'   , MemoTyp);
  SetField('RB_BUDJAL'    , MemoJal);
  SetField('RB_TYPERUB'   , MemoTyp);
  SetField('RB_SIGNERUB'  , MemoSig);
  SetField('RB_FAMILLES'  , MemoFam);
  SetField('RB_AXE'       , MemoAx );
  G.RowCount := MemoNbCol ;
  EffaceLeGrid;
  {$IFNDEF TRESO}
  ComboAxe(MemoAx <> '') ;
  ChargeListeFam;
  {$ENDIF}
  Result := True ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.MemoChampParDefaut;
{---------------------------------------------------------------------------------------}
begin
  if TFFiche(Ecran).FTypeAction <> taCreatenSerie then Exit ;
  {$IFDEF TRESO}
  MemoTyp   := 'TRE';
  MemoSig   := 'POS';
  MemoFam   := '';
  MemoAx    := '';
  MemoTl    := GetField('RB_TABLELIBRE');
  MemoJal   := '';
  {$ELSE}
  MemoTyp   := cbType .Value ;
  MemoSig   := cbSigne.Value ;
  MemoFam   := GetField('RB_FAMILLES');
  MemoAx    := cbAxe.Value;
  MemoTl    := GetField('RB_TABLELIBRE');
  MemoJal   := cbBudget.Value;
  {$ENDIF}
  MemoNbCol := G.RowCount ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.ChargeEnreg;
{---------------------------------------------------------------------------------------}
{$IFNDEF TRESO}
var
  bStdCegid : Boolean;
  ch        : string;
{$ENDIF}
begin
  if (VH^.PaysLocalisation=CodeISOES) and (GetField('RB_CLASSERUB') = 'RDR') then //XVI 24/02/20005
     Formulerubrique.Lines.Text := GetField('RB_COMPTE1')
  else begin
    EffaceLeGrid ;
    RetailLeGrid ;
    AfficheLibelle ;
    ComboAxe(GetField('RB_AXE') <> '') ;
    LitCpteExclu(GetField('RB_COMPTE1')   , 0);
    LitCpteExclu(GetField('RB_EXCLUSION1'), 1);
    LitCpteExclu(GetField('RB_COMPTE2')   , 3);
    LitCpteExclu(GetField('RB_EXCLUSION2'), 4);

    if not (TFFiche(Ecran).FTypeAction in [taCreat, taCreatOne, taCreatEnSerie]) then
      SetControlEnabled('RB_RUBRIQUE', False);

  end ;

  {$IFNDEF TRESO}
  if (VH^.PaysLocalisation<>CodeISOES) then //XVI 24/02/2005
  begin
     if FCtxSurRubrique = CtxBudget then
        PositionneFamille;
  end else
    RB_CLASSERUBchange(cbClasse);

  ChargeListeFam ;

  if FCtxSurRubrique in [CtxLiasse] then
    SetControlEnabled('BCOPIER', (TFFiche(Ecran).FTypeAction in [TaCreat, TaCreatEnSerie]) and not fbPersoStandard)
  else
    SetControlEnabled('BCOPIER', not (TFFiche(Ecran).FTypeAction = taConsult) and not fbPersoStandard);
  {$ENDIF}

  if TFFiche(Ecran).FTypeAction = taConsult then
  begin
(*   for n := 0 to Ecran.ComponentCount - 1 do
     if (UpperCase(Ecran.Components[n].Name) <> 'BFERME'   ) and
        (UpperCase(Ecran.Components[n].Name) <> 'BIMPRIMER') and
        (UpperCase(Ecran.Components[n].Name) <> 'GCPTE1'   ) then
       SetControlEnabled(UpperCase(Ecran.Components[n].Name), False);
  *)
    G.Enabled := True;
    G.Options := G.Options-[goEditing]+[goRowSelect] ;
  end;

  if G.Focused then
     if (VH^.PaysLocalisation=CodeISOEs) and (FormuleRubrique.Visible) then //XVI 24/02/2005
       FormuleRubrique.SetFocus
     else
       G.SetFocus ;
  QuelCarBour ;

  {$IFNDEF TRESO}
  if FCtxSurRubrique <> CtxBudget then
  begin
    SetControlVisible('BFAM', True);
    ch := Trim(GetControlText('RB_RUBRIQUE'));
    if (ch <> '') then begin
      bStdCegid := EstSpecif('51502');
      bStdCegid := (ctxStandard in V_PGI.PGIContexte) and (bStdCegid);
(*
      if ((ch[1]='@') and (not bStdCegid)) then
        SetControlVisible('BFAM', False)
      else if ((ch[1]='$') and (not (ctxStandard in V_PGI.PGIContexte))) then
        SetControlVisible('BFAM', False);
*)
      if ((GetField('RB_PREDEFINI')='CEG') and (not bStdCegid)) then SetControlVisible('BFAM', False)
      else if ((GetField('RB_PREDEFINI')='STD') and (not (ctxStandard in V_PGI.PGIContexte))) then
        SetControlVisible('BFAM', False);
    end;
  end;
  {$ENDIF}

  Ecran.Caption := TraduireMemoire('Rubrique') + ' : ' + GetControlText('RB_RUBRIQUE');
  UpdateCaption(Ecran);
end ;

{$IFNDEF TRESO}
//XMG 05/04/04 début
{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.RempliComboFamRub(C : THValComboBox) ;
{---------------------------------------------------------------------------------------}
//XMG 05/04/04 fin
var
  Q : TQuery ;
  T : TOB;
  n : Integer;
  O : Boolean ;
begin
  Q := OpenSql('SELECT CO_CODE,CO_LIBELLE FROM COMMUN WHERE CO_TYPE="RBB"',True) ;
  T := TOB.Create('', nil, -1);
  T.LoadDetailDB('COMMUN', '', '', Q, False);
  Ferme(Q) ;

  C.Values.Clear ;
  C.Items.Clear ;

  for n := 0 to T.Detail.Count - 1 do begin
    O := True;
    if T.Detail[n].GetValeur(0) = 'GEN' then
      C.Values.Add('CBG')
    else if T.Detail[n].GetValeur(0) = 'ANA' then
      C.Values.Add('CBS')
    else if T.Detail[n].GetValeur(0) = 'G/A' then begin
      if EstSerie(S3) then O := False
                      else C.Values.Add('G/S');
    end
    else if T.Detail[n].GetValeur(0) = 'A/G' then begin
      if EstSerie(S3) then O := False
                      else C.Values.Add('S/G');
    end;
    if O then C.Items.Add(T.Detail[n].GetValeur(1));
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.ChargeListeFam;
{---------------------------------------------------------------------------------------}
var
  LeLibFam : String ;
  QLib : TQuery ;
  ListeFamille : TStringList ;
  St,St1 : String ;
  i : Integer ;
begin
  LeLibFam := GetField('RB_FAMILLES');
  ListFam.Items.Clear ;
  ListeFamille := TStringList.create ;

  if FCtxSurRubrique <> CtxBudget then
  begin
    QLib := OpenSQL('Select YDS_CODE, YDS_LIBELLE FROM CHOIXDPSTD WHERE YDS_TYPE = "RBF"',TRUE) ;
    while not QLib.Eof do
    begin
      ListeFamille.Add(QLib.Fields[0].AsString+'='+QLib.Fields[1].AsString) ;
      QLib.Next ;
    end ;

    Ferme(QLib) ;
    while LeLibFam <> '' do
    begin
      St := ReadTokenSt(LeLibFam) ;
      i  := ListeFamille.IndexOfName(St) ;
      if i >= 0 then St1:=ListeFamille.Values[St]
                else St1:='Non renseigné' ;
      ListFam.Items.Add(St1) ;
    end ;
  end
  else
    ListFam.Items.Add(GetField('FAMBUD'));

  ListeFamille.free ;
end;
{$ENDIF TRESO}

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.ChampsParDefaut;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRESO}
  SetField('RB_TYPERUB', 'TRP');
  SetField('RB_NATURERUB', 'TRE');
  SetField('RB_CLASSERUB', 'TRE');
  SetField('RB_SIGNERUB', 'C');
  {$ELSE}
  SetField('RB_TYPERUB', 'GEN');
  SetField('RB_SIGNERUB', 'POS');
  {$ENDIF}
  SetField('RB_RUBRIQUE', '');
  SetField('RB_CODEABREGE', '');
  SetField('RB_LIBELLE', '');

  EffaceLeGrid ;
  if VH^.PaysLocalisation=CodeISOEs then
  begin
     FormuleRubrique.Lines.Clear;
     RB_CLASSERUBchange(cbClasse);
  End ; //XVI 24/02/2005
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.EffaceLeGrid ;
{---------------------------------------------------------------------------------------}
var
  i,
  j : Integer;
begin
  for i:=1 to G.RowCount-1 do
    for j:=0 to G.ColCount-1 do
      G.Cells[j,i]:='' ;
  G.RowCount:=2 ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.RetailLeGrid ;
{---------------------------------------------------------------------------------------}
{$IFNDEF TRESO}
var
  St : string;
{$ENDIF}
begin
{$IFDEF TRESO}
  G.ColCount := 2;
  G.ColWidths[0] := 300;
  G.ColWidths[1] := 300;
  G.ColCombo := -1;
  {JP 22/01/07 : Nécessaire en Tréso en D7}
  G.ValCombo := nil;
  THValComboBox(GetControl('FCALCUL')).Visible := False;
{$ELSE}
  G.Visible := False;
  SetControlVisible('FCALCUL', False);
  St := cbType.Value;
  if (St='ANA') or (St='BUD') or (St='GEN') or
   (St='TIE') or
   ((VH^.PaysLocalisation=CodeISOEs) and (cbClasse.Value = 'RDR')) then begin //XVI 24/02/2005
    G.ColCount := 3 ;
    G.ColWidths[0] := 185 ;
    G.ColWidths[1] := 125;
    G.ColWidths[2] := 85 ;
  end
  else
  begin
    G.ColCount := 5 ;
    G.ColWidths[0] := 135 ;
    G.ColWidths[1] := 85 ;
    G.ColWidths[2] := 75 ;
    G.ColWidths[3] := 135 ;
    G.ColWidths[4] := 85 ;
  end ;
  G.ColCombo := 2 ;

  G.Visible  := True ;
{$ENDIF}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.AfficheLibelle;
{---------------------------------------------------------------------------------------}
{$IFNDEF TRESO}
var
  St   : string[3] ;
  A, B : Char ;
  CB   : THValComboBox;
{$ENDIF}
begin
  {$IFDEF TRESO}
  G.Cells[0,0] := 'Comptes Généraux';
  G.Cells[1,0] := 'Comptes généraux à exclure';
  {$ELSE}
  CB := THValComboBox(GetControl('CBCAPTION'));
  if not Assigned(G) or not Assigned(CB) then Exit;

  St:= cbType.Value;
  A:=St[1] ;
  B:=St[3] ;
  G.Cells[2,0] := CB.Items[8] ;
  case A of
    'A' : begin G.Cells[0,0] := CB.Items[0] ; G.Cells[1,0] := CB.Items[1] ; end ;
    'B' : begin G.Cells[0,0] := CB.Items[2] ; G.Cells[1,0] := CB.Items[3] ; end ;
    'T' : begin G.Cells[0,0] := CB.Items[4] ; G.Cells[1,0] := CB.Items[5] ; end ;
    'G' : begin G.Cells[0,0] := CB.Items[6] ; G.Cells[1,0] := CB.Items[7] ; end ;
  end ;

  if G.ColCount=5 then begin
    case B of
      'A' : begin G.Cells[3,0] := CB.Items[0] ; G.Cells[4,0] := CB.Items[1] ; end ;
      'B' : begin G.Cells[3,0] := CB.Items[2] ; G.Cells[4,0] := CB.Items[3] ; end ;
      'T' : begin G.Cells[3,0] := CB.Items[4] ; G.Cells[4,0] := CB.Items[5] ; end ;
      'G' : begin G.Cells[3,0] := CB.Items[6] ; G.Cells[4,0] := CB.Items[7] ; end ;
    end ;
  end ;

  if (cbClasse.Value = 'RDR') then begin
    G.Cells[0,0] := CB.Items[9] ;
    G.Cells[1,0] := CB.Items[10] ;
  end;
  {$ENDIF}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.PositionneFamille ;
{---------------------------------------------------------------------------------------}
var
  Value: String;         {FP 15/11/2005 FQ13268}
begin
  {b FP 15/11/2005 FQ13268: Le champ FAMBUD n'existe pas.
  if FCtxSurRubrique = CtxBudget then Exit;

  if cbType.Value = 'GEN' then SetField('FAMBUD', 'CBG')
  else if cbType.Value = 'ANA' then SetField('FAMBUD', 'CBS')
  else if cbType.Value = 'G/A' then SetField('FAMBUD', 'G/S')
  else if cbType.Value = 'A/G' then SetField('FAMBUD', 'S/G');}

  Value := '';
  if cbType.Value = 'GEN' then Value := 'CBG'
  else if cbType.Value = 'ANA' then Value := 'CBS'
  else if cbType.Value = 'G/A' then Value := 'G/S'
  else if cbType.Value = 'A/G' then Value := 'S/G';

  if (Value = '') and (FCtxSurRubrique = CtxBudget) then
    Value := GetField('RB_FAMILLES');

  if Value <> '' then
  begin
    if FCtxSurRubrique = CtxBudget then
    begin
      THValComboBox(GetControl('FAMBUD')).Style := csDropDown;
      THValComboBox(GetControl('FAMBUD')).Text  := Value;
    end
    else
      SetField('FAMBUD', Value);
  end;
  {e FP 15/11/2005 FQ13268}
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.AjouteLigne ;
{---------------------------------------------------------------------------------------}
var
  i    : Integer ;
  Vide : Boolean ;
begin
  Vide := True ;
  for i:=0 to G.ColCount-1 do
    if (G.Cells[i, G.RowCount - 1]<>'') then Vide := False;

  if not Vide then G.RowCount := G.RowCount + 1;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.EffaceLigne ;
{---------------------------------------------------------------------------------------}
var
  ARow : Integer ;
begin
  for ARow:=1 to G.RowCount - 2 do begin
    if G.ColCount = 3 then begin
      if (G.Cells[0,ARow] = '') and (G.Cells[1,ARow] = '') then
        G.DeleteRow(ARow) ;
    end
    else begin
      if (G.Cells[0,ARow] = '') and (G.Cells[1,ARow] = '') and
         (G.Cells[3,ARow] = '') and (G.Cells[4,ARow] = '') then
        G.DeleteRow(ARow) ;
    end ;
  end ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.QuelCarBour;
{---------------------------------------------------------------------------------------}
Var
  St  : String[3] ;
  A,B : Char ;
begin
  St := cbType.Value;
  A := St[1] ;
  B := St[3] ;
  CarBour  := #0 ;
  CarBour1 := #0 ;

  Case A of
    'A' : CarBour := VH^.Cpta[AxeToFb(GetField('RB_AXE'))].Cb ;
    'T' : CarBour := VH^.Cpta[fbAux].Cb ;
    'G' : CarBour := VH^.Cpta[fbGene].Cb ;
  end ;
  if G.ColCount = 5 then begin
    case B of
      'A' : CarBour1 := VH^.Cpta[AxeToFb(GetField('RB_AXE'))].Cb ;
      'T' : CarBour1 := VH^.Cpta[fbAux].Cb ;
      'G' : CarBour1 := VH^.Cpta[fbGene].Cb ;
     end ;
   end ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.EcritCpteExclu;
{---------------------------------------------------------------------------------------}
begin
  if (VH^.PaysLocalisation=CodeISOES) and (GetField('RB_CLASSERUB') = 'RDR') then //XVI 24/02/2005
    SetField('RB_COMPTE1',FormuleRubrique.Lines.text)
  else begin
    EffaceLigne;
    if CpteVide then Exit ;
    Depasse := False ;
    SetField('RB_COMPTE1'   , Grid2St(0));
    SetField('RB_COMPTE2'   , Grid2St(3));
    SetField('RB_EXCLUSION1', Grid2St(1));
    {$IFNDEF TRESO}
    {JP 07/10/04 : en Trésorerie, RB_EXCLUSION2 sert à stocker les modes de paiement}
    SetField('RB_EXCLUSION2', Grid2St(4));
    {$ENDIF}
  end ;
end ;

{---------------------------------------------------------------------------------------}
function TOM_RUBRIQUE.CpteVide : Boolean ;
{---------------------------------------------------------------------------------------}
var
  ARow : Integer ;
begin
  Result := False ;
  for ARow := 1 to G.RowCount-1 do begin
    if (G.Cells[0,ARow] = '') and (G.Cells[1,ARow] <> '')then begin
      AfficheMsg(9,'','') ;
      Result := True ;
      Exit ;
    end;
  end ;

  if G.ColCount = 5 then
  for ARow := 1 to G.RowCount - 1 do begin
    if (G.Cells[3,ARow] = '') and (G.Cells[4,ARow] <> '')then begin
      AfficheMsg(9,'','') ;
      Result:=True ;
      Exit ;
    end;
  end;
end ;

{---------------------------------------------------------------------------------------}
function TOM_RUBRIQUE.EnregOK : Boolean ;
{---------------------------------------------------------------------------------------}
begin
  Result  := False;
  Depasse := False;
  if not (DS.state in [dsEdit,dsInsert]) then begin
    Result := True;
    Exit;
  end;


  if DS.state in [dsInsert] then begin
    if GetControlText('RB_RUBRIQUE') = '' then begin
      AfficheMsg(2, '', '') ;
      SetFocusControl('RB_RUBRIQUE');
      Exit ;
    end;
    {$IFDEF TRESO}
    if ExisteSQL('SELECT RB_RUBRIQUE FROM RUBRIQUE WHERE RB_NATRUB IN ("CPT", "BUD") ' +
                 'AND RB_RUBRIQUE = "' + GetField('RB_RUBRIQUE') + '"') then begin
      LastErrorMsg := TraduireMemoire('Le code est déjà utilisé dans la Comptabilité');
      Exit;
    end;
    {$ELSE}
    if EstComptaTreso and ExisteSQL('SELECT RB_RUBRIQUE FROM RUBRIQUE WHERE RB_NATRUB = "TRE" ' +
                                    'AND RB_RUBRIQUE = "' + GetField('RB_RUBRIQUE') + '"') then begin
      LastErrorMsg := TraduireMemoire('Le code est déjà utilisé dans la Trésorerie');
      Exit;
    end;
    {$ENDIF TRESO}
  end;

  if DS.state in [dsEdit,dsInsert] then
  begin
    if FCtxsurRubrique = CtxBudget then
    begin
      {b FP 15/11/2005 FQ13268 SetField('RB_FAMILLES', THValComboBox(GetControl('FAMBUD')).Value + ';') ;}
      SetField('RB_FAMILLES', THValComboBox(GetControl('FAMBUD')).Text + ';') ;
      {e FP 15/11/2005 FQ13268}
    end;

    if not ControleCodeRubriqueOk then Exit;
    if not ControleCpteAbrOk then Exit ;

    if GetControlText('RB_LIBELLE'  ) = '' then begin AfficheMsg(3 ,'',''); SetFocusControl('RB_LIBELLE'); Exit ; end ;
    {$IFNDEF TRESO}
    {Pas de notion de famille de rubriques en Tréso}
    if GetField('RB_FAMILLES' ) = ''       then begin AfficheMsg(7 ,'',''); Exit ; end ;
    if cbClasse.Value = ''                 then begin AfficheMsg(14,'',''); Exit ; end ;
    {$ENDIF}
    if not AnalyseCompteOk                 then begin AfficheMsg(5,'',''); Exit ; end ;
    if Not BudJalOk then Exit ;
  end;

  EcritCpteExclu;

  Result := True;
  if Depasse then AfficheMsg(8,'','');
{$IFDEF TRESO}
  SetField('RB_NATURERUB', 'TRE');
  SetField('RB_CLASSERUB', 'TRE');
{$ENDIF}
  MemoChampParDefaut ;
end ;

{---------------------------------------------------------------------------------------}
function TOM_RUBRIQUE.AnalyseCompteOk : Boolean ;
{---------------------------------------------------------------------------------------}
Var
  i, j, k : Integer ;
  St      : string ;
  TypChar : set of Char ;
  Okok    : Boolean ;
  VirgOk  : Boolean; {JP 10/10/2003}
begin
  Result := True;
  Okok   := False;

  if cbClasse.Value = 'RDR' then Exit;

  if GetField('RB_TABLELIBRE') = 'X' then
    TypChar:=['0'..'9','A'..'Z',';',':',',','#','*','-','&','|',' ','?',CarBourBud]
  else
    {JP 23/11/04 : a priori le point-virgule n'avait rien à faire dans TypChar}
    TypChar := ['0'..'9', 'A'..'Z', {';', }':', ' ', '?', CarBour, CarBour1];

  G.Col:=0 ; G.Row:=1 ;

  for j := 1 to G.RowCount-1 do
    for i := 0 to G.ColCount-1 do begin
      if i = 2 then Continue ;
      if G.Cells[i, j] <> '' then begin
        St := G.Cells[i, j];
        Okok := True;
        {JP 10/10/03 : Rien n'empêche de saisir deux fois ":"}
        VirgOk := False;
        for k := 1 to Length(St) do begin
          if ((St[k] = ':') and VirgOk) or
             not(St[k] in TypChar) then begin
            {$IFDEF TRESO}
            {23/11/04 : FQ TRESO 10178 : on autorise la vigule sur les comptes exclus comme séparateur
                        de comptes ou de fourchettes de comptes}
            if (St[k] = ',') and (i = 1) then Continue;
            {$ENDIF}
            G.SetFocus;
            G.Col  := i;
            G.Row  := j;
            Result := False;
            Exit ;
          end;
          if St[k] = ':' then VirgOk := True;
        end;
      end;
    end;

  if not Okok then Result := False;
end ;

{---------------------------------------------------------------------------------------}
function TOM_RUBRIQUE.BudJalOk : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := True;
(* A FAIRE : pb sur RB_BUDJAL.ExisteH
  Result := False ;
  if (Not Budget) or (Not (Action in [taCreat..taCreatOne])) then begin
    Result := True;
    Exit;
  end;

  if RB_BUDJAL.ExisteH <= 0 then begin
    AfficheMsg(11,'','') ;
    SetFocusControl('RB_BUDJAL');
    Exit;
  end;
  Result := True ;
  *)
end ;

{---------------------------------------------------------------------------------------}
function TOM_RUBRIQUE.ControleCodeRubriqueOk : Boolean;
{---------------------------------------------------------------------------------------}
{$IFDEF TRESO}
var
  ch        : string;
{$ENDIF}
begin
  Result := False ;
  if GetControlText('RB_RUBRIQUE') = '' then Exit;
  {$IFDEF TRESO}
  ch := GetControlText('RB_RUBRIQUE');
  if tCodeFlux.IndexOf(ch) > -1 then begin
    AfficheMsg(21, '', '');
    Exit;
  end;
  {$ELSE}
  {$ENDIF}
  Result := True ;
end;

{---------------------------------------------------------------------------------------}
function TOM_RUBRIQUE.ControleCpteAbrOk : Boolean ;
{---------------------------------------------------------------------------------------}
begin
  Result:=False ;
  if GetControlText('RB_CODEABREGE') = '' then begin
    AfficheMsg(12,'','') ;
    SetFocusControl('RB_CODEABREGE');
    Exit ;
  end ;
  Result:=True ;
end ;

{---------------------------------------------------------------------------------------}
function TOM_RUBRIQUE.Grid2St( ACol : Integer ) : String ;
{---------------------------------------------------------------------------------------}
Var
  St, StCal,
  StFin     : string ;
  ARow      : Integer ;
  CB      : THValComboBox;
begin
  CB := THValComboBox(GetControl('FCALCUL'));

  if ACol >= G.ColCount then begin
    Result := StFin;
    Exit ;
  end ;

  for ARow := 1 to G.RowCount-1  do begin
    CB.Libelle := G.Cells[2,ARow] ;
    {$IFNDEF TRESO}
    if CB.Value <> '' then StCal := '(' + CB.Value + ')'
                      else StCal := '(SM)' ;
    {$ENDIF}
    St:=G.Cells[ACol,ARow];

    if St <> '' then begin
      if (ACol = 1) or (ACol = 4) then StFin := StFin + St
                                  else StFin := StFin + St + StCal;
    end ;
    StFin := StFin + ';' ;
  end ;

  if Length(StFin) > 250 then begin
    Depasse := True ;
    TronqueLeCompte(StFin) ;
  end ;
  Result := StFin ;
end ;


{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.TransFormeCode(var Chaine : string ; PVirgule : Boolean);
{---------------------------------------------------------------------------------------}
begin
  if Chaine = '' then Exit;

  if PVirgule then While Pos(',',Chaine) > 0 do Chaine[Pos(',',Chaine)] := ';'
              else While Pos(';',Chaine) > 0 do Chaine[Pos(';',Chaine)] := ',' ;
  if PVirgule then Chaine := Chaine + ';'
              else Delete(Chaine,Length(Chaine),1) ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.TronqueLeCompte(var Chaine : string);
{---------------------------------------------------------------------------------------}
var
  St,
  St1 : String ;
begin
  while Chaine <> '' do begin
    St1 := ReadTokenSt(Chaine) ;
    if Length(St) + Length(St1) + 1 > 249 then Break ;
    St := St + St1 + ';' ;
  end ;
  Chaine := St ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.ComboAxe(Avec : Boolean) ;
{---------------------------------------------------------------------------------------}
begin
  if Avec then SetControlProperty('RB_AXE','COLOR', clWindow)
          else SetControlProperty('RB_AXE','COLOR', clBtnFace);
  SetControlEnabled('TRB_AXE', Avec);
  SetControlEnabled('RB_AXE', Avec);
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.LitCpteExclu(St : String ; Col : Byte);
{---------------------------------------------------------------------------------------}
var
  St1, St2,
  StCalc  : string ;
  i, j    : Byte ;
  CB      : THValComboBox;
begin
  i := 1;
  CB := THValComboBox(GetControl('FCALCUL'));
  if not Assigned(G) or not Assigned(CB) then Exit;

  while St<>'' do begin
    St1 := ReadTokenSt(St) ;
    StCalc := '' ;
    St2    := '' ;
    for j := 1 to Length(St1)do begin
       if St1[j]='(' then begin
         StCalc := Copy(St1, j+1, 2);
         Break;
       end
       else
         St2 := St2 + St1[j] ;
    end ;

    if StCalc<>'' then
      G.Cells[2,i] := CB.Items[CB.Values.IndexOf(StCalc)];
    G.Cells[Col,i] := St2 ;
    Inc(i);
    if i > G.RowCount - 1 then
      AjouteLigne ;
  end ;
end ;

{$IFDEF TRESO}
{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.RB_SIGNERUBChange (Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  {Dans la fiche de suivi, j'ai besoin de deux champs définissant le signe :
   RB_SIGNERUB = C(rédit) - > RB_TYPERUB := TRP(ositif)
   RB_SIGNERUB = D(ébit)  - > RB_TYPERUB := TRN(égatif)}
  if cbSigne.Value = 'C' then SetField('RB_TYPERUB', 'TRP')
                         else SetField('RB_TYPERUB', 'TRN')
end;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.OnArgumentTreso;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  L : THLabel;
begin
  cbModePaie := THMultiValComboBox(GetControl('CBMODEPAIE'));
  {La gestion des rubriques en Tréso permet juste de regrouper les comptes sous un libellé
   et de simuler un code flux}
  SetControlProperty('RB_RUBRIQUE', 'MAXLENGTH', 3);
  SetControlProperty('RB_CODEABREGE', 'MAXLENGTH', 3);

  SetControlVisible('BCOPIER', False);
  SetControlVisible('BIMPRIMER', False);
  SetControlVisible('TRB_CLASSERUB', False);
  SetControlVisible('RB_CLASSERUB', False);
  SetControlVisible('RB_TABLELIBRE', False);
  SetControlVisible('GBCARACTER', False);

  {19/07/04 : J'utilise ce boolean pour stocker si les écritures sont soumises
              au calcul des commissions de mouvements}
  SetControlVisible('RB_CATEGORIE', True);
  SetControlCaption('RB_CATEGORIE', TraduireMemoire('Soumis à commission de mvts'));

  {Gestion du "signe" de la rubrique}
  cbSigne.OnChange := RB_SIGNERUBChange;
  cbSigne.Parent   := TWinControl(GetControl('GBPARAM'));
  cbSigne.Top      := 39;
  cbSigne.Left     := 249;
  cbSigne.DataType := '';
  cbSigne.Width    := 130;

  cbSigne.TabOrder := 3;
  cbSigne.TabOrder := 3;

  cbSigne.Values.Clear;
  cbSigne.Values.Add('C');
  cbSigne.Values.Add('D');
  cbSigne.Items.Clear;
  cbSigne.Items.Add('Crédit');
  cbSigne.Items.Add('Débit');

  L := THLabel(GetControl('TRB_SIGNERUB'));
  L.Parent := TWinControl(GetControl('GBPARAM'));
  L.Top    := 43;
  L.Left   := 214;

  {27/09/04 : FQ 10119 : On stocke le mode de paiement dans le champ RB_BUDJAL}
  cbModePaie.Visible := True;
  cbModePaie.Top      := 39;
  cbModePaie.Left     := 480;
  cbModePaie.DataType := 'TTMODEPAIE';
  cbModePaie.Enabled  := True;
  cbModePaie.Width    := 135;
  cbModePaie.Color    := clWindow;
  cbModePaie.OnChange := MODEPAIEChange;

  L := THLabel(GetControl('TRB_BUDJAL'));
  L.Parent  := TWinControl(GetControl('GBPARAM'));
  L.Top     := 43;
  L.Left    := 388;
  L.Caption := 'Mode de paiement';
  L.Enabled := True;

  {10/08/06 : gestion des autres composants qui semblent avoir été déplacés pour les besoins de la compta}
  {$IFDEF EAGLCLIENT}
  THEdit(GetControl('RB_LIBELLE')).Left := cbSigne.Left;
  TCheckBox(GetControl('RB_CATEGORIE')).Left := THEdit(GetControl('RB_LIBELLE')).Left +
                                                THEdit(GetControl('RB_LIBELLE')).Width + 10;
  {$ELSE}
  THDBEdit(GetControl('RB_LIBELLE')).Left := cbSigne.Left;
  TDBCheckBox(GetControl('RB_CATEGORIE')).Left := THDBEdit(GetControl('RB_LIBELLE')).Left +
                                                  THDBEdit(GetControl('RB_LIBELLE')).Width + 10;
  {$ENDIF EAGLCLIENT}
  THLabel(GetControl('TRB_LIBELLE')).Left := THLabel(GetControl('TRB_SIGNERUB')).Left;
  {10/08/06 : gestion des autres composants qui semblent avoir été déplacés pour les besoins de la compta}

  {En tréso les rubriques sont assimilées à des codes flux pour les écritures importées
   de la compta}
  Q := OpenSQL('SELECT TFT_FLUX FROM FLUXTRESO', True);
  tCodeFlux := TStringList.Create;
  tCodeFlux.Duplicates := dupIgnore;
  tCodeFlux.Sorted     := True;
  while not Q.EOF do begin
    tCodeFlux.Add(Q.Fields[0].AsString);
    Q.Next;
  end;
  Ferme(Q);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.GetModePaiement;
{---------------------------------------------------------------------------------------}
begin
  if not Assigned(cbModePaie) then cbModePaie := THMultiValComboBox(GetControl('CBMODEPAIE'));
  cbModePaie.Value := GetField('RB_EXCLUSION2');
end;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.SetModePaiement;
{---------------------------------------------------------------------------------------}
begin
  SetField('RB_EXCLUSION2', cbModePaie.Value);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.MODEPAIEChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if THMultiValComboBox(Sender).Focused and not (DS.State in [dsInsert, dsEdit]) then
    DS.Edit;
end;
{$ENDIF}

{Gestion des message}
{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.InitMsg;
{---------------------------------------------------------------------------------------}
begin
  Msg := THMsgBox.Create(FMenuG);
  Msg.Mess.Add('0;Rubriques;Voulez-vous enregistrer les modifications?;Q;YNC;Y;C;');
  Msg.Mess.Add('1;Rubriques;Confirmez-vous la suppression de l''enregistrement?;Q;YNC;N;C;');
  Msg.Mess.Add('2;Rubriques;Vous devez renseigner un code.;W;O;O;O;');
  Msg.Mess.Add('3;Rubriques;Vous devez renseigner un libellé.;W;O;O;O;');
  Msg.Mess.Add('4;Rubriques;Le code que vous avez saisi existe déjà. Vous devez le modifier.;W;O;O;O;');
  Msg.Mess.Add('5;Rubriques;Votre saisie des comptes est incorrecte;W;O;O;O;');
  Msg.Mess.Add('L''enregistrement est inaccessible');
  Msg.Mess.Add('7;Rubriques;Vous devez renseigner une famille.;W;O;O;O;');
  Msg.Mess.Add('8;Rubriques;ATTENTION : Votre saisie des comptes est trop longue, elle sera tronquée.;E;O;O;O;');
  Msg.Mess.Add('9;Rubriques;ATTENTION : Vous avez saisi des comptes d''exclusions et pas de comptes à inclure;E;O;O;O;');
  Msg.Mess.Add('Paramétrage des variables:');
  Msg.Mess.Add('11;Rubriques;Le code budget renseigné est vide ou il n''existe pas.;W;O;O;O;');
  Msg.Mess.Add('12;Rubriques;Vous devez renseigner un code abrégé.;W;O;O;O;');
  Msg.Mess.Add('13;Rubriques;Le code abrégé que vous avez saisi existe déjà. Vous devez le modifier.;W;O;O;O;');
  Msg.Mess.Add('14;Rubriques;Vous devez renseigner une classe de rubriques;W;O;O;O;');
  Msg.Mess.Add('15;Rubriques;Le code rubrique doit commencer par @ (Standard Cegid).;W;O;O;O;');
  Msg.Mess.Add('16;Rubriques;Le code rubrique doit commencer par $ (Standard).;W;O;O;O;');
  Msg.Mess.Add('17;Rubriques;Le code rubrique ne doit pas commencer par $ ou @ (réservé aux standards).;W;O;O;O;');
  Msg.Mess.Add('18;Rubriques;Opération impossible : Vous n''avez pas renseigné le code identifiant de la rubrique.;W;O;O;O;');
  Msg.Mess.Add('Choix d''une rubrique à dupliquer');
  Msg.Mess.Add('20;Rubriques;Confirmez-vous la duplication de la rubrique sélectionnée ?;Q;YN;N;N;');
  Msg.Mess.Add('21;Rubriques;Le code saisi est déjà affecté à un code flux ! Veuillez en saisir un autre.;W;O;O;O;');
end;

{---------------------------------------------------------------------------------------}
function TOM_RUBRIQUE.AfficheMsg(num : Integer; Av,Ap : string ) : Word ;
{---------------------------------------------------------------------------------------}
begin
  Result := mrNone ;
  if Num > 21 then Exit ;
  Result := Msg.Execute(num,Av,Ap) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.FermeMsg ;
{---------------------------------------------------------------------------------------}
begin
  Msg.Free ;
end;
(*
{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.BRubriquesClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var lesfam,unefam,Where : String ;
    PosCaret            : Longint ;
    pp1                 : LongRec ;
    PP                  : TPoint ;
begin
   lesFam := GetField('RB_FAMILLES') ; //XMG 05/04/04
   where  := '' ;
   while lesfam <> '' do Begin
     if Trim(where) <> '' then Where:=Where+' or' ;
     UneFam := readtokenst(lesFam) ;
     Where  := Where+' RB_FAMILLES like "%' + UneFam + ';%"' ;
   end ;
   if trim(where)<>'' then begin
      FormuleRubrique.SetFocus ;
      laRubrique.Text := '' ;
      posCaret := FormuleRubrique.SelStart ;
      PP1 := LongRec(MaxIntValue([0,SendMessage(FormuleRubrique.Handle, EM_POSFROMCHAR, PosCaret, 0)]));
      pp.X := PP1.Lo;
      PP.Y := PP1.Hi;
      pp := Ecran.ScreentoClient(FormuleRubrique.ClientToScreen(PP)) ;
      LaRubrique.left := pp.X;
      LaRubrique.Top  := pp.Y;
      if Lookuplist(LaRubrique, TraduireMemoire('Les rubriques'), 'RUBRIQUE', 'RB_RUBRIQUE', 'RB_LIBELLE', Where, 'RB_RUBRIQUE',TRUE,0) then Begin
        LaRubrique.Text := '[' + LaRubrique.Text + ']';
        FormuleRubrique.Lines.Text := Copy(FormuleRubrique.Lines.Text , 1, FormuleRubrique.SelStart) + LaRubrique.Text+Copy(FormuleRubrique.Lines.Text,FormuleRubrique.SelStart + 1 + FormuleRubrique.Sellength, Length(FormuleRubrique.Lines.Text));
        FormuleRubrique.SelStart   := PosCaret + Length(LaRubrique.Text);
      end ;
   end ;
end;
//XVI 24/02/2005
*)

{JP 15/05/07 : FQ 20340 : problème lorsque la grille prend le focus => on force un changement de cellule
{---------------------------------------------------------------------------------------}
procedure TOM_RUBRIQUE.OnAfterFormShow;
{---------------------------------------------------------------------------------------}
begin
  SendMessage(G.Handle, WM_KEYDOWN, VK_TAB,  0);
  SendMessage(G.Handle, WM_KEYDOWN, VK_LEFT, 0);
end;

initialization
  RegisterClasses ( [ TOM_RUBRIQUE ] ) ;

end.



