{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 30/04/2003
Modifié le ... : 12/06/2003
Description .. : Source TOF de la FICHE : MULRELDETAIL ()
Suite ........ : 
Suite ........ : 
Mots clefs ... : TOF;MULRELDETAIL
*****************************************************************}
Unit UTOFMULRELDETAIL ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     AglInit,
     ed_tools,
     extCtrls,        // pour le TPanel
{$IFDEF EAGLCLIENT}
     MaineAGL,
     eMul,
     UtileAGL,       // pour le LanceDocument
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Mul,
     Fe_Main,
     EdtRDoc,       // pour le LanceDocument
     SaisBor ,
     Hdb,
{$ENDIF}
     Saisie,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HTB97,
     HEnt1,
     Ent1,
     HMsgBox,
     HStatus,         // pour FiniMove et MoveCur
     uRecupSQLModele, // pour RecupSQLModele
     SaisUtil,        // pour GetInfosDevise et RevisionActive
     EcheMono,        // pour SaisirMonoEcheance
     RelvUtil,        // pour EC_RELF
     LettUtil,        // pour CodeSuivant
     SaisComm,        // pour OBMToIdent
     UTOF,
     HQry,
     UTOB ;

{$IFDEF EAGLCLIENT}
  Procedure GenereReleves ( vTiers : TOB ; vIndex : Integer) ;
{$ELSE}
//  Procedure GenereReleves ( vTiers : TDataSource ) ;
  Procedure GenereReleves ( vTiers : THQuery ) ;
{$ENDIF}
Procedure CPLanceFiche_MulRelDetail ( vRange, vLequel, vArgs : String ) ;

Type
  TOF_MULRELDETAIL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    // Ancêtre
  	Q                   : THQuery ;
{$IFDEF EAGLCLIENT}
    FListe              : THGrid ;
{$ELSE}
    FListe              : THDBGrid ;
{$ENDIF}

    // Messages
    HRel           : THMsgBox ;

    // variables
    TPiece        : TList ;
    DateGene      : TDateTime ;

{$IFDEF EAGLCLIENT}
    FTiers     : TOB ;
    FIndex     : Integer ;
{$ELSE}
//    FTiers     : TDataSource ;
    FTiers     : THQuery ;
{$ENDIF}

    YATP          : boolean ;
    DernLettre    : String ;
    RegimeTva     : String ;
    Gene          : String ;
    Auxi          : String ;
    Libelle       : String ;
    Devise        : String ;
    DEV           : RDEVISE ;
    LeM           : RMVT ;
    GeneNumRel    : Longint ;
    GeneMasqueRel : String17 ;
    GeneSoucheP   : String3 ;
    NowFutur      : TDateTime ;
    LeMode        : String3 ;
    LaEche        : TDateTime ;
    LaValeur      : TDateTime ;
    DateMin       : TDateTime ;
    DateMax       : TDateTime ;
    REL           : EC_RELF ;
    TabTvaEnc     : Array[1..5] of Double ;

    gszCaption              : String ;
    bSelectAllClickParent   : TNotifyEvent ;

    // Initialisations
    procedure InitVariables ;
    procedure InitComposants ;
    procedure InitEvenements ;
    procedure InitMsgBox ;
    procedure InitCriteres ;
    procedure InitTiers ;
    // Calcul solde sélection
    Function  CalculSoldeSelection : Double ;
    Procedure AfficheSoldeSelection ;
    procedure CalculCreditDebit( var TDD,TCD,TDP,TCP : Double ) ;
    procedure CalculCreditDebitLigne( var TDD,TCD,TDP,TCP : Double ) ;
    // EVT GRILLE
    procedure FListeFlipSelection(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    // EVT BOUTONS
    procedure BOuvrirClick(Sender: TObject);
    procedure bSelectAllClick(Sender: TObject);
    procedure BPrecedentClick(Sender: TObject);
    procedure BSuivantClick(Sender: TObject);
    procedure BParamModeleClick(Sender: TObject);
    procedure BZoomPieceClick(Sender: TObject);
    // EVT AUTRES
    procedure DATEGENEREExit(Sender: TObject);
    // Traitements
    procedure MajNavigation ;
    function  VerifCoher : boolean ;
    function  JeValide : boolean ;
    procedure FabricReleve ;
    procedure LettrageReleveOrigine ( CodeL : String4 ; REL : EC_RELF ) ;
    function  UpdateTOBMOrigine( CodeL : String4 ; REL : EC_RELF ) : Boolean ;
    Function  MajTiersRelLet ( Auxi : String ) : String ;
    procedure EditeRelevesTP ;
    procedure RechercheDernCodeLett ;
  public
    FindFirst : boolean ;
  end ;

Implementation
uses
    {$IFDEF eAGLCLIENT}
    MenuOLX
    {$ELSE}
    MenuOLG
    {$ENDIF eAGLCLIENT}     
    ,edtdoc // EditDocument
    ;

//==============================================================================
{$IFDEF EAGLCLIENT}
  Procedure GenereReleves ( vTiers : TOB ; vIndex : Integer) ;
  begin
    TheData  := vTiers ;
    CPLanceFiche_MulRelDetail ( '', '', IntToStr( vIndex ) ) ;
  end ;
{$ELSE}
//  Procedure GenereReleves ( vTiers : TDataSource ) ;
  Procedure GenereReleves ( vTiers : THQuery ) ;
  begin
    TheData := vTiers ;
    CPLanceFiche_MulRelDetail ( '', '', '' ) ;
  end ;
{$ENDIF}

Procedure CPLanceFiche_MulRelDetail ( vRange, vLequel, vArgs : String ) ;
begin
  AGLLanceFiche( 'CP', 'CPMULRELDETAIL', vRange, vLequel, vArgs ) ;
end ;
//==============================================================================

{ TOF_MULRELDETAIL }


procedure TOF_MULRELDETAIL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_MULRELDETAIL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_MULRELDETAIL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_MULRELDETAIL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_MULRELDETAIL.OnArgument (S : String ) ;
begin
  Inherited ;

  InitVariables ;
  InitComposants ;
  InitEvenements ;
  InitMsgBox ;

{$IFDEF EAGLCLIENT}
  FTiers  := Tob(TheData) ;
  FIndex  := ValeurI( ReadTokenSt(S) ) ;
{$ELSE}
  FTiers := THQuery(TheData) ;
{$ENDIF}
  TheData := Nil ;

  // Init date de génération des pièces
  SetControlText('DateGenere', DateToStr(V_PGI.DateEntree) ) ;
  DateGene        := V_PGI.DateEntree ;

  // Init visibilité choix du modèle
  if ((Not VH^.OuiTP) or (EstSerie(S3))) then
    begin
    SetControlVisible('MODELETP', False ) ;
    SetControlVisible('TMODELETP', False ) ;
    end ;

  // On cache les onglets non gérés
  TTabSheet(GetControl('PComplement', True)).TabVisible := False ;
  TTabSheet(GetControl('Pavance', True)).TabVisible := False ;

  // Init Titres et Critères
  MajNavigation ;
  InitTiers ;
  InitCriteres ;

(*
  // Lance la recherche auto
  TFMul(Ecran).BChercheClick(Nil) ;

  // Si aucun enregistrement, on ferme
{$IFDEF EAGLCLIENT}
  if Q.TQ.Detail.Count = 0 then
{$ELSE}
  if Q.Bof and Q.Eof then
{$ENDIF}
    BEGIN
    HRel.Execute(4,'','') ;
    TFMul(Ecran).Close ;
    Exit ;
    END ;
*)
end ;

procedure TOF_MULRELDETAIL.OnClose ;
begin
  Inherited ;
  VideListe(TPIECE) ;
  TPiece.Free ;
end ;

procedure TOF_MULRELDETAIL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_MULRELDETAIL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_MULRELDETAIL.InitComposants;
begin
  // Ancêtre
 	Q      := THQuery(GetControl('Q', True)) ;
{$IFDEF EAGLCLIENT}
  FListe := THGrid(GetControl('FListe',True)) ;
{$ELSE}
  FListe := THDBGrid(GetControl('FListe',True)) ;
{$ENDIF}

  // Boutons    

{$IFDEF EAGLCLIENT}
  SetControlVisible('BSELECTALL', False );
{$ENDIF}
end;

procedure TOF_MULRELDETAIL.InitEvenements;
Var lBouton : TButton ;
begin
  // Grille
{$IFDEF EAGLCLIENT}
  FListe.OnDblClick := FListeDblClick ;
{$ELSE}
  FListe.OnDblClick := FListeDblClick ;
{$ENDIF}
  FListe.OnFlipSelection   := FListeFlipSelection ;

  // Affectation EVT Boutons ancêtre
    // -> tout sélectionner
    lBouton := TButton(GetControl('bSelectAll',True)) ;
    if Assigned(lBouton.OnClick)
      then bSelectAllClickParent    := lBouton.OnClick
      else bSelectAllClickParent    := nil ;
    lBouton.OnClick := bSelectAllClick ;
    // -> BOuvrir
  TButton(GetControl('BOuvrir',True)).OnClick      := BOuvrirClick ;

  // Autres Boutons
  TButton(GetControl('BParamModele',True)).OnClick := BParamModeleClick ;
  TButton(GetControl('BZoomPiece',True)).OnClick   := BZoomPieceClick ;

  // Filtres
  TButton(GetControl('BPrecedent',True)).OnClick   := BPrecedentClick ;
  TButton(GetControl('BSuivant',True)).OnClick     := BSuivantClick ;
  THEdit(GetControl('DateGenere',True)).OnExit     := DateGenereExit ;

end;

procedure TOF_MULRELDETAIL.InitMsgBox;
begin
  HRel := THMsgBox.create(FMenuG);
  HRel.Mess.Clear ;
  HRel.Mess.Add('0;Relevés de factures;Vous n''avez sélectionné aucune échéance.;E;O;O;O;') ;
  HRel.Mess.Add('1;Relevés de factures;Confirmez-vous la génération des relevés pour les échéances sélectionnées ?;Q;YN;Y;Y;') ;
  HRel.Mess.Add('2;Relevés de factures;Vous devez définir un compteur de relevé et un journal associé.;W;O;O;O;') ;
  HRel.Mess.Add('Relevés de facture du tiers') ;
  HRel.Mess.Add('4;Relevés de factures;Ce tiers n''a aucune échéance à générer.;W;O;O;O;') ;
  HRel.Mess.Add('5;Relevés de factures;Le journal associé au compteur est incorrect. Vous devez le modifier.;W;O;O;O;') ;
  HRel.Mess.Add('6;Relevés de facture;Echéances ou modes de paiement diffèrents : voulez-vous choisir et rendre identiques ces informations ?;Q;YN;Y;Y;') ;
  HRel.Mess.Add('ATTENTION! Certains mouvements en cours de traitement n''ont pas pu générer de relevés.') ;
  HRel.Mess.Add('8;Relevés de factures;Vous devez renseigner une date valide.;W;O;O;O;') ;
  HRel.Mess.Add('9;Relevés de factures;La date que vous avez renseignée est sur un exercice non ouvert.;W;O;O;O;') ;
  HRel.Mess.Add('10;Relevés de factures;La date que vous avez renseignée est sur un exercice non ouvert.;W;O;O;O;') ;
  HRel.Mess.Add('11;Relevés de factures;La date que vous avez renseignée est antérieure à la clôture provisoire.;W;O;O;O;') ;
  HRel.Mess.Add('12;Relevés de factures;La date que vous avez renseignée est antérieure à la clôture définitive.;W;O;O;O;') ;
  HRel.Mess.Add('13;Relevés de factures;Le journal associé au compteur n''est pas multi-devises. Vous devez le modifier.;W;O;O;O;') ;
  HRel.Mess.Add('EURO') ;
  HRel.Mess.Add('15;?caption?;Vous avez edité des relevés pour des tiers payeurs. Voulez-vous éditer des documents à destination des tiers facturés ?;Q;YN;Y;Y;') ;

end;

procedure TOF_MULRELDETAIL.InitVariables;
begin
  TPIECE      := TList.Create ;
  YATP        := False ;
  DernLettre  := '' ;
  RegimeTva   := '' ;
//  ModeOppose  := False ;
end;

procedure TOF_MULRELDETAIL.InitCriteres;
begin

  // Date comptable début
  if VH^.Precedent.Code<>''
    then SetControlText('E_DATECOMPTABLE' , DateToStr(VH^.Precedent.Deb) )
    else SetcontrolText('E_DATECOMPTABLE' , DateToStr(VH^.Encours.Deb) ) ;

  // date comptable fin
  SetControlText('E_DATECOMPTABLE_', DateToStr(V_PGI.DateEntree) ) ;

  // Devise
  if GetControlText('E_DEVISE') = '' then
    SetControlProperty('E_DEVISE', 'ItemIndex', 0 ) ;

  // Etablissement
  PositionneEtabUser( THValComboBox(GetControl('E_ETABLISSEMENT', True) ) ) ;
  if GetControlText('E_ETABLISSEMENT') = '' then
    SetControlProperty('E_ETABLISSEMENT', 'ItemIndex', 0 ) ;

end;

procedure TOF_MULRELDETAIL.InitTiers;
begin

{$IFDEF EAGLCLIENT}
  Auxi    := FTiers.Detail[FIndex].GetValue('REL_AUXILIAIRE') ;
  Gene    := FTiers.Detail[FIndex].GetValue('REL_GENERAL') ;
  Libelle := FTiers.Detail[FIndex].GetValue('REL_LIBELLE') ;
  Devise  := FTiers.Detail[FIndex].GetValue('REL_DEVISE') ;
{$ELSE}
  Auxi    := FTiers.FindField('REL_AUXILIAIRE').AsString ;
  Gene    := FTiers.FindField('REL_GENERAL').AsString ;
  Libelle := FTiers.FindField('REL_LIBELLE').AsString ;
  Devise  := FTiers.FindField('REL_DEVISE').AsString ;
{$ENDIF}

  gszCaption := HRel.Mess[3] + ' ' + Auxi + ' ' + Libelle ;
  Ecran.Caption := gszCaption ;

  FillChar( DEV, Sizeof(DEV), #0 ) ;
  DEV.Code  := Devise ;
  GetInfosDevise(DEV) ;
  DEV.Taux  := GetTaux( DEV.Code, DEV.DateTaux, DateGene) ;

  SetControlText( 'E_AUXILIAIRE' ,  Auxi ) ;
  SetControlText( 'E_GENERAL' ,     Gene ) ;
  SetControlText( 'E_DEVISE' ,      Devise ) ;

  UpdateCaption(Ecran) ;

  // Init Dernier relevé pour le tiers
  RechercheDernCodeLett ;

end;

Function  TOF_MULRELDETAIL.CalculSoldeSelection : Double ;
Var i 				  		: Integer ;
    PCumul          : TPanel ;
{$IFNDEF EAGLCLIENT}
		debit, credit  	: Double ;
{$ENDIF}
begin
	Result := 0 ;
  PCumul := TPanel( GetControl('PCumul', True) ) ;
  if (PCumul.FindChildControl('__QRYPCumul_E_DEBIT') = nil)
  	 or (PCumul.FindChildControl('__QRYPCumul_E_CREDIT') = nil) then Exit ;
{$IFDEF EAGLCLIENT}
  // Pas de mode "Tout sélectionné en CWAS
{$ELSE}
  if FListe.AllSelected then
    begin
	  debit  := Valeur(THNumEdit(PCumul.FindChildControl('__QRYPCumul_E_DEBIT')).text ) ;
  	credit := Valeur(THNumEdit(PCumul.FindChildControl('__QRYPCumul_E_CREDIT')).text ) ;
	  Result := debit - credit ;
    end
  else	// Sinon, on parcours le Bookmark
{$ENDIF}
    for i:=0 to FListe.NbSelected-1 do
      BEGIN
      FListe.GotoLeBookmark(i) ;
{$IFDEF EAGLCLIENT}
      Q.TQ.Seek(FListe.row - 1) ;
{$ENDIF}
      Result := Result + Q.FindField('E_DEBIT').asFloat - Q.FindField('E_CREDIT').asFloat ;
      END ;
End ;

Procedure TOF_MULRELDETAIL.AfficheSoldeSelection ;
Var Solde : Double;
begin
	// Calcul du solde
	Solde := CalculSoldeSelection;
  // Ajout affichage du cumul
  if Solde < 0
  	then Ecran.Caption := gszCaption + ' (Cumul de la sélection : ' + StrFMontant(-1*Solde,0,V_PGI.OkDecV,'',True) + ' C)'
  	else Ecran.Caption := gszCaption + ' (Cumul de la sélection : ' + StrFMontant(Solde,0,V_PGI.OkDecV,'',True) + ' D)';
	// Raffraichissement Titre
	UpdateCaption(Ecran) ;
end ;

procedure TOF_MULRELDETAIL.FListeFlipSelection(Sender: TObject);
begin
	AfficheSoldeSelection ;
end;

procedure TOF_MULRELDETAIL.bSelectAllClick(Sender: TObject);
begin
  if Assigned(bSelectAllClickParent) then
    bSelectAllClickParent(Sender) ;
	AfficheSoldeSelection ;
end ;

procedure TOF_MULRELDETAIL.BPrecedentClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  if FIndex <= 0 then Exit ;
  FIndex := FIndex - 1 ;
{$ELSE}
  if FTiers.Bof then Exit ;
  FTiers.Prior ;
{$ENDIF}
  MajNavigation ;
  InitTiers ;
  TFMul(Ecran).BChercheClick(nil) ;
end ;

procedure TOF_MULRELDETAIL.BSuivantClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  if FIndex >= (FTiers.Detail.Count - 1) then Exit ;
  FIndex := FIndex + 1 ;
{$ELSE}
  if FTiers.Eof then Exit ;
  FTiers.Next ;
{$ENDIF}
  MajNavigation ;
  InitTiers ;
  TFMul(Ecran).BChercheClick(nil) ;
end ;

procedure TOF_MULRELDETAIL.BParamModeleClick(Sender: TObject);
begin
{ FQ 19667 BVE 23.04.07 }
  EditDocument('L','RLV','RLW',True) ;
(*
{$IFDEF EAGLCLIENT}
  // Non compatible eAGL
{$ELSE}
  EditDocumentS5S7('L','RLV','RLV',True) ;
{$ENDIF}
*)
{ END FQ 19667 }            //YMOO
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 10/04/2002
Modifié le ... : 18/08/2004
Suite ........ : - LG - 18/08/2004 - Suppression de la fct debutdemois pour 
Suite ........ : l'appel de la saisie bor, ne fct pas avec les exercices 
Suite ........ : decalees
Mots clefs ... : 
*****************************************************************}
procedure TOF_MULRELDETAIL.BZoomPieceClick(Sender: TObject);
Var
    lStModeSaisie     : String ;
  {$IFDEF EAGLCLIENT}
	  TOBListe,TOBLigne : TOB ;
  {$ELSE}
    P                 : RParFolio ;
  {$ENDIF}
begin

  {$IFDEF EAGLCLIENT}
  Q.TQ.Seek(FListe.Row - 1) ;
  {$ENDIF}

  lStModeSaisie := Q.FindField('E_MODESAISIE').AsString ;
  if ((lStModeSaisie<>'-') and (lStModeSaisie<>'')) then
    begin
  {$IFDEF EAGLCLIENT}
     // A FAIRE
    PGIInfo('Fonction non compatible CWAS', 'Consultation d''un bordereau' ) ;
  {$ELSE}
    FillChar(P, Sizeof(P), #0) ;
    P.ParPeriode  := DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime) ;
    P.ParCodeJal  := Q.FindField('E_JOURNAL').AsString ;
    P.ParNumFolio := IntToStr(Q.FindField('E_NUMEROPIECE').AsInteger) ;
    P.ParNumLigne := Q.FindField('E_NUMLIGNE').AsInteger ;
    ChargeSaisieFolio(P, taConsult) ;
  {$ENDIF}
    end
  else
    begin
	{$IFDEF EAGLCLIENT}
  	TOBListe := TOB.Create('Liste Ecriture',nil,-1) ;
    TOBLigne := Q.TQ.Detail[FListe.row-1] ;
    TOB.Create('ECRITURE',TOBListe,-1) ;
    TOBListe.Detail[0].Dupliquer(TOBLigne, False, True);
		TrouveEtLanceSaisie(TOBListe, taConsult, 'N' ) ;
    FreeAndNil(TOBListe) ;
	{$ELSE}
		TrouveEtLanceSaisie(Q, taConsult, 'N' ) ;
	{$ENDIF}
    end ;
end ;

procedure TOF_MULRELDETAIL.FListeDblClick(Sender: TObject);
begin
  BZoomPieceClick(Nil) ;
end;

procedure TOF_MULRELDETAIL.BOuvrirClick(Sender: TObject);
begin
  if Not JeValide then Exit ;
  if Not VerifCoher then Exit ;
  if HRel.Execute(1,'','')<>mrYes then Exit ;
  NowFutur := NowH ;
  FabricReleve ;
  {Ré-init}
//  DernLettre := '' ;
  TFMul(Ecran).BChercheClick(Nil) ;
end ;

function  TOF_MULRELDETAIL.JeValide : boolean ;
Var Q,QP    : TQuery ;
    JalRel  : String ;
    MM      : String17 ;
    NumP    : Longint ;
begin
  Result := False ;
  // Vérification au moins 1 ligne sélectionnée
  if not( ( FListe.NbSelected >0 ) or (FListe.AllSelected) ) then
    begin
    HRel.Execute(0,'','') ;
    Exit ;
    end ;

  // Vérification mode révision
  if RevisionActive(DateGene) then Exit ;

  // Récupération Souche de relevé
  Q := OpenSQL('Select * from SOUCHE Where SH_TYPE="REL" and SH_JOURNAL<>""',False) ;
  if Q.EOF then
    begin
    HRel.Execute(2,'','') ;
    Ferme(Q) ;
    Exit ;
    end
  else
    begin
    JalRel        := Q.FindField('SH_JOURNAL').AsString ;
    GeneNumRel    := Q.FindField('SH_NUMDEPART').AsInteger ;
    GeneMasqueRel := Q.FindField('SH_MASQUENUM').AsString ;
    // Inc Num compteur ref relevé
    ExecuteSQL('UPDATE SOUCHE SET SH_NUMDEPART = ' + IntToStr( GeneNumRel+1 ) +
                ' WHERE SH_TYPE="' + Q.FindField('SH_TYPE').AsString + '"' +
                ' AND SH_SOUCHE="' + Q.FindField('SH_SOUCHE').AsString + '"' ) ;
    end ;
  Ferme(Q) ;

  // Récupération Souche du Journal de la Souche de relevé
  QP := OpenSQL('SELECT J_COMPTEURNORMAL, J_MULTIDEVISE FROM JOURNAL ' +
                'WHERE J_JOURNAL="' + JalRel + '" AND J_COMPTEURNORMAL<>""', True) ;
  if QP.EOF then
    begin
    HRel.Execute(5,'','') ;
    Ferme(QP) ;
    Exit ;
    end
  else
    begin
    GeneSoucheP :=  QP.FindField('J_COMPTEURNORMAL').AsString ;
    if GeneSoucheP='' then
      begin
      Ferme(QP) ;
      Exit ;
      end ;
    if ((QP.FindField('J_MULTIDEVISE').AsString='-') and (DEV.Code<>V_PGI.DevisePivot)) then
      begin
      Ferme(QP) ;
      HRel.Execute(13,'','') ;
      Exit ;
      end ;
    NumP := GetNum(EcrGen,GeneSoucheP,MM,0) ;
    end ;
  Ferme(QP) ;
  Result:=True ;

  // Remplissage structure M
  FillChar( LeM, Sizeof(LeM), #0 ) ;
  LeM.Jal         := JalRel ;
  LeM.Nature      := 'OD' ;
  LeM.Num         := NumP ;
  LeM.Exo         := QuelExoDT(DateGene) ;
  LeM.DateC       := DateGene ;
  LeM.CodeD       := DEV.Code ;
  LeM.TauxD       := DEV.Taux ;
  LeM.DateTaux    := DEV.DateTaux ;
  LeM.Valide      := False ;
  LeM.EtapeRegle  := False ;
  LeM.FromGuide   := False ;
  LeM.Simul       := 'N' ;
  LeM.Etabl       := GetControlText('E_ETABLISSEMENT') ;
  if LeM.Etabl='' then
    LeM.Etabl     := VH^.EtablisDefaut ;
end;

function  TOF_MULRELDETAIL.VerifCoher : boolean ;
Var i         : integer ;
    DD        : TDateTime ;
    Eche      : TDateTime ;
    Valeu     : TDateTime ;
    MP        : String ;
    X         : T_MONOECH ;
    Premier   : boolean ;
    Coher     : boolean ;
begin
  Result    := False ;
  Premier   := True ;
  Coher     := True ;
  DateMin   := IDate2099 ;
  DateMax   := IDate1900 ;

{$IFDEF EAGLCLIENT}
  // Pas de mode "Tout sélectionné en CWAS
{$ELSE}
  if FListe.AllSelected then
    begin
    Q.First ;
    While not Q.Eof do
      begin
      DD      := Q.FindField('E_DATECOMPTABLE').asDateTime ;
      MP      := Q.FindField('E_MODEPAIE').AsString ;
      Eche    := Q.FindField('E_DATEECHEANCE').asDateTime ;
      Valeu   := Q.FindField('E_DATEVALEUR').asDateTime ;
      if Valeu <= Encodedate(1901,01,01)
        then Valeu := Eche ;
      if Premier then
        begin
        LeMode    := MP ;
        LaEche    := Eche ;
        LaValeur  := Valeu ;
        end
      else
        if ((Eche<>LaEche) or (MP<>LeMode))
          then Coher:=False ;
      if DD>DateMax
        then DateMax:=DD ;
      if DD<DateMin
        then DateMin:=DD ;
      Premier:=False ;

      Q.Next ;
      end ;
    end
  else	// Sinon, on parcours le Bookmark
{$ENDIF}
    for i := 0 to (FListe.NbSelected - 1) do
      begin
      FListe.GotoLeBookmark(i) ;
      {$IFDEF EAGLCLIENT}
        Q.TQ.Seek(FListe.row - 1) ;
      {$ENDIF}
      DD      := Q.FindField('E_DATECOMPTABLE').asDateTime ;
      MP      := Q.FindField('E_MODEPAIE').AsString ;
      Eche    := Q.FindField('E_DATEECHEANCE').asDateTime ;
      Valeu   := Q.FindField('E_DATEVALEUR').asDateTime ;
      if Valeu <= Encodedate(1901,01,01)
        then Valeu := Eche ;
      if Premier then
        begin
        LeMode    := MP ;
        LaEche    := Eche ;
        LaValeur  := Valeu ;
        end
      else
        if ((Eche<>LaEche) or (MP<>LeMode))
          then Coher:=False ;
      if DD>DateMax
        then DateMax:=DD ;
      if DD<DateMin
        then DateMin:=DD ;
      Premier:=False ;
      end ;

  if DateGene > DateMax
    then DateMax := DateGene ;
  if DateGene < DateMin
    then DateMin := DateGene ;
  if Not Coher then
    begin
    if HRel.Execute(6,'','')<>mrYes then Exit ;
    X.DateEche    := LaEche ;
    X.ModePaie    := LeMode ;
    X.Treso       := False ;
    X.Cat         := '' ;
    X.OkInit      := False ;
    X.OkVal       := False ;
    X.Action      := taModif ;
    X.DateValeur  := LaValeur ;
    X.DateMvt     := 0 ;
    if Not SaisirMonoEcheance(X) then Exit ;
    LaEche        := X.DateEche ;
    LeMode        := X.ModePaie ;
    LaValeur      := X.DateValeur ;
    END ;

  Result := True ;
end;


procedure TOF_MULRELDETAIL.FabricReleve ;
Var TDD,TCD,TDP,TCP : Double ;
    CodeL : String ;
    SQ    : AnsiString ;
    LL    : TStrings ;
    TT    : TList ;
    i     : integer ;
begin
//  Application.ProcessMessages ; // ??
  InitMove(3,'') ;
  // MAJ Tiers
  MoveCur(False) ;
  CodeL := MajTiersRelLet( GetControlText('E_AUXILIAIRE') ) ;
  if CodeL='' then
    begin
    V_PGI.IoError := oeUnknown ;
    Exit ;
    end ;
  // Génération pièce comptable
  REL.General     := GetControlText('E_GENERAL') ;
  REL.Auxiliaire  := GetControlText('E_AUXILIAIRE') ;
  REL.DebitD      := 0 ;
  REL.CreditD     := 0 ;
  REL.DebitP      := 0 ;
  REL.CreditP     := 0 ;
//  REL.DebitE      := 0 ;
//  REL.CreditE     := 0 ;
  REL.ModePaie    := LeMode ;
  REL.Echeance    := LaEche ;
  REL.DateMin     := DateMin ;
  REL.DateMax     := DateMax ;
  REL.CodeL       := CodeL ;

  CalculCreditDebit( TDD, TCD, TDP, TCP) ;

  if (TDD - TCD) >= 0 then
    begin
    REL.DebitD  := (TDD - TCD) ;
    REL.DebitP  := (TDP - TCP) ;
    if REL.DebitD <> 0
      then LeM.TauxD := (REL.DebitP * DEV.Quotite) / REL.DebitD ;
    end
  else
    begin
    REL.CreditD :=  (TCD - TDD) ;
    REL.CreditP :=  (TCP - TDP) ;
    if REL.CreditD<>0
      then LeM.TauxD:=(REL.CreditP*DEV.Quotite)/REL.CreditD ;
    end ;

  if VH^.OuiTvaEnc then
    for i:=1 to 5 do
      REL.TabTvaEnc[i] := TabTvaEnc[i] ;

  REL.RegimeTva := RegimeTva ;
  REL.RefRel    := Trim(GeneMasqueRel) + ' ' + Inttostr(GeneNumRel) ;
  MoveCur(False) ;
  GenerePieceReleve( REL, LeM, DEV) ;

  // Mises à jour
  MoveCur(False) ;
  SetIncNum( EcrGen, GeneSoucheP, LeM.Num, 0 ) ;
  LettrageReleveOrigine( CodeL, REL ) ;
  FiniMove ;
  if V_PGI.IOError<>oeOK then Exit ;

  // Edition relevé

  SQ := RecupSQLModele('L','RLV','RLW',CodeL,REL.Auxiliaire,REL.General) ;
  TT := TList.Create ;
  LL := TStringList.Create ;
  try
     LL.Add(SQ) ;
     TT.Add(LL) ;
     {JP 23/02/04 : L'état 2/3 utilisait des champs spéciaux définis en dur dans l'AGL. Pour
                    l'état eAgl, ces variables ont été remplacées par les champs de la table
                    ÉCRITURE, d'où le nouvel état}
     LanceDocument('L', 'RLV', 'RLW',                      // Document à éditer
                     TT,                                   // TList des requêtes
                     nil,                                  // non géré en eAGL
                     GetCheckBoxState('CApp')=cbChecked,   // Aperçu avant impression
                     True ) ;
  finally
     LL.Free ;
     TT.Free ;
  end;
  if ((YATP) and (GetControlText('MODELETP')<>'')) then
    if HRel.Execute( 15, gszCaption, '' )=mrYes then
      EditeRelevesTP ;
  YATP := False ;
end;

procedure TOF_MULRELDETAIL.MajNavigation ;
begin
{$IFDEF EAGLCLIENT}
  SetControlEnabled('BPRECEDENT', FIndex > 0) ;
  SetControlEnabled('BSUIVANT',   FIndex < (FTiers.Detail.Count-1) ) ;
{$ELSE}
  SetControlEnabled('BPRECEDENT', not FTiers.Bof) ;
  SetControlEnabled('BSUIVANT',   not FTiers.Eof) ;
{$ENDIF}
end;

procedure TOF_MULRELDETAIL.LettrageReleveOrigine ( CodeL : String4 ; REL : EC_RELF ) ;
Var i         : integer ;
begin
  YATP := False ;

{$IFDEF EAGLCLIENT}
  // Pas de mode "Tout sélectionné en CWAS
{$ELSE}
  if FListe.AllSelected then
    begin
    Q.First ;
    While not Q.Eof do
      begin
      MoveCur(False) ;
      // Recup et maj TOBM
      if not UpdateTOBMOrigine ( CodeL, REL ) then Exit ;
      Q.Next ;
      end ;
    end
  else	// Sinon, on parcours le Bookmark
{$ENDIF}
    for i := 0 to (FListe.NbSelected - 1) do
      begin
      // Déplacement grille
      FListe.GotoLeBookmark( i ) ;
      {$IFDEF EAGLCLIENT}
        Q.TQ.Seek(FListe.row - 1) ;
      {$ENDIF}
      MoveCur(False) ;
      // Recup et maj TOBM
      if not UpdateTOBMOrigine ( CodeL, REL ) then exit ;
      end ;

end ;

Function  TOF_MULRELDETAIL.MajTiersRelLet ( Auxi : String ) : String ;
Var CodeL      : String ;
    Q          : TQuery ;
    lStSQL     : String ;
begin
  CodeL := '' ;
  Q := OpenSQL('Select T_DATEDERNRELEVE, T_DERNLETTRAGE from TIERS Where T_AUXILIAIRE="'+Auxi+'" AND T_DERNLETTRAGE="'+DernLettre+'"', True) ;
  if Not Q.EOF then
    begin
    // Calcul code de dernier relevé
    CodeL := CodeSuivant( Q.FindField('T_DERNLETTRAGE').AsString ) ;
    DernLettre := CodeL ;
    // MAJ du tiers en base
    lStSQL := 'UPDATE TIERS SET T_DATEDERNRELEVE = "' + USDateTime(V_PGI.DateEntree) + '", ' +
                               'T_DERNLETTRAGE = "'   + CodeL + '" ' +
                         'WHERE T_AUXILIAIRE="'       + Auxi + '" ';
    ExecuteSQL( lStSQL ) ;
    END ;
  Ferme(Q) ;
  Result := CodeL ;
end ;

procedure TOF_MULRELDETAIL.EditeRelevesTP ;
Var Q  : TQuery ;
    St,SubSQ : String ;
    TT : Tlist ;
    GenePremier : boolean ;
    LL,ListAux : TStrings ;
    SQL,SQ,MinExo,MaxExo,LeExo,SWhere,Auxi,OldEtat,OldWhere2 : String ;
    MinDate,MaxDate,LaDate : TDateTime ;
    MinNum,MaxNum,LeNum,i : integer ;
begin
  OldEtat   := GetControlText('E_ETATLETTRAGE') ;
  OldWhere2 := GetControlText('XX_WHERE2') ;
  SetControlText('E_ETATLETTRAGE', '' ) ;
  SetControlText('XX_WHERE2',      '' ) ;
  St := 'SELECT E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_CONTREPARTIEAUX FROM ECRITURE '
         + RecupWhereCritere( TPageControl(GetControl('Pages', True)) )
         + ' AND E_MULTIPAIEMENT="TP" AND E_FLAGECR="ROR" AND E_QUALIFORIGINE="TP"' ;
  MinNum  := 999999999 ;
  MaxNum  := 0 ;
  MinDate := iDate2099 ;
  MaxDate := iDate1900 ;
  MinExo  := 'zzz' ;
  MaxExo  := '   ' ;
  ListAux := TStringList.Create ;

  Q := OpenSQL(St,True) ;
  While Not Q.EOF do
    begin
    // Exo
    LeExo := Q.FindField('E_EXERCICE').AsString ;
    if LeExo>MaxExo then MaxExo := LeExo ;
    if LeExo<MinExo then MinExo := LeExo ;
    // Date comptable
    LaDate := Q.FindField('E_DATECOMPTABLE').AsDateTime ;
    if LaDate>MaxDate then MaxDate := LaDate ;
    if LaDate<MinDate then MinDate := LaDate ;
    // Numéro de pièce
    LeNum := Q.FindField('E_NUMEROPIECE').AsInteger ;
    if LeNum>MaxNum then MaxNum := LeNum ;
    if LeNum<MinNum then MinNum := LeNum ;
    // Auxiliaire
    Auxi := Q.FindField('E_CONTREPARTIEAUX').AsString ;
    if ListAux.indexOf(Auxi) < 0 then ListAux.Add(Auxi) ;
    // MAJ Ecriture BASE
    SQL := 'UPDATE ECRITURE SET E_MULTIPAIEMENT="TF" ' +
           'WHERE E_JOURNAL = "' + Q.FindField('E_JOURNAL').AsString + '" AND ' +
                 'E_EXERCICE = "' + LeExo + '" AND ' +
                 'E_DATECOMPTABLE = "' + UsDateTime(LaDate) + '" AND ' +
                 'E_NUMEROPIECE = ' + IntToStr(LeNum) + ' AND ' +
                 'E_NUMLIGNE = 1 AND ' +
                 'E_MULTIPAIEMENT = "" ' ;
    ExecuteSQL(SQL) ;
    Q.Next ;
    end ;

  Ferme(Q) ;
  GenePremier := True ;
  SWhere := ' E_JOURNAL = "' + VH^.JalVTP + '" AND ' +
             'E_EXERCICE >= "' + MinExo + '" AND ' +
             'E_EXERCICE <= "' + MaxExo + '" AND ' +
             'E_DATECOMPTABLE >= "' + UsDateTime(MinDate) + '" AND ' +
             'E_DATECOMPTABLE <= "' + UsDateTime(MaxDate) + '" AND ' +
             'E_NUMEROPIECE >= ' + IntToStr(MinNum) + ' AND ' +
             'E_NUMEROPIECE <= ' + IntToStr(MaxNum) + ' AND ' +
             'E_NUMLIGNE = 1 ' ;

  for i:=0 to ListAux.Count-1 do
    begin
    SubSQ := ' AND ' + SWhere + ' AND E_MULTIPAIEMENT="TF" ' ;
    Auxi  := ListAux[i] ;
    TT    := TList.Create ;
    LL    := TStringList.Create ;
    SQ    := RecupSQLModele( 'L', 'RLC', GetControlText('MODELETP'), '', Auxi, '') ;
    LL.Add(SQ) ;
    LL.Add(SubSQ) ;
    TT.Add(LL) ;
    LanceDocument(  'L', 'RLC', GetControlText('ModeleTP'),
                    TT,
                    Nil,
                    GetCheckBoxState('CApp') = cbChecked,
                    GenePremier ) ;
    LL.Free ;
    TT.Free ;
    GenePremier := False ;
    end ;
  if ListAux.Count>0 then
    begin
    SWhere := ' E_JOURNAL = "' + VH^.JalVTP + '" AND ' +
               'E_EXERCICE >= "' + MinExo + ' " AND ' +
               'E_EXERCICE <= "' + MaxExo + '" AND ' +
               'E_DATECOMPTABLE >= "' + UsDateTime(MinDate) + '" AND ' +
               'E_DATECOMPTABLE <= "' + UsDateTime(MaxDate) + '" AND ' +
               'E_NUMEROPIECE >= ' + IntToStr(MinNum) + ' AND ' +
               'E_NUMEROPIECE <= ' + IntToStr(MaxNum) + ' AND ' +
               'E_MULTIPAIEMENT <> "" ' ;
    ExecuteSQL('UPDATE ECRITURE SET E_MULTIPAIEMENT="" WHERE ' + SWhere ) ;
    end ;
  ListAux.Clear ;
  ListAux.Free ;

  SetControlText('E_ETATLETTRAGE', OldEtat) ;
  SetControlText('XX_WHERE2', OldWhere2 ) ;

end ;

procedure TOF_MULRELDETAIL.DATEGENEREExit(Sender: TObject);
Var DD : TDateTime ;
    Err : integer ;
begin
  if Not IsValidDate(GetControlText('DATEGENERE')) then
    begin
    HRel.Execute(8,'','') ;
    SetControlText('DATEGENERE', DateToStr(V_PGI.DateEntree)) ;
    DateGene := V_PGI.DateEntree ;
    end
  else
    begin
    DD  := StrToDate( GetControlText('DATEGENERE') ) ;
    Err := DateCorrecte(DD) ;
    if Err>0 then
      begin
      HRel.Execute(8+Err,'','') ;
      SetControlText('DATEGENERE' , DateToStr(V_PGI.DateEntree) ) ;
      DateGene := V_PGI.DateEntree ;
      end
    else
      begin
      if RevisionActive(DD) then
        begin
        SetControlText('DATEGENERE', DateToStr(V_PGI.DateEntree) ) ;
        DateGene := V_PGI.DateEntree ;
        end
      else
        begin
        DateGene := DD ;
        end ;
      end ;
    end ;
end;



procedure TOF_MULRELDETAIL.CalculCreditDebit( var TDD,TCD,TDP,TCP : Double ) ;
Var i : integer ;
begin
  TDD := 0 ;
  TCD := 0 ;
  TDP := 0 ;
  TCP := 0 ;
  FillChar(TabTvaEnc,Sizeof(TabTvaEnc),#0) ;
{$IFDEF EAGLCLIENT}
  // Pas de mode "Tout sélectionné en CWAS
{$ELSE}
  if FListe.AllSelected then
    begin
    Q.First ;
    While not Q.Eof do
      begin
      CalculCreditDebitLigne( TDD, TCD, TDP, TCP) ;
      Q.Next ;
      end ;
    end
  else	// Sinon, on parcours le Bookmark
{$ENDIF}
    for i:=0 to FListe.NbSelected-1 do
      begin
      FListe.GotoLeBookmark(i) ;
{$IFDEF EAGLCLIENT}
      Q.TQ.Seek(FListe.row - 1) ;
{$ENDIF}
      CalculCreditDebitLigne( TDD, TCD, TDP, TCP) ;
      end ;
end;

procedure TOF_MULRELDETAIL.CalculCreditDebitLigne(var TDD, TCD, TDP, TCP: Double);
Var k   : integer;
begin
  TDD := TDD + Q.FindField('E_DEBITDEV').AsFloat ;
  TCD := TCD + Q.FindField('E_CREDITDEV').AsFloat ;
  TDP := TDP + Q.FindField('E_DEBIT').AsFloat ;
  TCP := TCP + Q.FindField('E_CREDIT').AsFloat ;
  if VH^.OuiTvaEnc then
    begin
    for k:=1 to 4 do
      TabTvaEnc[k] := TabTvaEnc[k] + Q.FindField('E_ECHEENC'+IntToStr(k)).AsFloat ;
    TabTvaEnc[5] := TabTvaEnc[5] + Q.FindField('E_ECHEDEBIT').AsFloat ;
    end ;
  RegimeTva := Q.FindField('E_REGIMETVA').AsString ;
end;

procedure TOF_MULRELDETAIL.RechercheDernCodeLett ;
Var Q : TQuery ;
begin
  DernLettre := '' ;
  Q := OpenSQL('Select T_DERNLETTRAGE from TIERS Where T_AUXILIAIRE="'+Auxi+'"', True) ;
  if Not Q.EOF then
    DernLettre := Q.FindField('T_DERNLETTRAGE').AsString ;
  Ferme(Q) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 09/05/2003
Modifié le ... : 13/06/2003
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_MULRELDETAIL.UpdateTOBMOrigine( CodeL : String4 ; REL : EC_RELF ) : Boolean ;
Var St,StUp   : String ;
    SQL,EMP   : String ;
    AncDate   : TDateTime ;
    M         : RMVT ;
    O         : TOBM ;
  {$IFDEF EAGLCLIENT}
    lTobEcr   : TOB ;
  {$ENDIF}
begin

  Result := False ;
  EMP := '' ;

  // Récupération TOBM
  O := TOBM.Create( EcrGen, '', True ) ;
  {$IFDEF EAGLCLIENT}
//    O.ChargeMvt( Q.TQ.Detail[FListe.Row - 1] );
    lTobEcr := Q.TQ.Detail[FListe.Row - 1] ;
    O.SelectDB('"' + lTobEcr.GetValue('E_JOURNAL') + '";' +
               '"' + lTobEcr.GetValue('E_EXERCICE') + '";' +
               '"' + USDateTime(lTobEcr.GetValue('E_DATECOMPTABLE')) + '";' +
               IntToStr(lTobEcr.GetValue('E_NUMEROPIECE')) + ';' +
               IntToStr(lTobEcr.GetValue('E_NUMLIGNE')) + ';' +
               IntToStr(lTobEcr.GetValue('E_NUMECHE')) + ';' +
               '"' + lTobEcr.GetValue('E_QUALIFPIECE') + '"', nil, False);

  {$ELSE}
    O.ChargeMvt( Q );
  {$ENDIF}
  if O=Nil then Exit ;

  // MAJ objet mémoire
  O.PutMvt('E_REFRELEVE',     REL.RefRel)  ;
  O.PutMvt('E_FLAGECR',       'ROR') ;
  O.PutMvt('E_LETTRAGE',      CodeL)        ;
  O.PutMvt('E_ETATLETTRAGE',  'TL') ;
  O.PutMvt('E_MODEPAIE',      LeMode)       ;
  O.PutMvt('E_DATEECHEANCE',  LaEche) ;
  O.PutMvt('E_DATEPAQUETMIN', DateMin) ;
  O.PutMvt('E_DATEPAQUETMAX', DateMax) ;
  if DEV.Code<>V_PGI.DevisePivot then
    begin
    O.PutMvt('E_LETTRAGEDEV', 'X') ;
//    O.PutMvt('E_LETTRAGEEURO','-') ;
    end
  else
    begin
    O.PutMvt('E_LETTRAGEDEV', '-') ;
{    if ModeOppose
      then O.PutMvt('E_LETTRAGEEURO','X')
      else O.PutMvt('E_LETTRAGEEURO','-') ;}
    end ;
  if ((VH^.OuiTP) and (O.GetMvt('E_QUALIFORIGINE')='TP') and (GetControlText('MODELETP')<>'')) then
    begin
    YATP:=True ;
    EMP:='TP' ;
    end ;
  O.PutMvt('E_COUVERTURE',      O.GetMvt('E_DEBIT') + O.GetMvt('E_CREDIT') ) ;
  O.PutMvt('E_COUVERTUREDEV',   O.GetMvt('E_DEBITDEV') + O.GetMvt('E_CREDITDEV') ) ;
//  O.PutMvt('E_COUVERTUREEURO',  O.GetMvt('E_DEBITEURO') + O.GetMvt('E_CREDITEURO') ) ;
  O.PutMvt('E_DATEVALEUR',      LaValeur ) ;
  O.PutMvt('E_MULTIPAIEMENT',   EMP ) ;
  AncDate := O.GetMvt('E_DATEMODIF') ;
  O.PutMvt('E_DATEMODIF', NowFutur) ;

  // Gestin de l'état // ???
  St:=O.GetMvt('E_ETAT') ;
  if Length(St)>=7 then
    begin
    St[1]:='-'; St[2]:='-'; St[3]:='-'; St[4]:='0'; St[5]:='R'; St[6]:='M'; St[7]:='X';
    // Lettrage Relevé, sur Montants, sans Regul, ni Ecartchange, ni Pasécart, Originerelevé
    O.PutMvt('E_ETAT',  St) ;
    end ;

  StUp := O.StPourUpdate ;
  if St='' then Exit ;

  // MAJ fichier
  M := OBMToIdent(O,True) ;

  SQL := 'UPDATE ECRITURE SET ' + StUp +
          ' Where  ' + WhereEcriture(tsGene,M,True) +
          ' AND E_DATEMODIF="' + USTime(AncDate) + '"' ;

  if ExecuteSQL(SQL)<>1 then
    begin
    V_PGI.IOError:=oeUnknown ;
    Exit ;
    end ;

  // Tout s'est bien passé
  Result := True ;
end;

Initialization
  registerclasses ( [ TOF_MULRELDETAIL ] ) ;
end.

