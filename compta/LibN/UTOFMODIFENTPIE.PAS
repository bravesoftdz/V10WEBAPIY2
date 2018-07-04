{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 19/09/2001
Modifié le ... : 19/10/2001
Description .. : Source TOF de la TABLE : MODIFENTPIE ()
Mots clefs ... : TOF;MODIFENTPIE
*****************************************************************}
Unit UTOFMODIFENTPIE ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     eMUL,       // MUL
     Maineagl,   // AGLLanceFiche
{$ELSE}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     MUL,        // MUL
     Fe_Main,    // AGLLanceFiche
{$ENDIF}
{$IFDEF VER150}
   Variants,
{$ENDIF}
     forms,      // Objet Application
     sysutils,   // AnsiUpperCase,StrToDateTime,EncodeDate,StrToInt,IntToStr,StrToDate,Date,DateToStr
     ComCtrls,   // TTabSheet,TPageControl
     HCtrls,     // Contrôles Halley
     HEnt1,      // EnableControls,TraduireMemoire,DebutDeMois,FinDeMois,V_PGI,Arrondi,NumSemaine
     HMsgBox,    // THMsgBox,PGIBox,PGIAsk,PGIInfo
     UTOF,       // TOF
     Ent1,       // VH,PositionneEtabUser
     UTOB,       // TOB
     Saisie,     // TrouveEtLanceSaisie
     Saisbor,    // LanceSaisieFolio
     HDB,        // THDBGrid
     SaisUtil,   // QuelExoDT,GetNewNumJal,EcrGen, RevisionActive
     AGLInit,    // TheTob
     ed_tools,   // Pour VideListe
     DelVisuE,   // Pour VisuPiecesGenere
     uLibEcriture, // WhereEcritureTob
     UTOFMODIFENTPIEPARAM  // CCLanceFiche_ModifEntPieParam
     ;
Const
  Titre = 'Modification entête de pièces' ;
  Dec = 2 ;

procedure CCLanceFiche_ModifEntPie(Range : string = ''; Arg : string = '');
{JP 21/03/07 : Suppression de pièces sur critères}
procedure ModifEntPieSurCriteres(Jal, Exo, Nat, Num, dtC : string); 

Type
  TOF_MODIFENTPIE = Class (TOF)
    Etablissement : THValComboBox ;
    Journal : THValComboBox ;
    NatPiece : THValComboBox ;
    Mois : THValComboBox ;
    OldAnnee : THValComboBox ;
    Exercice : THValComboBox ;
    DateComptaDe : THEdit;
    DateComptaA  : THEdit;
    LMsg : THMsgBox ;
    TobEcr : TOB ;
    TobAux : TOB ;
    TobJal : TOB ;
    TobPiece : TOB ;
    TobSecAna : TOB ;
    TobExercice : TOB ;
    TobGen : TOB ;
    TobCompl : TOB ;
{$IFDEF EAGLCLIENT}
    ListeEcr : THGrid;
{$ELSE}
    ListeEcr : THDBGrid;
{$ENDIF}
    QlisteEcr : TQuery ;
    PCritere : TTabSheet ;
    Pages : TpageControl ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    gbInsert,gbJalBOR,bAllSelected,bUseBouton : boolean;
    JDebit,JCredit,TDebit,TCredit : Double ;
    gszNEWJALPLUS, gszNATUREDATATYPE : String;
    gszNewJal, gszNewNatPie, gszNewEtab : String;
    gszNewDate, gszNewAnnee, gszNewmois : String;
    gbRBDATE, gbRBMOIS : Boolean;
    function  Incompatibilite(FermeNewJal, FermeJournal, MDNewjal, MDJournal, MSNewJal, MSJournal: String): Boolean ;
    function  VeriChgExo(var ChgExo : Boolean ; var Exo1, Exo2 : String) : Boolean ;
    procedure JournalOnChange (Sender : TObject) ;
    procedure ExerciceOnChange(Sender: TObject);
    procedure BOuvrirOnClick(Sender: TObject);
    //SG6 10/01/05 FQ 14227
    {$IFNDEF EAGLCLIENT}
    procedure BSelectAllClick(Sender : TObject);
    {$ENDIF}

    procedure RecupMultiSel(pbChgperiode : Boolean; pszPeriode : String);
    procedure PrechargeOrigines;
    procedure MarqueOrigine;
    procedure VireInutiles(pbChgperiode : Boolean; pszPeriode : String);
    procedure RecupEcritAna ;
    Procedure MajEcrAna (TobEcrdetail : TOB ; ZoneMaJ : String ; NouvelValeur : Variant ) ;
    Procedure RecalculCompteTreso (var TobEcrDet : TOB ; AncienCptTreso, NouvCptTreso, NatJournal  : String) ;
    procedure CalculSiChangeExo (TobEcrDet : TOB ; NouvCptTreso, Exo2 : String);
    procedure CreationTobs ;
    procedure CreationTobJournal(bAvecNew : Boolean) ;
    procedure Arret ;
    procedure ChangeEtablissement(var TobEcrDet: TOB);
    procedure ChangeJournal(var TobEcrDetail: TOB);
    procedure ChangeNaturePiece;
    procedure ChangeDateOuPeriode(Var TobEcrDetail : TOB ; ChgDate : Boolean ; szDay : String);
    procedure MajCumulTob(var TobDetail: Tob; TobEcrDet: Tob ; Signe : Integer ; Lettre : String ; Ana : Boolean);
    procedure LibereTobs;
    function  MajTobs : boolean;
    Procedure RecalculSolde (ChgExo, Journaux : Boolean ; Exo2, AncienCptTreso, NouvCptTreso : String );
    procedure FinTraitement;
    procedure CumulJournal (TobEcrDetail : TOB);
    procedure AjoutCptGenerauxTreso(AncienCptTreso, NouvCptTreso : String);
    procedure FListeDblClick(Sender: TObject);
    procedure FListeFlipSelection(Sender: TObject);  //SG6 23/11/04 FQ 14227
    function  AfficheMsg(num: integer; Av, Ap: string): Word;
//    function  UpdateSet ( NumL : integer) : String ;
    procedure FermeMsg;
    procedure InitMsg;
    function RecupWhereCritere(PC : TPageControl) : String;

  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPVersion,
  CPTypeCons,
  CPProcGen,
  ULibExercice,
  {$ENDIF MODENT1}
  {$IFDEF eAGLCLIENT}
  MenuOLX
  {$ELSE}
  MenuOLG
  {$ENDIF eAGLCLIENT}
  {$IFDEF TRSYNCHRO}
  , uLibTrSynchro
  {$ENDIF}
  , Constantes
  ;

{JP 29/07/05 : FQ 15124 : exclusion des pièces GC/GA/GI/GP
 JP 26/06/07 : FQ TRESO 10491 : exclusion des pièces issues de la Tréso}
const
  CST_WHERE1 = 'EX_ETATCPTA = "OUV" AND E_QUALIFPIECE = "N" AND E_CREERPAR <> "DET" AND E_VALIDE = "-" ' +
               'AND E_MODESAISIE <> "LIB" AND ((E_REFGESCOM = "") OR (E_REFGESCOM IS NULL)) ' +
               'AND (E_QUALIFORIGINE <> "' + QUALIFTRESO + '" OR E_QUALIFORIGINE IS NULL OR E_QUALIFORIGINE = "") ';

procedure CCLanceFiche_ModifEntPie(Range : string = ''; Arg : string = '');
begin
  AGLLanceFiche ('CP' , 'CPMODIFENTPIE', Range, '' , Arg);
end;

{JP 21/03/07 : Suppression de pièces sur critères
{---------------------------------------------------------------------------------------}
procedure ModifEntPieSurCriteres(Jal, Exo, Nat, Num, dtC : string);
{---------------------------------------------------------------------------------------}
var
  Range : string;
begin
  Range := 'E_JOURNAL=' + Jal + ';E_EXERCICE=' + Exo + ';E_NATUREPIECE=' + Nat + ';E_DATECOMPTABLE=' +
            dtC  + ';E_DATECOMPTABLE_=' + dtC + ';';
  CCLanceFiche_ModifEntPie(Range);
end;

procedure TOF_MODIFENTPIE.OnLoad ;
var
  szSQL,szWhere : String;
  Q : TQuery ;
  i : integer;
BEGIN
  Inherited ;
  // Pour les pièces en mode bordereau
  // Si Journal avec mode saisie BORdereau
  if (ExisteSQL('SELECT J_MODESAISIE FROM JOURNAL WHERE J_MODESAISIE="BOR" AND J_JOURNAL="'+Journal.Value+'"')) then begin
    SetControlText('XX_WHERE','E_ETATLETTRAGE<="RI" AND E_REFPOINTAGE="" '); // 13533 // 14227
    szSQL := 'SELECT E_JOURNAL, E_NUMEROPIECE, MIN(E_NUMLIGNE), E_NUMGROUPEECR, E_PERIODE FROM ECRITURE LEFT OUTER JOIN EXERCICE ON E_EXERCICE=EX_EXERCICE ';
    szSQL := szSQL + AnsiUpperCase(RecupWhereCritere(Pages)) ;
    szSQL := szSQL + 'GROUP BY E_JOURNAL, E_NUMEROPIECE, E_NUMGROUPEECR, E_PERIODE';
    Q := OpenSql(szSQL,True);
    while not Q.EOF do begin
      szWhere := szWhere +'(E_NUMEROPIECE='+Q.Fields[1].AsString+' AND E_NUMLIGNE='+Q.Fields[2].AsString+' AND E_PERIODE='+Q.Fields[4].AsString+') OR ';
      Q.Next;
    end;
    Ferme(Q);
    i := length(szWhere)-3;
    szWhere := Copy(szWhere,0,i);
    SetControlText('XX_WHERE',szWhere);
    end
  else
    SetControlText('XX_WHERE','E_ETATLETTRAGE<="RI" AND E_NUMLIGNE=1 AND E_REFPOINTAGE="" '); // 13533 // 14227
end;

procedure TOF_MODIFENTPIE.OnArgument (S : String ) ;
BEGIN
  Inherited ;

  Ecran.HelpContext:=999999420;

  //SG6 10/01/05 FQ 14227
  bAllSelected := False;
  bUseBouton := False;

  {$IFDEF EAGLCLIENT}
    SetControlVisible('BSELECTALL',False);
  {$ELSE}
    Tbutton(GetControl('BSELECTALL',True)).OnClick := BSelectAllClick;
  {$ENDIF}


  gbInsert := False;
  Journal       := THValComboBox(GetControl('E_JOURNAL',True)) ;
  Exercice      := THValComboBox(GetControl('E_EXERCICE',True)) ;
  DateComptaDe	:= THEdit(GetControl('E_DATECOMPTABLE',True));
  DateComptaA	  := THEdit(GetControl('E_DATECOMPTABLE_',True));
  Etablissement := THValComboBox(GetControl('E_ETABLISSEMENT',True)) ;
  NatPiece      := THValComboBox(GetControl('E_NATUREPIECE',True)) ;
  PCritere      := TTabSheet(GetControl('PCRITERE',True)) ;
  Pages         := TpageControl(GetControl('PAGES',True)) ;

  If Journal <> NIL then Journal.OnChange := JournalOnChange ;
  If Exercice <> NIL then Exercice.OnChange := ExerciceOnChange;
  TButton(GetControl('BOUVRIR',True)).OnClick := BOuvrirOnClick ;

{$IFDEF EAGLCLIENT}
  THGrid(GetControl('FLISTE',True)).OnDblClick := FListeDblClick ;
{$ELSE}
  THDBGrid(GetControl('FLISTE',True)).OnDblClick := FListeDblClick ;
{$ENDIF}
  TFmul(Ecran).FListe.OnFlipSelection:= FListeFlipSelection ; //SG6 23/11/04 FQ 14227

  // Initialise les messages
  InitMsg;

  // Journal + Nature pièce + Etablissement + Exercice
//  SetControlProperty('E_JOURNAL','PLUS','AND J_MODESAISIE<>"LIB" AND (J_NATUREJAL="OD" OR J_NATUREJAL="ACH" OR J_NATUREJAL="VTE" OR J_NATUREJAL="CAI" OR J_NATUREJAL="BQE")') ;
// 11831 Enlève les journaux de type Bordereau
  SetControlProperty('E_JOURNAL','PLUS','AND J_MODESAISIE<>"LIB" AND J_MODESAISIE<>"BOR" AND (J_NATUREJAL="OD" OR J_NATUREJAL="ACH" OR J_NATUREJAL="VTE" OR J_NATUREJAL="CAI" OR J_NATUREJAL="BQE")') ;
  Journal.ItemIndex := 0; JournalOnChange(nil);
  Etablissement.ItemIndex := 0; PositionneEtabUser (Etablissement, False);
  Exercice.ItemIndex := 0; ExerciceOnChange(nil);

  // Permet de prendre les écritures Normale, non DETruites, non Validées, non saisie en mode LIBre et qui oné étés saisies après la dernière clôture périodique
  {JP 29/07/05 : FQ 15124 et non issues de GC/GA/GI/GP}
  SetControlText('XX_WHERE1',CST_WHERE1 + 'AND E_DATECOMPTABLE>"'+UsDateTime(VH^.DateCloturePer)+'"');
END ;

procedure TOF_MODIFENTPIE.OnClose ;
BEGIN
  LibereTobs ;
  FermeMsg;
  Inherited ;
END ;

// Vérification si les journaux ont le même mode multidevises, le même mode de saisie et s'ils ne sont pas fermés.
function TOF_MODIFENTPIE.Incompatibilite(FermeNewJal,FermeJournal,MDNewjal,MDJournal, MSNewJal, MSJournal : String ) : Boolean ;
begin
Result := False ;
  //  Au moins un des 2 journaux est fermé.
  If (FermeNewJal = 'X') or (FermeJournal = 'X') then begin AfficheMsg(2,'',''); Result := True; Exit; end;

  // Les journaux n'ont pas le même état multidevises.
  If (MDNewJal <> MDJournal) then begin AfficheMsg(3,'',''); Result := True; Exit; end;

  // Les journaux n'ont pas le même mode de saisie.
  If (MSNewJal <> MSJournal) then begin AfficheMsg(4,'',''); Result := True; Exit; end;
end;

// Vérification du changement ou non d'exercice entre la(les) date(s) de(s) l'écriture(s) sélectionnée(s) et la nouvelle date
Function TOF_MODIFENTPIE.VeriChgExo (Var ChgExo : Boolean ; Var Exo1, Exo2 : string) : Boolean ;
var
  dtNewDate : TDatetime ;
  ExoOrigineEstEnCours, ExoOrigineEstSuivant, ExoDestinationEstEnCours, ExoDestinationEstSuivant : Boolean ;
  Q : TQuery ;
begin
  Result := True ;
  If gbRBDATE then dtNewDate:= StrToDateTime(gszNewDate)
               Else dtNewDate := EncodeDate(StrToInt(gszNewAnnee), StrToInt(gszNewmois),1) ;

  TobExercice := Tob.Create('_EXERCICE', Nil,-1) ;
  Q:=OpenSql('SELECT EX_EXERCICE, EX_DATEDEBUT, EX_DATEFIN, EX_ETATCPTA FROM EXERCICE WHERE EX_ETATCPTA="OUV" ORDER BY EX_DATEDEBUT', True) ;
  TobExercice.LoadDetailDB('EXERCICE','', '', Q, False, False) ;
  Ferme(Q) ;

  ExoOrigineEstEnCours := (Exercice.Value = TobExercice.Detail[0].GetValue('EX_EXERCICE'));
  if (TobExercice.Detail.Count = 1) then ExoOrigineEstSuivant := False
                                    else ExoOrigineEstSuivant := (Exercice.Value = TobExercice.Detail[1].GetValue('EX_EXERCICE'));

  ExoDestinationEstEnCours := (dtNewDate >= TobExercice.Detail[0].getValue('EX_DATEDEBUT')) and (dtNewDate <= TobExercice.detail[0].getValue('EX_DATEFIN')) ;
  if (TobExercice.Detail.Count = 1) then ExoDestinationEstSuivant := False
                                    else ExoDestinationEstSuivant := (dtNewDate >= TobExercice.Detail[1].getValue('EX_DATEDEBUT')) and (dtNewDate <= TobExercice.Detail[1].getValue('EX_DATEFIN')) ;
  ChgExo := ((ExoOrigineEstEnCours) and (ExoDestinationEstSuivant)) or ((ExoDestinationEstEnCours) and (ExoOrigineestSuivant)) ;

  If ExoOrigineEstEnCours then begin Exo1 := VH^.EnCours.Code ; Exo2 := VH^.EnCours.Code ; If ChgExo then Exo2 := VH^.Suivant.Code ; end
                          else begin Exo1 := VH^.Suivant.Code ; Exo2 := VH^.Suivant.Code ; If ChgExo then Exo2 := VH^.EnCours.Code ; end;

  // Les dates des exercices ne correspondent pas.
  If ((Not ExoOrigineEstEnCours) And (Not ExoOrigineEstSuivant)) or ((Not ExoDestinationEstEnCours) and (Not ExoDestinationEstSuivant))then begin AfficheMsg(1,'',''); Result := False; end;
END ;

// Modifie la NATURE lors du changement de JOURNAL
Procedure TOF_MODIFENTPIE.JournalOnChange (Sender : TObject) ;
Var RSql, NatJal : string ;
    Q : TQuery  ;
BEGIN
  RSql := 'SELECT J_NATUREJAL FROM JOURNAL WHERE J_JOURNAL="'+GetControlText('E_JOURNAL')+'"' ;
  Q :=OpenSql(RSql, True) ;
  NatJal := Q.Fields[0].AsString ;
  If NatJal <>'OD' then gszNEWJALPLUS := 'AND J_NATUREJAL="'+NatJal+'"'
                   Else gszNEWJALPLUS := 'AND J_NATUREJAL="OD" OR J_NATUREJAL="ACH" OR J_NATUREJAL="VTE" OR J_NATUREJAL="CAI" OR J_NATUREJAL="BQE"';
  Ferme(Q) ;
  SetControlProperty('E_NATUREPIECE','DATATYPE','TTNATUREPIECE') ; // Par défaut
  If (Natjal= 'BQE') or (NatJal= 'CAI') Then SetControlProperty('E_NATUREPIECE','DATATYPE','TTNATPIECEBANQUE') ;
  If (Natjal= 'ACH')                    Then SetControlProperty('E_NATUREPIECE','DATATYPE','TTNATPIECEACHAT') ;
  If (Natjal= 'VTE')                    Then SetControlProperty('E_NATUREPIECE','DATATYPE','TTNATPIECEVENTE') ;
  NatPiece.ItemIndex := 0 ;
END ;

// Traitement des écritures
Procedure TOF_MODIFENTPIE.BOuvrirOnClick (Sender : TObject) ;
Var
  ChgJal, ChgEtab, ChgNatPie, Chgdate, ChgExo, Chgperiode : boolean ;
  SoldeCpteTresoARecalculer, bIncJal, bExisteFolio, bCutOff : Boolean ;
  OldNumeropiece, NumeroPiece, NewNumeropiece : Integer ;  // Pour les journaux en mode non bordereau
  iNewNumPie, iPeriode, iOldPeriode, iNumLig, iOldNumGrp, iNumGrp, iNewNumGrp : Integer;  // Pour les journaux en mode BORdereau
  iOldNumPie, iNumPie, iOldNumLig,iNumeroGrp : Integer;  // Pour le journaux en mode BORdereau
  iNewNumPie2, iNumLig2, iNewNumGrp2 : Integer;
  Fiche : TFMul ;
  i,j,k : integer ;
  NatNewJal, NatJournal : String ;
  MsNewJal, MdNewJal, CpNewJal, MsJournal, MdJournal, CpJournal, FermeNewJal, FermeJournal, CptNewJal, CptJournal : String ;
  szModeSaisie, S, SY, SSect, Debit, Credit, szDate, Exo1, Exo2, AncienCptTreso, NouvCptTreso, szMtt : String ;
  TobMAJBase, TobEcrDet, TobComplDet : Tob ;
  DDeb, DFin : TDateTime ;
  TInfo, TDetruite, TobEcrBOR,TobEcrBORSource, TobEcrAnaBOR : TOB;
  ListePieces : TList ;
  QBor : TQuery;
  lTobTmp : Tob ;
  lInfo : TInfoEcriture ;
  lRecError : TRecError ;
  lMess : TMessageCompta ;
BEGIN
  Fiche := TFMul(Ecran) ;
  ListeEcr := Fiche.FListe ;
{$IFDEF EAGLCLIENT}
  QListeEcr := Fiche.Q.TQ;
{$ELSE}
  QListeEcr := Fiche.Q ;
{$ENDIF}
  TDetruite:= nil;
  ListePieces:=nil;

  // Si pas de sélection
  If (not ListeEcr.AllSelected) and (ListeEcr.NbSelected=0) then begin
    PgiBox (TraduireMemoire('Aucune ligne sélectionnée. Traitement impossible.'), TraduireMemoire(Titre));
    Exit;
  end;

  // Les critères ne sont pas appliqués.
  if TfMul(Ecran).CritModified then begin
    PgiBox (TraduireMemoire('Vous n''avez pas appliqué les critères.'), TraduireMemoire(Titre));
    Exit;
  end;

  EnableControls(Ecran, False );

  // Vérifie si le mode de saisie du journal est BORdereau
  if (ExisteSQL('SELECT J_MODESAISIE FROM JOURNAL WHERE J_MODESAISIE="BOR" AND J_JOURNAL="'+QListeEcr.FindField('E_JOURNAL').AsString+'"') = True) then begin
    szModeSaisie := 'BOR';
    gbJalBOR := True;
    end
  else begin
    szModeSaisie := '-';
    gbJalBOR := False;
  end;

  TInfo := TOB.Create('', nil, -1);
  TInfo.AddChampSupValeur('ETABLISSEMENT',Etablissement.Value);
  TInfo.AddChampSupValeur('JOURNAL',Journal.Value);
  TInfo.AddChampSupValeur('NATURE',NatPiece.Value);
  TInfo.AddChampSupValeur('DATE','');
  TInfo.AddChampSupValeur('RBMOIS','');
  TInfo.AddChampSupValeur('RBDATE','');
  TInfo.AddChampSupValeur('NEWJALPLUS',gszNEWJALPLUS);
  TInfo.AddChampSupValeur('NATUREDATATYPE',gszNATUREDATATYPE);
  TInfo.AddChampSupValeur('MODESAISIE',szModeSaisie);

  // FQ 12038
  TInfo.AddChampSupValeur('BLOQNATURE','-');
  for i:=0 to ListeEcr.NbSelected-1 do
  begin
    ListeEcr.GotoLeBookmark(i) ;
    {$IFDEF EAGLCLIENT}
      Fiche.Q.TQ.Seek(Fiche.FListe.Row-1);
    {$ENDIF}
    if (QListeEcr.FindField('E_NATUREPIECE').AsString = 'OD') then
    begin
      TInfo.PutValue('BLOQNATURE','X');
      Break;
    end;
  end;
  TheTOB := TInfo;
  CCLanceFiche_ModifEntPieParam;
  TInfo.Free;
  if TheTOB = nil then begin EnableControls(Ecran, True ); exit; end;
  gszNewEtab := TheTOB.GetValue('ETABLISSEMENT');
  gszNewJal := TheTOB.GetValue('JOURNAL');
  gszNewNatPie := TheTOB.GetValue('NATURE');
  gszNewMois := TheTOB.GetValue('MOIS');
  if (TheTOB.GetValue('MOIS')<>'') then
    if (TheTOB.GetValue('MOIS')<10) then gszNewMois := '0'+gszNewMois;
  gszNewAnnee := TheTOB.GetValue('ANNEE');
  gszNewDate := TheTOB.GetValue('DATE');
  gbRBMOIS := TheTOB.GetValue('RBMOIS');
  gbRBDATE := TheTOB.GetValue('RBDATE');

  Chgjal := (Journal.Value <> gszNewJal) and (gszNewJal<>'');
  ChgEtab := (Etablissement.Value <> gszNewEtab) and (gszNewEtab<>'');;
  ChgNatPie := (NatPiece.Value <> gszNewNatPie) and (gszNewNatPie<>'');;
  ChgExo := False ; Exo1:='000' ; Exo2:='000' ;
  ChgDate := (gbRBDATE=True) and (gszNewDate <> '01/01/1900');
  Chgperiode := ((gbRBMOIS=True) and (gszNewMois <> '') and (gszNewAnnee <> '')) or ChgDate;
  if Chgperiode then if not VeriChgExo(ChgExo, Exo1, Exo2) then BEGIN Arret ; exit ; end ;
  If (not ChgJal) and (Not ChgEtab) and (Not ChgNatPie) and (Not Chgperiode) then BEGIN Arret ; Exit ; end ;

  If PGIAsk(TraduireMemoire('Confirmez-vous le traitement ?'),TraduireMemoire(Titre))=MrNo then begin Arret; Exit; end;

  //Creation TOB JOURNAL
  SoldeCpteTresoARecalculer := False;
  if (Journal.Value <> '') then begin
    CreationTobJournal(True);
    If TobJal.detail[0].GetValue('J_JOURNAL')=gszNewJal then BEGIN i:=0 ; j:=1 ; end
                                                        else BEGIN i:=1 ; j:=0 ; end;
    If not ChgJal then BEGIN i:=0 ; j:= 0 ; end ;
    NatNewJal := TobJal.Detail[i].GetValue('J_NATUREJAL');
    MsNewJal := TobJal.Detail[i].GetValue('J_MODESAISIE');
    MdNewJal := TobJal.Detail[i].GetValue('J_MULTIDEVISE');
    CpNewJal := TobJal.Detail[i].GetValue('J_CONTREPARTIE');
    CptNewJal := TobJal.Detail[i].GetValue('J_COMPTEURNORMAL');
    FermeNewJal := TobJal.Detail[i].GetValue('J_FERME') ;
    Natjournal := TobJal.Detail[j].GetValue('J_NATUREJAL') ;
    MsJournal := TobJal.Detail[j].GetValue('J_MODESAISIE') ;
    MdJournal := TobJal.Detail[j].GetValue('J_MULTIDEVISE') ;
    CpJournal := TobJal.Detail[j].GetValue('J_CONTREPARTIE') ;
    CptJournal := TobJal.Detail[j].GetValue('J_COMPTEURNORMAL');
    FermeJournal := TobJal.Detail[j].GetValue('J_FERME') ;
    If Incompatibilite(FermeNewJal,FermeJournal,MdNewjal,MdJournal, MsNewJal, MsJournal) then BEGIN Arret ; Exit ; end ;
    SoldeCpteTresoARecalculer := ((NatNewJal='BQE') or (NatNewJal='CAI')) ;
    If SoldeCpteTresoARecalculer then BEGIN AncienCptTreso := CpJournal ; NouvCptTreso := CpNewJal ; End ;
  end;

  RecupMultisel(Chgperiode,gszNewAnnee+gszNewMois); //Création TOB ECRITURE (+ analytiques correspondantes)
  if (TobEcr.Detail.Count <=0) then begin Arret; Exit ; end ;

  If ChgExo then CreationTobs ;
  If SoldeCpteTresoARecalculer then AjoutCptGenerauxTreso (AncienCptTreso, NouvCptTreso) ;

  // Changement d'établissement, de journal, de mois, d'exercice ou de date
  Try
    BeginTrans ;
    {JP 17/11/04 : FQ 14972 : on détruit les écriture de Trésorerie correspondant aux
                   écritures que l'on va supprimer}
    if EstComptaTreso then
      for i := 0 to TobEcr.Detail.Count - 1 do
        DetruitEcritureTresoTOB(TobEcr.Detail[i]);

    // Supprime les écritures de la base
    for i := 0 to TobEcr.Detail.Count-1 do
      begin
      // Ecriture
      ExecuteSQL('DELETE FROM ECRITURE WHERE ' + WhereEcritureTob( tsGene, TobEcr.Detail[i], True, False ) ) ;
      // S'il y a de l'analytique : Supprime l'analytique
      if ( TobEcr.Detail[i].GetString('E_ANA') = 'X' ) then
        ExecuteSQL('DELETE FROM ANALYTIQ WHERE ' + WhereEcritureTob( tsAnal, TobEcr.Detail[i], True, False ) ) ;
      end;

    // Si mode BORdereau
    if gbJalBOR then begin
      // Permet de savoir s'il existe un folio sur la période de destination : A faire impérativement AVANT la suppression des écritures (par la ligne suivante)
      if (ExisteSQL('SELECT E_NUMEROPIECE FROM ECRITURE WHERE E_JOURNAL="'+Journal.Value+'" AND E_PERIODE='+gszNewAnnee+gszNewMois)) then bExisteFolio := True else bExisteFolio := False;

      // Recherche le N° folio + n° groupe écriture de l'écriture de destination
      QBor := OpenSQL('SELECT MAX(E_NUMEROPIECE) FROM ECRITURE WHERE E_JOURNAL="'+Journal.Value+'" AND E_PERIODE='+gszNewAnnee+gszNewMois,True);
      iNewNumPie2:=QBor.Fields[0].AsInteger;
      Ferme(QBor);
      QBor := OpenSQL('SELECT MAX(E_NUMLIGNE), MAX(E_NUMGROUPEECR) FROM ECRITURE WHERE E_JOURNAL="'+Journal.Value+'" AND E_PERIODE='+gszNewAnnee+gszNewMois+' AND E_NUMEROPIECE='+IntToStr(iNewNumPie2),True);
      iNumLig2:=   QBor.Fields[0].AsInteger;
      iNewNumGrp2:=QBor.Fields[1].AsInteger;
      Ferme(QBor);

      // Duplique la tob et supprime les écritures de la base
      TobEcrBOR := Tob.Create('_ECRBOR',NIL,-1) ;
      TobEcrBOR.Dupliquer(TobEcr,True,True,False);
      TOBEcr.DeleteDB(True);

      // Si on a fait un changement de date
      if Chgperiode then begin
        // Permet de savoir si on change de folio (J_INCNUM = X) sinon on met l'écriture si le dernier folio
        if (ExisteSQL('SELECT * FROM JOURNAL WHERE J_JOURNAL="'+Journal.Value+'" AND J_INCNUM="X"')= True) then bIncJal := True else bIncJal := False;

        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // Rénumérote la TOB qu'on souhaite modifier
        iOldPeriode:=0; iOldNumGrp:=0;
        // Récupère le N° folio + n° groupe écriture + N° de ligne
        if (bExisteFolio) then begin
          if (bIncJal) then begin
            iNewNumPie:=iNewNumPie2+1;
            iNumLig   :=0;
            iNewNumGrp:=1;
            end
          else begin
            iNewNumPie:=iNewNumPie2;
            iNumLig   :=iNumLig2;
            iNewNumGrp:=iNewNumGrp2+1;
          end;
        end
        else begin
          iNewNumPie:=1;
          iNumLig   :=0;
          iNewNumGrp:=1;
        end;

        // Rénumérote la TobEcr (n° pièce + n° ligne + n° Groupe écriture)
        for i :=0 to TobEcr.Detail.Count-1 do begin
          // Si on est déja dans la période : passe au suivant
          if (TobEcr.Detail[i].GetValue('E_PERIODE')=gszNewAnnee+gszNewMois) then continue;
          if (i = 0) then begin
            iOldNumGrp := TobEcr.Detail[i].GetValue('E_NUMGROUPEECR');
            iOldPeriode:= TobEcr.Detail[i].GetValue('E_PERIODE');
          end;
          // Récupère les valeurs
          iNumLig := iNumLig+1;
          iNumGrp := TobEcr.Detail[i].GetValue('E_NUMGROUPEECR');
          iPeriode:= TobEcr.Detail[i].GetValue('E_PERIODE');

          // Affecte les valeurs
          TobEcr.Detail[i].PutValue('E_NUMEROPIECE',iNewNumPie);  // N° folio
          TobEcr.Detail[i].PutValue('E_NUMLIGNE',iNumLig);        // N° de ligne

          // Si la ligne de détail a de l'analytique : Met à jour ces dernières
          if (TobEcr.Detail[i].Detail.Count >0) then begin
            for j:=0 to TobEcr.Detail[i].Detail.Count-1 do begin
              TobEcr.Detail[i].Detail[j].PutValue('Y_NUMEROPIECE',iNewNumPie);
              TobEcr.Detail[i].Detail[j].PutValue('Y_NUMLIGNE',iNumLig);
            end;
          end;

          // Numéro de groupe d'écriture
          if ((iOldNumGrp<>iNumGrp) or (iOldPeriode<>iPeriode)) then begin
            iNewNumGrp:=iNewNumGrp+1;
            iOldNumGrp:=iNumGrp;
            iOldPeriode:=iPeriode;
          end;
          TobEcr.Detail[i].PutValue('E_NUMGROUPEECR',iNewNumGrp);
        end;
        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // Rénumérote les TOB sources (celle où on a pris les écritures) car le n° groupeecr,n° pièce et n°ligne peuvent ne plus être correct
        // Cas où on a pris une écriture au milieu d'un groupe pour la mettre ailleurs, tous ce qui est après doit être rénuméroté

        // Garde uniquement la première ligne de chaque écriture
        iOldPeriode := TobEcrBOR.Detail[TOBEcrBOR.Detail.Count-1].GetValue('E_PERIODE');
        iOldNumGrp  := TobEcrBOR.Detail[TOBEcrBOR.Detail.Count-1].GetValue('E_NUMGROUPEECR');
        i := 0;
        while (True) do begin
          iPeriode := TobEcrBOR.Detail[i].GetValue('E_PERIODE');
          iNumGrp  := TobEcrBOR.Detail[i].GetValue('E_NUMGROUPEECR');
          if (iOldNumGrp=iNumGrp) and (iOldPeriode=iPeriode) then TobEcrBOR.Detail[i].Free
                                                             else begin iOldPeriode := iPeriode; iOldNumGrp := iNumGrp; Inc(i); end;
          if (i>=TOBEcrBOR.Detail.Count) then break;
        end;

        // Modifie le groupe d'écriture auquelle appartient chaque ligne d'écriture
        TobEcrBORSource := Tob.Create('_ECRBOR2',NIL,-1);
        for i:=0 to TobEcrBOR.Detail.Count-1 do begin
          // Charge les écritures du bordereau
          if TobEcrBorSource <> nil then TobEcrBorSource.ClearDetail;
          S := 'SELECT * FROM ECRITURE WHERE E_JOURNAL="'+Journal.Value+'" AND E_EXERCICE="'+Exercice.Value+'" AND E_DATECOMPTABLE="'+USDateTime(TobEcrBOR.Detail[i].GetValue('E_DATECOMPTABLE'))+'" AND E_NUMEROPIECE='+IntToStr(TobEcrBOR.Detail[i].GetValue('E_NUMEROPIECE'))+' ORDER BY E_NUMLIGNE';
          QBor := OpenSQL(S, True);
          If not QBor.Eof then
            TobEcrBorSource.LoadDetailDB('_ECRBOR2','','',QBor,False,False);
          Ferme(QBor);

          // Charge l'analytique correspondant
          S := 'SELECT * FROM ANALYTIQ WHERE Y_JOURNAL="'+Journal.Value+'" AND Y_EXERCICE="'+Exercice.Value+'" AND Y_DATECOMPTABLE="'+USDateTime(TobEcrBOR.Detail[i].GetValue('E_DATECOMPTABLE'))+'" AND Y_NUMEROPIECE='+IntToStr(TobEcrBOR.Detail[i].GetValue('E_NUMEROPIECE'))+' ORDER BY Y_NUMLIGNE';
          QBor:=OpenSql(S,True) ;
          If not QBor.Eof then begin
            TobEcrAnaBOR := TOB.Create('_ANALY2', Nil, -1) ;
            TobEcrAnaBOR.LoadDetailDB('ANALYTIQ', '', '', QBor, True, True) ;
            //Changement de Parent pour la TOB analytique
            For j := 0 to TobEcrBorSource.Detail.Count -1 do begin
              If (TobEcrBorSource.Detail[j].GetValue('E_ANA')='X') then begin
                While (TobEcrAnaBOR.Detail.Count > 0)
                  And (TobEcrBorSource.Detail[j].GetValue('E_NUMEROPIECE')=TobEcrAnaBOR.Detail[0].GetValue('Y_NUMEROPIECE'))
                  And (TobEcrBorSource.Detail[j].GetValue('E_NUMLIGNE')=TobEcrAnaBOR.Detail[0].GetValue('Y_NUMLIGNE'))
                  And (TobEcrBorSource.Detail[j].GetValue('E_JOURNAL')=TobEcrAnaBOR.Detail[0].GetValue('Y_JOURNAL'))
                  And (TobEcrBorSource.Detail[j].GetValue('E_QUALIFPIECE')=TobEcrAnaBOR.Detail[0].GetValue('Y_QUALIFPIECE'))
                  And (TobEcrBorSource.Detail[j].GetValue('E_GENERAL')=TobEcrAnaBOR.Detail[0].GetValue('Y_GENERAL')) do begin
                    TobEcrAnaBOR.Detail[0].ChangeParent(TobEcrBorSource.Detail[j], -1) ;
                end;
              end;
            end;
            TobEcrAnaBOR.free ;
          end;
          if QBor <> Nil Then Ferme(QBor);

          // Supprime le folio de la base
          TobEcrBorSource.DeleteDB(True);

          iNumLig := 0; iNumeroGrp:=0;
          for k:=0 to TobEcrBorSource.Detail.Count-1 do begin
            if (k = 0) then begin
              iOldNumGrp := TobEcrBorSource.Detail[k].GetValue('E_NUMGROUPEECR');
            end;
            // Récupère les valeurs
            iNumLig := iNumLig+1;
            iNumGrp := TobEcrBorSource.Detail[k].GetValue('E_NUMGROUPEECR');

            // Numéro de groupe d'écriture
            if ((iOldNumGrp<>iNumGrp) or (iNumeroGrp=0)) then begin iNumeroGrp:=iNumeroGrp+1; iOldNumGrp:=iNumGrp; end;

            // Mise à jour
            TobEcrBorSource.Detail[k].PutValue('E_NUMLIGNE',iNumLig);
            TobEcrBorSource.Detail[k].PutValue('E_NUMGROUPEECR',iNumeroGrp);

            // Si la ligne de détail a de l'analytique : Met à jour ces dernières
            if (TobEcrBorSource.Detail[k].Detail.Count >0) then begin
              for j:=0 to TobEcrBorSource.Detail[k].Detail.Count-1 do begin
                TobEcrBorSource.Detail[k].Detail[j].PutValue('Y_NUMLIGNE',iNumLig);
              end;
            end;
          end;

          // Réécrit le folio dans la base (avec les bons n° de ligne)
          If TobEcrBorSource <> NIL then begin
            TobEcrBorSource.SetAllModifie(True);
            TobEcrBorSource.InsertOrUpdateDb ;
          end;
        end;
        // Libère la TOB
        If TobEcrBorSource <> NIL then begin TobEcrBorSource.Free; end;
        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
      end;  // if Chgperiode then begin
      // Libère la TOB
      If TobEcrBOR <> NIL then begin TobEcrBOR.Free; end;
    end; // if gbJalBOR then begin

    JDebit:=0 ; JCredit:=0 ; TDebit:=0 ; TCredit:=0 ;
    // Si le compteur du nouveau et de l'ancien journal sont différents : Duplique la tob Ecriture
    if (CptNewJal<>CptJournal) then begin
      ListePieces:=TList.Create;
      TDetruite := Tob.Create('ECR',NIL,-1) ;
      TDetruite.Dupliquer(TobEcr,True,True,True);
      {$IFDEF TRSYNCHRO}
      {JP 17/11/04 : FQ 14972 : Les écritures supprimées vont être recrées avec TobMAJBaseà 1 / -1
                     or ces écritures ne doivent pas être synchronisées}
      if EstComptaTreso then
        for i := 0 to TDetruite.Detail.Count - 1 do
          TDetruite.Detail[i].SetString('E_TRESOSYNCHRO', 'RIE');
      {$ENDIF}
    end;

    Numeropiece := 0;
    If ChgEtab or ChgJal or ChgPeriode then
    BEGIN
      TobCompl := Tob.create('ECRCOMPL',NIL,-1);

      For k:=0 TO TobEcr.Detail.Count-1 do BEGIN
        TobEcrDet := TobEcr.Detail[k];
        bCutOff:=False;
        {FQ21880 12.02.2008 YMO Changement de clef sur ECRCOMPL}
        if (ExisteSQL('SELECT G_CUTOFF FROM GENERAUX WHERE G_GENERAL="'+TobEcrDet.GetString('E_GENERAL')+'" AND G_CUTOFF="X"')) then
          bCutOff:=True;

        if bCutOff then
          TobComplDet := CSelectDBTOBCompl( TobEcrDet, TobCompl ) ; {FQ21880 05/12/2007 YMO Maj EcrCompl}

        if (CptNewJal<>CptJournal) then begin
          if (k>0) then begin
            if (TDetruite.Detail[k].GetValue('E_NUMEROPIECE') <> TDetruite.Detail[k-1].GetValue('E_NUMEROPIECE')) then
              ListePieces.add(TDetruite.Detail[k]);
            end
          else ListePieces.add(TDetruite.Detail[k]);
        end;

        If ChgJal then BEGIN
          // Si on est en mode BORdereau, on ne passe JAMAIS par ici (car on ne peut pas changer le journal)
          if (CptNewJal<>CptJournal) then begin
            gbInsert := True;
            szDate := TobEcrDet.GetValue('E_DATECOMPTABLE'); //'01/'+Mois.Value+'/'+OldAnnee.Value;
            DDeb:=DebutdeMois(StrToDate(szDate)); DFin:=FindeMois(StrToDate(szDate));

            S:='E_JOURNAL="'+Journal.Value+'" AND E_EXERCICE="'+QuelExoDT(StrToDate(szDate))+'" '
              +'AND E_DATECOMPTABLE>="'+UsDateTime(DDeb)+'" AND E_DATECOMPTABLE<="'+UsDateTime(DFin)+'" '
              +'AND E_NUMEROPIECE='+IntToStr(TobEcrDet.GetValue('E_NUMEROPIECE'))+' AND E_QUALIFPIECE="'+TobEcrDet.GetValue('E_QUALIFPIECE')+'" ';

            SY:='Y_JOURNAL="'+Journal.Value+'" AND Y_EXERCICE="'+QuelExoDT(StrToDate(szDate))+'" '
               +'AND Y_DATECOMPTABLE>="'+UsDateTime(DDeb)+'" AND Y_DATECOMPTABLE<="'+UsDateTime(DFin)+'" '
               +'AND Y_NUMEROPIECE='+IntToStr(TobEcrDet.GetValue('E_NUMEROPIECE'))+' AND Y_QUALIFPIECE="'+TobEcrDet.GetValue('E_QUALIFPIECE')+'" ';

            // S'il y a de l'analytique
            if (TobEcrDet.Detail.Count >0) then begin
              for i :=0 to TobEcrDet.Detail.Count-1 do begin
                // Mise à jour de la table SECTION
                Debit :=  IntToStr(TobEcrDet.Detail[i].GetValue('Y_DEBIT'));
                Credit := IntToStr(TobEcrDet.Detail[i].GetValue('Y_CREDIT'));
                SSect:='UPDATE SECTION SET S_TOTALDEBIT=S_TOTALDEBIT-'+Debit
                      +', S_TOTALCREDIT=S_TOTALCREDIT-'+Credit
                      +', S_TOTDEBE=S_TOTDEBE-'+Debit
                      +', S_TOTCREE=S_TOTCREE-'+Credit;
                // L'écriture passe sur l'exercice suivant
                if (StrToInt(Exo1)<StrToInt(Exo2)) then begin
                  SSect:=SSect+', S_TOTDEBS=S_TOTDEBS+'+Debit
                              +', S_TOTCRES=S_TOTCRES+'+Credit;
                end;
                // L'écriture passe sur l'exercice précédent
                if (StrToInt(Exo1)>StrToInt(Exo2)) then begin
                  SSect:=SSect+', S_TOTDEBP=S_TOTDEBP+'+Debit
                              +', S_TOTCREP=S_TOTCREP+'+Credit;
                end;
                SSect:=SSect+' WHERE S_AXE="'+TobEcrDet.Detail[i].GetValue('Y_AXE')
                            +'" AND S_SECTION="'+TobEcrDet.Detail[i].GetValue('Y_SECTION')+'"';
                ExecuteSQL(SSect);
              end;
            end;

            // Suppression de la pièce via le compte d'attente
            if (Numeropiece <> TobEcrDet.GetValue('E_NUMEROPIECE')) then begin
              // Supprime l'analytique puis toutes les lignes d'écritures > à 2
              ExecuteSQL('Delete from Analytiq Where '+SY);
              ExecuteSQL('Delete from Ecriture Where '+S+'AND (E_NUMLIGNE>2 OR E_NUMECHE>1)');

              // On ne modifie plus les écritures car elles sont supprimées de la base au début du traitement // 14226
              TobMAJBase := Tob.Create('ECR',NIL,-1) ;
              // FQ 17385 on recopie les tobs sans leur détail sinon on ré-insère l'analytique !
              for i:=0 to TDetruite.Detail.count-1 do
                begin
                lTobTmp := Tob.Create('ECRITURE', TobMAJBase, -1) ;
                lTobTmp.Dupliquer( TDetruite.detail[i] ,False ,True );
                end ;

              for i := TobMAJBase.Detail.Count-1 downto 0 do begin
                iNumLig := TobMAJBase.Detail[i].GetInteger('E_NUMLIGNE');
                if (iNumLig > 2) then TobMAJBase.Detail[i].Free
                else begin
                  if iNumLig=1 then szMtt:='1' else szMtt:='-1' ;
                  TobMAJBase.Detail[i].SetString('E_GENERAL',VH^.Cpta[fbGene].Attente);
                  TobMAJBase.Detail[i].SetString('E_AUXILIAIRE', '');
                  TobMAJBase.Detail[i].SetString('E_DEBIT',szMtt);
                  TobMAJBase.Detail[i].SetDouble('E_CREDIT',0);
                  TobMAJBase.Detail[i].SetString('E_DEBITDEV',szMtt);
                  TobMAJBase.Detail[i].SetDouble('E_CREDITDEV',0);
                  TobMAJBase.Detail[i].SetString('E_VALIDE','X');
                  TobMAJBase.Detail[i].SetString('E_LIBELLE',Copy('DETRUITE LE '+DateToStr(Date)+' PAR '+V_PGI.UserName,1,35) );
                  TobMAJBase.Detail[i].SetDouble('E_COUVERTURE',0);
                  TobMAJBase.Detail[i].SetDouble('E_COUVERTUREDEV',0);
                  TobMAJBase.Detail[i].SetString('E_LETTRAGE','');
                  TobMAJBase.Detail[i].SetString('E_LETTRAGEDEV','-');
                  TobMAJBase.Detail[i].SetInteger('E_TAUXDEV',1);
                  TobMAJBase.Detail[i].SetDouble('E_QTE1',0);
                  TobMAJBase.Detail[i].SetDouble('E_QTE2',0);
                  TobMAJBase.Detail[i].SetString('E_ECHE','-');
                  TobMAJBase.Detail[i].SetString('E_ANA','-');
                  TobMAJBase.Detail[i].SetInteger('E_NUMECHE',0);
                  TobMAJBase.Detail[i].SetString('E_TYPEMVT','DIV');
                  TobMAJBase.Detail[i].SetString('E_DEVISE',V_PGI.DevisePivot);
                  TobMAJBase.Detail[i].SetString('E_TVAENCAISSEMENT','-');
                  TobMAJBase.Detail[i].SetString('E_CONTREPARTIEGEN','');
                  TobMAJBase.Detail[i].SetString('E_CONTREPARTIEAUX','');
                  TobMAJBase.Detail[i].SetString('E_CREERPAR','DET');
                  TobMAJBase.Detail[i].SetDateTime('E_DATEMODIF',NowH);
                end;
              end;

              TobMAJBase.InsertOrUpdateDb;
              TobMAJBase.Free;

              // PLus valables à partir du moment où on supprime les écritures en bases (voir plus haut)
              // Modifie les 2 lignes de l'écriture restantes
//              if ExecuteSQL('UPDATE Ecriture SET '+UpdateSet(1)+' Where '+S+'AND E_NUMLIGNE=1')<>1 then BEGIN V_PGI.IOError:=oeSaisie ; FinTraitement; LibereTobs; Exit ; END ;
//              if ExecuteSQL('UPDATE Ecriture SET '+UpdateSet(2)+' Where '+S+'AND E_NUMLIGNE=2')<>1 then BEGIN V_PGI.IOError:=oeSaisie ; FinTraitement; LibereTobs; Exit ; END ;

              NewNumeropiece := GetNewNumJal(gszNewjal,True,Date);
              OldNumeropiece := NewNumeropiece;
            end
            else NewNumeropiece := OldNumeropiece;
            Numeropiece := TobEcrDet.GetValue('E_NUMEROPIECE');
          end
          else NewNumeropiece:=TobEcrDet.GetValue('E_NUMEROPIECE');
        end
        else NewNumeropiece:=TobEcrDet.GetValue('E_NUMEROPIECE');

        TobEcrdet.PutValue('E_NUMEROPIECE',NewNumeropiece) ;
        TobEcrdet.PutValue('E_IO','X') ;
        If TobEcrDet.GetValue('E_ANA')= 'X' then MajEcrAna (TobEcrDet, 'Y_NUMEROPIECE', NewNumeropiece) ;

        If ChgEtab then ChangeEtablissement (TobEcrDet) ; // Changement d'établissement
        If ChgExo or ChgJal then CumulJournal (TobEcrDet) ;
        If ChgJal then ChangeJournal (TobEcrDet) ; //Changement de journal
        If ChgPeriode then ChangeDateOuPeriode (TobEcrDet, ChgDate, Copy(TobEcrdet.GetValue('E_DATECOMPTABLE'),0,2)) ; // Changement de date ou de periode
        If SoldeCpteTresoARecalculer then RecalculCompteTreso (TobEcrDet, AncienCptTreso, NouvCptTreso, NatJournal);
        If ChgExo Then CalculSiChangeExo (TobEcrDet, NouvCptTreso, Exo2) ;

        {FQ21880 12.02.2008 YMO Changement de clef sur ECRCOMPL}
        if bCutOff then
        begin
          TobComplDet.DeleteDB ;
          CMAJTOBCompl(TobEcrDet);
        end;

      END ; // Fin du For k

      If ChgExo or ChgJal or ChgPeriode or ChgDate then RecalculSolde (ChgExo, True, Exo2, AncienCptTreso, NouvCptTreso); //Recalcul solde des journaux
      If SoldeCpteTresoARecalculer then RecalculSolde (ChgExo, False, Exo2, AncienCptTreso, NouvCptTreso) ; // Recalcul des comptes de tresorerie
    END ; //Fin du If ChgEtab or ChgPeriode then

    If ChgNatPie Then ChangeNaturePiece ; //Changement de Nature de piece

    { FQ 18966 BVE 24.04.07 }
    lMess := TMessageCompta.Create('Modification entête de pièces') ;
    CChargeTInfoEcr(TobEcr,lInfo) ;
    try
       lRecError := CIsValidSaisiePiece(TobEcr,lInfo) ;
       // Vérification des données :
       if (lRecError.RC_Error = RC_PASERREUR) and MajTobs then
          CommitTrans
       else
       begin
          if (lRecError.RC_Error <> RC_PASERREUR) then
             lMess.Execute(lRecError.RC_Error) ;
          Rollback;
       end;
    finally
       lMess.Free ;
       lInfo.Free ;
    end;
    { END FQ 18966 }
  Except
   on e : Exception do
    begin
     Rollback ;
     PGIInfo( 'ATTENTION : Programme interrompu. Relancer le traitement.' + #10#13 + E.Message );
    // AfficheMsg(5,'','') ; // Programme interrompu. Relancer le traitement.
    end ;
  END ;

  if (CptNewJal<>CptJournal) then begin
    if LMsg.Execute(6,'','')=mrYes then VisuPiecesGenere(ListePieces,EcrGen,0) ;
    VideListe(ListePieces) ; ListePieces.Free ;
    if TDetruite <> nil then begin TDetruite.free; end;
  end;
  LibereTobs ;
  FinTraitement ;
END ;

// Récupération des écritures générales et analytiques
procedure TOF_MODIFENTPIE.RecupMultiSel(pbChgperiode : Boolean; pszPeriode : String);
Var
  i :integer ;
  TobEcrDetail : TOB ;
  {$IFDEF EAGLCLIENT}
  Fiche : TFMul ;
  {$ENDIF}
begin
  // On récupère toutes les écritures du mul ainsi que toutes les analytiques correspondantes
  // puis on supprime celles qui ne sont pas sélectionnées.
  // Si on est sur un Journal de type BORdereau, on crée aussi un tob qui permettra de récupérer le reste du bordereau (afin de renuméroter les lignes)
  // Nota: Les écritures analytiques sont des filles de la tob détail
  // Tob Virtuelle
  //  ->Tob fille : Contient une ligne d'écriture comptable
  //   ->Tob fille : Contient l'analytique de l'écriture comptable
  TobEcr := Tob.create('_ECRITURE',NIL,-1);
  PrechargeOrigines;
  RecupEcritAna; // Recupère les écritures analytiques correspondantes

  If (ListeEcr.AllSelected) then
  begin
    // Si on est sur un journal de type BORdereau et qu'on change de période
    if (gbJalBOR and pbChgperiode) then
    begin
      // Supprime les écritures qui se trouveraient sur la période de destination
      for i:=TOBEcr.Detail.Count-1 downto 0 do
      begin
        TobEcrDetail:=TOBEcr.Detail[i] ;
        // Si elle est dans la période de destination : La supprime
        If (TobEcrDetail.GetValue('E_PERIODE')=pszPeriode) then TobEcrDetail.Free;
        // BPY le 25/06/2004 => Fiche 13297 : Mode Revision
        If (RevisionActive(TobEcrDetail.GetValue('E_DATECOMPTABLE'))) then TobEcrDetail.Free;
        // Fin BPY
      end;
    end;
  end
  else
  begin
    for i:=0 to ListeEcr.NbSelected-1 do
    begin
      ListeEcr.GotoLeBookmark(i);
      {$IFDEF EAGLCLIENT}
      Fiche := TFMul(Ecran) ;
      Fiche.Q.TQ.Seek(Fiche.FListe.Row-1);
      {$ENDIF}
      // BPY le 25/06/2004 => Fiche 13297 : Mode Revision
      if (not (RevisionActive(QlisteEcr.FindField('E_DATECOMPTABLE').AsDateTime))) then MarqueOrigine; // Marque les lignes sélectionnées
      // Fin BPY
    end;
    VireInutiles(pbChgperiode,pszPeriode) ; // Enlève les lignes non sélectionnées
  end;
end;

// Chargement dans la TOBECR des écritures sélectionnées
procedure TOF_MODIFENTPIE.PrechargeOrigines ;
Var Q : TQuery ;
    St, St1, St2 : String ;
BEGIN
  TobEcr.ClearDetail ;
  St := 'SELECT ECRITURE.* FROM ECRITURE LEFT OUTER JOIN EXERCICE ON E_EXERCICE=EX_EXERCICE '; // Ne pas supprimer la liaison avec la table EXERCICE // 13854
  St2:= GetControlText('XX_WHERE');
  SetControlText('XX_WHERE','');
  Pages.Pages[2].TabVisible := False; // Evite de récupèrer les critères de l'onglet Avancé
  St1:=AnsiUpperCase(RecupWhereCritere(Pages)) ;
  Pages.Pages[2].TabVisible := True;
  SetControlText('XX_WHERE',St2);
  St := St + St1 + ' ORDER BY E_JOURNAL, E_PERIODE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMGROUPEECR';
  Q := OpenSQL(St, True) ;
  TobEcr.LoadDetailDB('ECRITURE','','',Q,False,False) ;
  Ferme(Q) ;
  If TOBEcr.Detail.Count>0 then TOBEcr.Detail[0].AddChampSup('MARQUE',True) ;
END ;

// Marquage des écritures générales sélectionnées dans la liste de la mul
Procedure TOF_MODIFENTPIE.MarqueOrigine ;
Var
  TobEcrDetail : TOB ;
BEGIN
  if QListeEcr = nil then Exit;
  TobEcrDetail := TobEcr.FindFirst(['E_JOURNAL','E_EXERCICE','E_PERIODE','E_NUMEROPIECE','E_NUMGROUPEECR','E_QUALIFPIECE'],
                                   [QListeEcr.FindField('E_JOURNAL').AsString,
                                    QListeEcr.FindField('E_EXERCICE').AsString,
                                    QListeEcr.FindField('E_PERIODE').AsInteger,
                                    QListeEcr.FindField('E_NUMEROPIECE').AsInteger,
                                    QListeEcr.FindField('E_NUMGROUPEECR').AsInteger,
                                    QListeEcr.FindField('E_QUALIFPIECE').AsString],False) ;
  while TobEcrDetail <> Nil do BEGIN
    TobEcrDetail.PutValue('MARQUE','X') ;
    TobEcrDetail := TobEcr.FindNext(['E_JOURNAL','E_EXERCICE','E_PERIODE','E_NUMEROPIECE','E_NUMGROUPEECR','E_QUALIFPIECE'],
                                     [QListeEcr.FindField('E_JOURNAL').AsString,
                                      QListeEcr.FindField('E_EXERCICE').AsString,
                                      QListeEcr.FindField('E_PERIODE').AsInteger,
                                      QListeEcr.FindField('E_NUMEROPIECE').AsInteger,
                                      QListeEcr.FindField('E_NUMGROUPEECR').AsInteger,
                                      QListeEcr.FindField('E_QUALIFPIECE').AsString],False) ;
  END ;
END ;

// Destruction des écritures non sélectionnées
procedure TOF_MODIFENTPIE.VireInutiles(pbChgperiode : Boolean; pszPeriode : String);
Var i : integer ;
    TobEcrDetail  : TOB ;
BEGIN
  for i:=TOBEcr.Detail.Count-1 downto 0 do BEGIN
    TobEcrDetail:=TOBEcr.Detail[i] ;
    // Si l'écriture n'est pas sélectionnée
    if TobEcrDetail.GetValue('MARQUE')<>'X' then begin
//      // Si on est sur un journal de type BORdereau et qu'on change de période et qu'elle est dans la période de destination
//      if (gbJalBOR and pbChgperiode) and (TobEcrDetail.GetValue('E_PERIODE')=pszPeriode) then begin
//        // Change le parent
//        TobEcrDetail.ChangeParent(TobEcrBOR,0);
//      end
//      // Sinon
//      else begin
        // La supprime
        TobEcrDetail.Free;
//      end;
    end
    // Si l'écriture est sélectionnée
    else begin
      // Si on est sur un journal de type BORdereau et qu'on change de période
      if (gbJalBOR and pbChgperiode) then begin
        // Si elle est dans la période de destination
        If (TobEcrDetail.GetValue('E_PERIODE')=pszPeriode) then begin
          // La supprime
          TobEcrDetail.Free;
//          // La désélectionne
//          TobEcrDetail.PutValue('MARQUE','-');
//          // Change le parent
//          TobEcrDetail.ChangeParent(TobEcrBOR,0);
        end;
      end;
    end;
  end;
end;

// Récupération des lignes analytiques et rattachement de ces lignes à la TOB principale TOBECR
procedure TOF_MODIFENTPIE.RecupEcritAna  ;
Var  St1, St2, RSql : String ;
     i,y : Integer ;
     Q : TQuery ;
     TobAna : Tob ;
BEGIN
  //Modification du RecupWhereCriteres pour rechercher les écritures analytiques
  St2:= GetControlText('XX_WHERE');
  SetControlText('XX_WHERE','');
  SetControlText('XX_WHERE1','EX_ETATCPTA="OUV" AND E_QUALIFPIECE="N" AND E_CREERPAR<>"DET" AND E_DATECOMPTABLE>"'+UsDateTime(VH^.DateCloturePer)+'"');
  Pages.Pages[2].TabVisible := False;  // Evite de récupèrer les critères de l'onglet Avancé
  St1:=AnsiUpperCase(RecupWhereCritere(Pages)) ;
  Pages.Pages[2].TabVisible := True;
  SetControlText('XX_WHERE',St2);
  {JP 29/07/05 : FQ 15124}
  SetControlText('XX_WHERE1', CST_WHERE1 + 'AND E_DATECOMPTABLE>"'+UsDateTime(VH^.DateCloturePer)+'"');
  St2:='E_' ;
  i:=Pos(St2,St1) ;
  While i>0 do BEGIN
    System.Delete(St1,i,Length(St2)) ;
    System.Insert('Y_',St1,i) ;
    i:=Pos(St2,St1) ;
  END ;
  RSql := 'SELECT ANALYTIQ.* FROM ANALYTIQ LEFT OUTER JOIN EXERCICE ON Y_EXERCICE=EX_EXERCICE ';
  RSql := RSql + St1 + ' ORDER BY Y_JOURNAL, Y_PERIODE, Y_NUMEROPIECE, Y_NUMLIGNE' ;
  Q:=OpenSql(Rsql,True) ;
  If not Q.Eof then BEGIN
    TobAna := TOB.Create('_ANALYTIQ', Nil, -1) ;
    TobAna.LoadDetailDB('ANALYTIQ', '', '', Q, True, True) ;
    //Changement de Parent pour la TOB analytique TobAna
    For i := 0 to TobEcr.Detail.Count -1 do BEGIN
      If (TobEcr.Detail[i].getValue('E_ANA')='X') then BEGIN
        For y:=TobAna.Detail.Count-1 downto 0 do begin
          if (TobEcr.Detail[i].GetValue('E_NUMEROPIECE')=TobAna.Detail[y].GetValue('Y_NUMEROPIECE'))
            And (TobEcr.Detail[i].GetValue('E_NUMLIGNE')=TobAna.Detail[y].GetValue('Y_NUMLIGNE'))
            And (TobEcr.Detail[i].GetValue('E_JOURNAL')=TobAna.Detail[y].GetValue('Y_JOURNAL'))
            And (TobEcr.Detail[i].GetValue('E_QUALIFPIECE')=TobAna.Detail[y].GetValue('Y_QUALIFPIECE'))
            And (TobEcr.Detail[i].GetValue('E_GENERAL')=TobAna.Detail[y].GetValue('Y_GENERAL')) then begin
            TobAna.Detail[y].ChangeParent(TobEcr.Detail[i], -1) ;
          END ;
        END ;
      END ;
    END ;
    TobAna.free ;
  END ;
  if Q <> Nil Then Ferme(Q) ;
END ;

// Mise à jour des lignes analytiques correspondant à l'ecriture générale
Procedure TOF_MODIFENTPIE.MajEcrAna (TobEcrDetail : TOB ; ZoneMaj : String ;  NouvelValeur : Variant) ;
Var i : integer ;
BEGIN
  for i:= 0 to TobEcrDetail.Detail.Count-1 do BEGIN
    TobEcrDetail.Detail[i].PutValue(ZoneMaJ, NouvelValeur) ;
  END ;
END ;

// Cumul débit et crédit pour mise à jour ultérieure des comptes de trésorerie (Banque ou caisse)
Procedure TOF_MODIFENTPIE.RecalculCompteTreso (Var TobEcrDet : TOB ; AncienCptTreso, NouvCptTreso, NatJournal : String) ;
BEGIN
  If (TobEcrDet.GetValue('E_GENERAL')=AncienCptTreso) or ((TobEcrDet.GetValue('E_GENERAL')=NouvCptTreso) and (Natjournal='OD')) then BEGIN
    TDebit := Arrondi(TDebit+TobEcrDet.GetValue('E_DEBIT'),Dec) ;
    TCredit := Arrondi(TCredit+TobEcrDet.GetValue('E_CREDIT'),Dec) ;
  end ;
  If TobEcrDet.GetValue('E_GENERAL')=AncienCptTreso then TobEcrDet.PutValue('E_GENERAL', NouvCptTreso) ;
END ;

// Mise à jour des cumuls des Tobs généraux et si besoin des tobs auxiliaires et sections
// s'il y a changement d'exercice des écritures sélectionnées
Procedure TOF_MODIFENTPIE.CalculSiChangeExo(TobEcrDet : TOB ; NouvCptTreso, Exo2 : String) ;
Var TobGenDet, TobAuxDet, TobSecAnaDet, TobAnaDet : TOB ;
    i : integer ;
BEGIN
  // Changement d'Exercice En Cours à exercice suivant
  If exo2 = VH^.Suivant.Code then BEGIN
    TobEcrDet.PutValue('E_EXERCICE',Exo2) ;
    TobGenDet := TobGen.FindFirst(['G_GENERAL'],[TobEcrDet.GetValue('E_GENERAL')],False) ;
    If TobEcrDet.GetValue('E_GENERAL') <> NouvCptTreso then MajCumulTob(TobGenDet, TobEcrDet , 1, 'G', False) ;
    If TobEcrDet.GetValue('E_AUXILIAIRE')<> '' then BEGIN
      TobAuxDet := TobAux.FindFirst(['T_AUXILIAIRE'],[TobEcrDet.GetValue('E_AUXILIAIRE')],False) ;
      MajCumulTob(TobAuxDet, TobEcrDet , 1, 'T', False ) ;
    END ;
    If TobEcrDet.GetValue('E_ANA') = 'X' then BEGIN
      for i := 0 to TobEcrDet.Detail.Count-1 do BEGIN
        TobAnadet := TobEcrDet.detail[i] ;
        TobSecAnaDet := TobSecAna.FindFirst(['S_SECTION'],[TobAnaDet.GetValue('Y_SECTION')],False) ;
        TobAnaDet.PutValue('Y_EXERCICE',Exo2) ;
        MajCumulTob(TobSecAnaDet, TobAnaDet , 1, 'S', True ) ;
      end ;
    END ;
    END
  // Changement d'Exercice Suivant à Exercice En Cours
  Else BEGIN
    TobEcrDet.PutValue('E_EXERCICE',VH^.EnCours.Code) ;
    TobGenDet := TobGen.FindFirst(['G_GENERAL'],[TobEcrDet.GetValue('E_GENERAL')],False) ;
    If TobEcrDet.GetValue('E_GENERAL') <> NouvCptTreso then MajCumulTob(TobGenDet, TobEcrDet , -1, 'G', False ) ;
    If TobEcrDet.GetValue('E_AUXILIAIRE')<> '' then BEGIN
      TobAuxDet := TobAux.FindFirst(['T_AUXILIAIRE'],[TobEcrDet.GetValue('E_AUXILIAIRE')],False) ;
      MajCumulTob(TobAuxDet, TobEcrDet , -1, 'T', False ) ;
    END ;
    If TobEcrDet.GetValue('E_ANA') = 'X' then BEGIN
      for i := 0 to TobEcrDet.Detail.Count-1 do BEGIN
        TobAnadet := TobEcrDet.detail[i] ;
        TobSecAnaDet := TobSecAna.FindFirst(['S_SECTION'],[TobAnaDet.GetValue('Y_SECTION')],False) ;
        TobAnaDet.PutValue('Y_EXERCICE',VH^.EnCours.Code) ;
        MajCumulTob(TobSecAnaDet, TobAnaDet , -1, 'S', True ) ;
      end ;
    END ;
  END ;
END ;

// Création des TOB Généraux, Auxiliaires, Sections avec les éléments correspondants des écritures sélectionnées
Procedure TOF_MODIFENTPIE.CreationTobs ;
Var i,j : integer ;
    st,ax,sec : string ;
    Q : TQuery ;
    TobDet : TOB ;
BEGIN
  TobAux := Tob.Create ('_AUXILIAIRE', Nil, -1 ) ;
  TobGen := Tob.Create ('_GENERAUX', Nil, -1 ) ;
  TobSecAna := Tob.Create ('_SECTION', Nil, -1 ) ;

  // Création de la TOB Auxiliaire
  for i:= 0 to TobEcr.detail.count-1 do BEGIN
    st:=TobEcr.Detail[i].getValue('E_AUXILIAIRE') ;
    if st <>'' then BEGIN
      If  Tobaux.FindFirst (['T_AUXILIAIRE'],[st],false) = Nil  then BEGIN
        Q := OpenSql('SELECT T_AUXILIAIRE,T_TOTCREE,T_TOTDEBE,T_TOTCRES,T_TOTDEBS FROM TIERS WHERE T_AUXILIAIRE="'+st+'"',true);
        TobAux.LoadDetailDB('AUXILIAIRE','','',Q,true);
        Ferme(Q) ;
      END ;
    END ;

  // Création de la TOB Generaux
  st:=TobEcr.detail[i].getValue('E_GENERAL') ;
  If st <>'' then BEGIN
    If  TobGen.FindFirst (['G_GENERAL'],[st],False) = Nil then BEGIN
      Q := OpenSql('SELECT G_GENERAL, G_FERME, G_TOTALCREDIT, G_TOTALDEBIT, G_TOTCREE, G_TOTCRES, G_TOTDEBE, G_TOTDEBS, G_NATUREGENE FROM GENERAUX WHERE G_GENERAL="'+st+'"', True)  ;
      TobGen.LoadDetailDB('GENERAUX','','',Q,true);
      Ferme(Q) ;
    END ;
  END ;

  // Création de la TOB des sections analytiques
  TobDet:=TobEcr.detail[i] ;
  If TobDet.GetValue('E_ANA') = 'X' then
    for j:=0 to TobDet.detail.Count-1 do BEGIN
      Ax := TobDet.Detail[j].GetValue('Y_AXE') ; Sec := TobDet.Detail[j].GetValue('Y_SECTION') ;
      If  TobSecAna.FindFirst (['S_AXE','S_SECTION'],[Ax,Sec],False) = nil  then BEGIN
        Q := OpenSql('SELECT S_AXE, S_SECTION, S_TOTCREE, S_TOTCRES, S_TOTDEBE, S_TOTDEBS FROM SECTION WHERE S_SECTION ="'+Sec+'"', True) ;
        TobSecAna.LoadDetailDB('SECTION','','',Q,true);
        Ferme(Q) ;
      END ;
    END ;
  END;
END ;

// Création de la tob journal avec les 2 journaux (l'ancien et le nouveau)
Procedure TOF_MODIFENTPIE.CreationTobJournal(bAvecNew : Boolean) ;
Var RSql : String ;
    Q : TQuery ;
BEGIN
  RSql := 'SELECT J_JOURNAL, J_NATUREJAL, J_FERME, J_COMPTEURNORMAL, J_MULTIDEVISE, J_CONTREPARTIE, ' ;
  RSql := RSql + 'J_TOTALDEBIT, J_TOTALCREDIT, J_TOTDEBE, J_TOTCREE, J_TOTDEBS, J_TOTCRES, J_MODESAISIE' ;
  RSql := RSql + ' FROM JOURNAL WHERE ';
  if bAvecNew then Rsql := RSql + 'J_JOURNAL="'+gszNewJal+'" OR ';
  Rsql := RSql + 'J_JOURNAL="'+Journal.Value+'"';
  Q := OpenSql(Rsql,True) ;
  TobJal := Tob.Create ('_JOURNAL', Nil, -1 ) ;
  TobJal.LoadDetailDB('JOURNAL', '', '', Q, True, True) ;
  Ferme(Q);
END ;

procedure TOF_MODIFENTPIE.Arret ;
BEGIN
  FinTraitement ;
  LibereTobs ;
  PGIInfo(TraduireMemoire('Aucun changement à faire et / ou aucune écriture sélectionnée.'),TraduireMemoire(Titre));
  SetActiveTabSheet('PCritere') ;
  Journal.SetFocus;
END;

// Mise à jour de l'établissement
Procedure TOF_MODIFENTPIE.ChangeEtablissement (Var TobEcrDet : TOB ) ;
BEGIN
  TobEcrdet.PutValue('E_ETABLISSEMENT',gszNewEtab) ;
  if TobEcrDet.GetValue('E_ANA')='X' then MajEcrAna(TobEcrDet, 'Y_ETABLISSEMENT', gszNewEtab) ;
END ;

// Mise à jour du code journal de l'écriture par le nouveau code
Procedure TOF_MODIFENTPIE.ChangeJournal(Var TobEcrDetail : TOB) ;
BEGIN
  TobEcrdetail.PutValue('E_JOURNAL',gszNewJal) ;
  If TobEcrDetail.getValue('E_ANA')='X' then MajEcrAna(TobEcrDetail,'Y_JOURNAL',gszNewJal) ;
END ;

// Mise à jour de la nature de la pièce comptable
Procedure TOF_MODIFENTPIE.ChangeNaturePiece ;
Var i : Integer ;
    TobEcrDet : TOB ;
BEGIN
  For i:=0 TO TobEcr.Detail.Count-1 do BEGIN
    TobEcrDet := TobEcr.Detail[i];
    TobEcrDet.PutValue('E_NATUREPIECE',gszNewNatPie);
    if TobEcrDet.GetValue('E_ANA')= 'X' then MajEcrAna (TobEcrDet, 'Y_NATUREPIECE', gszNewNatPie);
  END ;
END ;

// Mise à jour des zones de date si changement de date (période, exercice)
Procedure TOF_MODIFENTPIE.ChangeDateOuPeriode (Var TobEcrDetail : TOB ; ChgDate : Boolean; szDay : String);
Var
  NouDate : TDateTime ;
  iDay,iMois,iAnnee : Integer;
BEGIN
  iDay:=0; iMois:=0; iAnnee:=0;
  if (gbRBMOIS) then begin
    iDay := StrToInt(szDay);
    iMois:= StrToInt(gszNewMois);
    iAnnee:=StrToInt(gszNewAnnee);
    // Pour passer au dernier jour du mois
    if ((iDay>28) and (iMois = 2)) then begin
      if (iAnnee mod 4 = 0) then iDay:=29 else iDay:=28;
    end;
    if ((iDay=31) and ((iMois=4) or (iMois=6) or (iMois=9) or (iMois=11))) then iDay:=30;
  end;
  If ChgDate then NouDate := StrToDate(gszNewdate)
             else NouDate := EncodeDate(iAnnee,iMois,iDay);
  TobEcrdetail.PutValue('E_DATECOMPTABLE',NouDate) ;
  TobEcrDetail.PutValue('E_PERIODE',GetPeriode(NouDate)) ;
  TobEcrDetail.PutValue('E_SEMAINE',NumSemaine(NouDate)) ;
  If TobEcrDetail.GetValue('E_ANA')='X' then BEGIN
    MajEcrAna(TobEcrDetail,'Y_DATECOMPTABLE',NouDate) ;
    MajEcrAna(TobEcrDetail,'Y_PERIODE',GetPeriode(NouDate)) ;
    MajEcrAna(TobEcrDetail,'Y_SEMAINE',NumSemaine(NouDate)) ;
  END ;
END ;

// Mise à jour des cumuls
procedure TOF_MODIFENTPIE.MajCumulTob (Var TobDetail : Tob ; TobEcrDet : Tob ; Signe : Integer ; Lettre : String ; Ana : Boolean ) ;
Var L : string ;
BEGIN
  if ana then L:='Y' Else L:='E' ;
  TobDetail.PutValue(Lettre +'_TOTDEBE',arrondi(TobDetail.GetValue(Lettre +'_TOTDEBE')+(-1*Signe*TobEcrDet.GetValue(L+'_DEBIT')),Dec)) ;
  TobDetail.PutValue(Lettre +'_TOTCREE',arrondi(TobDetail.GetValue(Lettre +'_TOTCREE')+(-1*Signe*TobEcrDet.GetValue(L+'_CREDIT')),Dec)) ;
  TobDetail.PutValue(Lettre +'_TOTDEBS',arrondi(TobDetail.GetValue(Lettre +'_TOTDEBS')+(Signe*TobEcrDet.GetValue(L+'_DEBIT')),Dec)) ;
  TobDetail.PutValue(Lettre +'_TOTCRES',arrondi(TobDetail.GetValue(Lettre +'_TOTCRES')+(Signe*TobEcrDet.GetValue(L+'_CREDIT')),Dec)) ;
END ;

// Libération des TOBS
Procedure TOF_MODIFENTPIE.LibereTobs ;
BEGIN
  If TobEcr <>      NIL then begin TobEcr.Free ;      TobEcr:=Nil ;      end ;
  If TobJal <>      NIL then begin TobJal.Free ;      TobJal:=Nil ;      end ;
  If TobGen <>      NIL Then begin TobGen.Free ;      TobGen:=Nil ;      end ;
  If TobExercice <> NIL then Begin TobExercice.Free ; TobExercice:=Nil ; end ;
  If TobAux <>      NIL Then begin Tobaux.free ;      TobAux:=Nil ;      end ;
  If TobSecAna <>   NIL then begin TobSecAna.free ;   TobSecAna:=Nil ;   end ;
  If TobCompl  <>   NIL then begin TobCompl.free ;    TobCompl:=Nil ;    end ;
//  If TobEcrBOR <>   NIL then begin TobEcrBOR.Free ;   TobEcrBOR:=Nil ;   end ;
END ;

// Mise à jour des TOBS
function TOF_MODIFENTPIE.MajTobs : boolean;
BEGIN
  Result := False;
  // Pour forcer la mise à jour des tob
  If TobEcr <> NIL then TobEcr.SetAllModifie(True); // 14771

    {$IFDEF TRSYNCHRO}
    {JP 17/11/04 : FQ 14972 : Mise à jour du champ E_TRESOSYNCHRO}
    if EstComptaTreso then
      MajE_TRESOSYNCHROTob(TobEcr, TobJal, TobGen);
    {$ENDIF}

  if gbInsert or gbJalBOR then begin
//    If TobEcrBOR <> NIL then TobEcrBOR.SetAllModifie(True);
    If TobEcr <> NIL then Result := TobEcr.InsertOrUpdateDb ;
    If TobJal <> NIL then Result := Result and TobJal.InsertOrUpdateDb ;
    If TobGen <> NIL Then  Result := Result and TobGen.InsertOrUpdateDb ;
    If TobAux <> NIL Then Result := Result and Tobaux.InsertOrUpdateDb ;
    If TobSecAna <> NIL then Result := Result and TobSecAna.InsertOrUpdateDb ;
    If TobCompl <> NIL then Result := Result and TobCompl.InsertOrUpdateDb ;
//    If TobEcrBOR <> NIL then Result := Result and TobEcrBOR.InsertOrUpdateDb ;
    end
  else begin
    // 13854
    If TobEcr <> NIL then Result := TobEcr.InsertOrUpdateDb ;
    If TobJal <> NIL then Result := Result and TobJal.InsertOrUpdateDb ;
    If TobGen <> NIL Then  Result := Result and TobGen.InsertOrUpdateDb ;
    If TobAux <> NIL Then Result := Result and Tobaux.InsertOrUpdateDb ;
    If TobSecAna <> NIL then Result := Result and TobSecAna.InsertOrUpdateDb ;
    If TobCompl <> NIL then Result := Result and TobCompl.InsertOrUpdateDb ; {FQ21880 05/12/2007  YMO Maj EcrCompl}
//    If TobEcrBOR <> NIL then Result := Result and TobEcrBOR.InsertOrUpdateDb ;
  end;
end;

Procedure TOF_MODIFENTPIE.RecalculSolde (ChgExo, Journaux  : Boolean ; Exo2, AncienCptTreso, NouvCptTreso : String );
Var DebSuiv, CreSuiv, DebEC, CreEC, Rech, Anc, Nou, L : String ;
Deb, Cre : Double ;
TOB1, TobDet : TOB ;
BEGIN
  if Journaux then BEGIN DebSuiv:='J_TOTDEBS'; CreSuiv:= 'J_TOTCRES'; DebEC:='J_TOTDEBE'; CreEC:='J_TOTCREE'; Deb:=JDebit; Cre:=JCredit; TOB1:=TobJal; Rech:='J_JOURNAL'; Anc:=Journal.Value;  Nou:=gszNewJal   ; L:='J'; End
              else BEGIN DebSuiv:='G_TOTDEBS'; CreSuiv:= 'G_TOTCRES'; DebEC:='G_TOTDEBE'; CreEC:='G_TOTCREE'; Deb:=TDebit; Cre:=TCredit; TOB1:=TobGen; Rech:='G_GENERAL'; Anc:=AncienCptTreso; Nou:=NouvCptTreso; L:='G'; End ;
  if Exo2=VH^.Suivant.Code then BEGIN
    if not ChgExo Then BEGIN DebEC:=DebSuiv ; CreEC:=CreSuiv ; end ;
    TobDet := TOB1.FindFirst([Rech],[Nou],False) ;
    if TobDet <> Nil then BEGIN
      TobDet.PutValue(DebSuiv, Arrondi(TobDet.GetValue(DebSuiv)+Deb,Dec)) ;
      TobDet.PutValue(CreSuiv, Arrondi(TobDet.GetValue(CreSuiv)+Cre,Dec)) ;
      TobDet.PutValue(L+'_TOTALDEBIT', Arrondi(TobDet.GetValue(L+'_TOTALDEBIT')+Deb,Dec)) ;
      TobDet.PutValue(L+'_TOTALCREDIT', Arrondi(TobDet.GetValue(L+'_TOTALCREDIT')+Cre,Dec)) ;
    END ;
    TobDet := TOB1.FindFirst([Rech],[Anc],False) ;
    if TobDet <> NIL THEN BEGIN
      TobDet.PutValue(DebEC, Arrondi(TobDet.GetValue(DebEC)-Deb,Dec)) ;
      TobDet.PutValue(CreEC, Arrondi(TobDet.GetValue(CreEC)-Cre,Dec)) ;
      TobDet.PutValue(L+'_TOTALDEBIT', Arrondi(TobDet.GetValue(L+'_TOTALDEBIT')-Deb,Dec)) ;
      TobDet.PutValue(L+'_TOTALCREDIT', Arrondi(TobDet.GetValue(L+'_TOTALCREDIT')-Cre,Dec)) ;
    END ;
    END
  Else BEGIN
    if not ChgExo Then BEGIN DebSuiv:=DebEC ; CreSuiv:=CreEC ; end ;
    TobDet := TOB1.FindFirst([Rech],[Nou],False) ;
    If TobDet <> NIL THEN BEGIN
      TobDet.PutValue(DebEC, Arrondi(TobDet.GetValue(DebEC)+Deb,Dec)) ;
      TobDet.PutValue(CreEC, Arrondi(TobDet.GetValue(CreEC)+Cre,Dec)) ;
      TobDet.PutValue(L+'_TOTALDEBIT', Arrondi(TobDet.GetValue(L+'_TOTALDEBIT')+Deb,Dec)) ;
      TobDet.PutValue(L+'_TOTALCREDIT', Arrondi(TobDet.GetValue(L+'_TOTALCREDIT')+Cre,Dec)) ;
    END ;
    TobDet := TOB1.FindFirst([Rech],[Anc],False) ;
    If TobDet <> NIl then BEGIN
      TobDet.PutValue(DebSuiv, Arrondi(TobDet.GetValue(DebSuiv)-Deb,Dec)) ;
      TobDet.PutValue(CreSuiv, Arrondi(TobDet.GetValue(CreSuiv)-Cre,Dec)) ;
      TobDet.PutValue(L+'_TOTALDEBIT', Arrondi(TobDet.GetValue(L+'_TOTALDEBIT')-Deb,Dec)) ;
      TobDet.PutValue(L+'_TOTALCREDIT', Arrondi(TobDet.GetValue(L+'_TOTALCREDIT')-Cre,Dec)) ;
    END ;
  END ;
END ;

Procedure TOF_MODIFENTPIE.FinTraitement ;
BEGIN
  TButton(GetControl('BCHERCHE',True)).Click ;
  EnableControls(Ecran, True ) ;
END ;

// Cumul Debit et credit pour mise à jour ultérieure des soldes
Procedure TOF_MODIFENTPIE.CumulJournal (TobEcrDetail : TOB);
BEGIN
  JDebit := Arrondi(JDebit+TobEcrDetail.GetValue('E_DEBIT'),Dec) ;
  JCredit := Arrondi(JCredit+TobEcrDetail.GetValue('E_CREDIT'),Dec) ;
END ;

// Création d'une tob de comptes generaux (2 au maximum) représentant l'ancien et le nouveau compte de
// trésorerie) ; Uniquement si les journaux sélectionnés sont de type "trésorerie"
Procedure TOF_MODIFENTPIE.AjoutCptGenerauxTreso (AncienCpttreso, NouvCptTreso : String) ;
Var Q : TQuery ;
BEGIN
  if TobGen = NIL then TobGen := Tob.Create('_GENERAUX', Nil, -1) ;
  if TobGen.Findfirst(['G_GENERAL'],[AncienCptTreso],False) = nil then begin
    Q := OpenSql('SELECT G_GENERAL, G_FERME, G_TOTALCREDIT, G_TOTALDEBIT, G_TOTCREE, G_TOTCRES, G_TOTDEBE, G_TOTDEBS, G_NATUREGENE FROM GENERAUX WHERE G_GENERAL="'+AncienCptTreso+'"', True)  ;
    TobGen.LoadDetailDB('GENERAUX','','',Q,true); Ferme(Q) ;
  end ;
  if TobGen.Findfirst(['G_GENERAL'],[NouvCptTreso],False) = nil then begin                                        {JP 17/12/04 : FQ 15122}
    Q := OpenSql('SELECT G_GENERAL, G_FERME, G_TOTALCREDIT, G_TOTALDEBIT, G_TOTCREE, G_TOTCRES, G_TOTDEBE, G_TOTDEBS, G_NATUREGENE FROM GENERAUX WHERE G_GENERAL="'+NouvCptTreso+'"', True)  ;
    TobGen.LoadDetailDB('GENERAUX','','',Q,true); Ferme(Q) ;                                            
  end ;
END ;

Procedure TOF_MODIFENTPIE.FListeDblClick (Sender: TObject);
Var sMode : String ;
    Fiche : TFMul ;
begin
  inherited;
  Fiche := TFMul(Ecran) ;
  ListeEcr := Fiche.FListe ;
{$IFDEF EAGLCLIENT}
  QListeEcr := Fiche.Q.TQ;
  Fiche.Q.TQ.Seek(Fiche.FListe.Row-1);
{$ELSE}
  QListeEcr := Fiche.Q;
{$ENDIF}
  if ((QListeEcr.EOF) and (QListeEcr.BOF)) then Exit ;
  if QListeEcr.FindField('E_MODESAISIE')=Nil then begin
    PGiBox(TraduireMemoire('Le Champ E_MODESAISIE n''est pas dans la liste.'),TraduireMemoire('Champ E_MODESAISIE INEXISTANT')) ;
    Exit ;
  end ;
  sMode:=QListeEcr.FindField('E_MODESAISIE').AsString ;
  if ((sMode<>'') and (sMode<>'-')) then LanceSaisieFolio(QListeEcr, TaConsult)
                                    else TrouveEtLanceSaisie(QListeEcr, TaConsult, 'N') ;
end;

function TOF_MODIFENTPIE.AfficheMsg(num : integer;Av,Ap : string ) : Word ;
begin
  Result:=mrNone ;
  if LMsg=nil then exit ;
  if (Num<0) or (Num>LMsg.Mess.Count-1) then Num:=0 ; //erreur inconnue
  Result:=LMsg.Execute(num,Av,Ap) ;
end;

// Pour la destruction des écritures
// (Journal avec un compteur différent)
{function TOF_MODIFENTPIE.UpdateSet(NumL: integer): String;
Var StM,StL : String ;
BEGIN
StL:=Copy('DETRUITE LE '+DateToStr(Date)+' PAR '+V_PGI.UserName,1,35) ;
if NumL=1 then StM:='1' else StM:='-1' ;
Result:=' E_GENERAL="'+VH^.Cpta[fbGene].Attente+'", E_AUXILIAIRE="", E_DEBIT='+StM+','
       +' E_CREDIT=0, E_DEBITDEV='+StM+', E_CREDITDEV=0, E_VALIDE="X", E_LIBELLE="'+StL+'",'
       +' E_COUVERTURE=0, E_COUVERTUREDEV=0, E_LETTRAGE="", E_LETTRAGEDEV="-", E_TAUXDEV=1,'
       +' E_QTE1=0, E_QTE2=0, E_ECHE="-", E_ANA="-", E_NUMECHE=0, E_TYPEMVT="DIV",'
       +' E_DEVISE="'+V_PGI.DevisePivot+'", E_TVAENCAISSEMENT="-", '
       +' E_CONTREPARTIEGEN="", E_CONTREPARTIEAUX="", E_CREERPAR="DET", '
       +' E_DATEMODIF="'+UsTime(NowH)+'"' ;
end;}

procedure TOF_MODIFENTPIE.FermeMsg ;
begin LMsg.Free ; end;

procedure TOF_MODIFENTPIE.InitMsg ;
begin
  LMsg:=THMsgBox.create(FMenuG) ;
  {00}LMsg.Mess.Add(traduirememoire('0;Modification entête de pièces;Erreur Inconnue.;W;O;O;O'));
  {01}LMsg.Mess.Add(traduirememoire('1;Modification entête de pièces;ARRET DU TRAITEMENT : Les dates des exercices ne correspondent pas.;W;O;O;O'));
  {02}LMsg.Mess.Add(traduirememoire('2;Modification entête de pièces;ARRET DU TRAITEMENT : Au moins un des 2 journaux est fermé.;W;O;O;O'));
  {03}LMsg.Mess.Add(traduirememoire('3;Modification entête de pièces;ARRET DU TRAITEMENT : Les journaux n''ont pas le même état multidevises.;W;O;O;O'));
  {04}LMsg.Mess.Add(traduirememoire('4;Modification entête de pièces;ARRET DU TRAITEMENT : Les journaux n''ont pas le même mode de saisie.;W;O;O;O'));
  {05}LMsg.Mess.Add(traduirememoire('5;Modification entête de pièces;ATTENTION : Programme interrompu. Relancer le traitement.;W;O;O;O'));
  {06}LMsg.Mess.Add(traduirememoire('6;Modification entête de pièces;Voulez-vous voir la liste des écritures détruites ?;Q;YN;N;N'));
end;

// Modifie les DATES lors du changement d'EXERCICE
procedure TOF_MODIFENTPIE.ExerciceOnChange(Sender: TObject);
begin
  ExoToDates(Exercice.Value, DateComptaDe, DateComptaA);
end;

function TOF_MODIFENTPIE.RecupWhereCritere(PC: TPageControl): String;
var
  i, j : Integer;
  C : TControl;
  sz, szName, szValue : String;
  Ch : Char;
begin
  for i := 0 to PC.PageCount-1 do begin
    // Si invisible : passe au suivant
    if not PC.Pages[i].TabVisible then Continue;
    for j := 0 to PC.Pages[i].ControlCount-1 do begin
      C := PC.Pages[i].Controls[j];


      if (Pos('_', C.Name) > 0) then begin
        if Copy(C.Name, Length(C.Name), 1) = '_' then szName := UpperCase(Copy(C.Name, 1, Length(C.Name)-1))
                                                 else szName := UpperCase(C.Name);

        if (C is THValComboBox) then begin
          szValue := THValComboBox(C).Value;
          if (szValue = '') then Continue;
          case THValComboBox(C).Operateur of
             Commence : sz := szName+' LIKE '+szValue+'%';
             Egal     : sz := szName+'="'+szValue+'"';
             Contient : sz := szName+' LIKE %'+szValue+'%';
             Superieur: sz := szName+' >="'+szValue+'"';
             Inferieur: sz := szName+' <="'+szValue+'"';
             Different: sz := szName+' <>"'+szValue+'"';
           end;
          if (Result = '') then Result := 'WHERE '+sz
                           else Result := Result+' AND '+sz;
          Continue;
          end
        else if (C is THEdit) then begin
          szValue := THEdit(C).Text;
          if (szValue = '') then Continue;
          // Cas du XX_WHERE
          if (Pos('XX_WHERE', UpperCase(szName)) > 0) then begin
            if (Result = '') then Result := 'WHERE '+szValue
                             else Result := Result+' AND ('+szValue+')';
            Continue;
          end;
          Ch := '"';
          case THEdit(C).OpeType of
            otDate   : szValue := USDateTime(StrToDate(szValue));
            otHeure  :;
            otString :;
            otReel   : Ch := #0;
          end;

          case THEdit(C).Operateur of
             Commence : sz := szName+' LIKE '+Ch+szValue+'%'+Ch;
             Egal     : sz := szName+'='+Ch+szValue+Ch;
             Contient : sz := szName+' LIKE '+Ch+'%'+szValue+'%'+Ch;
             Superieur: sz := szName+' >='+Ch+szValue+Ch;
             Inferieur: sz := szName+' <='+Ch+szValue+Ch;
             Different: sz := szName+' <>'+Ch+szValue+Ch;
           end;
          if (Result = '') then Result := 'WHERE '+sz
                           else Result := Result+' AND '+sz;
          Continue;
        end;
      end;
    end;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 23/11/2004
Modifié le ... :   /  /
Description .. : Vérification qd l'utilisateur veut séléctionner la piece si elle
Suite ........ : possède une écriture pointée ou lettée
Suite ........ : SG6 23/11/04 FQ 14227
Mots clefs ... : VERIFICATION,LETTRAGE,POINTAGE
*****************************************************************}
procedure TOF_MODIFENTPIE.FListeFlipSelection(Sender:TObject);
var
  condition_select:boolean;
  condition_check_pointage,condition_check_lettrage:boolean;
  Q,ListeQ:TQuery;
  sSql:string;
 begin

  {$IFDEF EAGLCLIENT}
    ListeQ := TFMul(Ecran).Q.TQ;
    if ListeQ<>nil then ListeQ.Seek(TFMul(Ecran).FListe.Row-1);
    condition_select:=TFMul(Ecran).FListe.IsSelected(TFMul(Ecran).FListe.Row);
  {$ELSE}
    ListeQ := TFMul(Ecran).Q ;
    condition_select:=TFMul(Ecran).FListe.IsCurrentSelected;
  {$ENDIF}

  //si la ligne est sélectionnée (et pas en cours de deselectionnement)
  if condition_select then
  begin
     //Vérification Pointage
     sSql:= 'SELECT MAX(E_REFPOINTAGE) as MAXREF,MAX(E_LETTRAGE) as MAXLETT FROM ECRITURE ';
     sSql:= sSql + 'WHERE E_JOURNAL="' + ListeQ.FindField('E_JOURNAL').AsString + '" AND E_EXERCICE="'+ListeQ.FindField('E_EXERCICE').AsString;
     sSql:= sSql + '" AND E_DATECOMPTABLE="'+usdatetime(ListeQ.FindField('E_DATECOMPTABLE').AsDateTime);
     sSql:= sSql + '" AND E_NUMEROPIECE='+IntToStr(ListeQ.FindField('E_NUMEROPIECE').AsInteger);
     sSql:= sSql + ' AND E_QUALIFPIECE="'+ListeQ.FindField('E_QUALIFPIECE').AsString+'"';
     Q:=OpenSql(sSql,true);
     if not Q.eof then
     begin
       condition_check_pointage:=VarIsNull(Q.FindField('MAXREF').AsVariant);
       condition_check_lettrage:=VarIsNull(Q.FindField('MAXLETT').AsVariant);

       //Vérification pointage
       if (Not(condition_check_pointage) and (Q.FindField('MAXREF').AsString<>'')) or (Q.FindField('MAXREF').AsString<>'') then
       begin
          if not bUseBouton then PGIBox('Cette pièce contient des écritures pointées. Vous ne pouvez pas modifier son entête','Erreur');
          {$IFDEF EAGLCLIENT}
          TFMul(Ecran).FListe.FlipSelection(TFMul(Ecran).FListe.Row);
          {$ELSE}
          TFMul(Ecran).FListe.FlipSelection;
          {$ENDIF}
          Exit;
       end;
       //Vérification lettrage
       if (Not(condition_check_lettrage) and (Q.FindField('MAXLETT').AsString<>'')) or (Q.FindField('MAXLETT').AsString<>'') then
       begin
          if not bUseBouton then PGIBox('Cette pièce contient des écritures lettrées. Vous ne pouvez pas modifier son entête','Erreur');
          {$IFDEF EAGLCLIENT}
          TFMul(Ecran).FListe.FlipSelection(TFMul(Ecran).FListe.Row);
          {$ELSE}
          TFMul(Ecran).FListe.FlipSelection;
          {$ENDIF}
          Exit;
       end;
     end;
     Ferme(Q);
  end;
end;

{$IFNDEF EAGLCLIENT}
procedure TOF_MODIFENTPIE.BSelectAllClick(Sender : TObject);
var
  condition_select:boolean;
  ListeQ:TQuery;
begin
  ListeQ := TFMul(Ecran).Q ;

  if ListeQ.eof then exit;

  bAllSelected:= not bAllSelected;
  bUseBouton := True;

  if bAllSelected then PGIBox('Les pièces contenant des écritures lettrées ou pointées ne seront pas sélectionnées','Avertissement');

  ListeQ.First;

  while (not ListeQ.eof) do
  begin
     condition_select:=TFMul(Ecran).FListe.IsCurrentSelected;

     if condition_select then
     begin
       if not bAllSelected then
       begin
         TFMul(Ecran).FListe.FlipSelection;
       end;
       ListeQ.Next;
       Continue;
     end
     else
     begin
       if not bAllSelected then
       begin
         ListeQ.Next;
         continue;
       end
       else
       begin
         TFMul(Ecran).FListe.FlipSelection;
       end;
     end;

     ListeQ.Next;
   end;

   bUseBouton := False;
   ListeQ.First;
end;
{$ENDIF}

Initialization
  registerclasses ( [ TOF_MODIFENTPIE ] ) ;
END.
