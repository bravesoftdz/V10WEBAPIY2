{***********UNITE*************************************************
Auteur  ...... : Lionel SANTUCCI
Créé le ...... : 16/10/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : REGLACOMPTE ()
Mots clefs ... : TOF;REGLACOMPTE
*****************************************************************}
Unit BTReglDemAcomptes ;

Interface

Uses StdCtrls,               
     Controls,
     Classes,
     M3FP,
     AglInit,
{$IFNDEF EAGLCLIENT}
     fe_main,
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
     Maineagl,
{$ENDIF}
     vierge,
     forms,
     grids,
     graphics,
     windows,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HTB97,
     UTOB,
     HMsgBox,
     utofAfBaseCodeAffaire,
     ENtGC,
     BTPUtil,
     UtilPGI,
     UTOF,
     UTOFGCACOMPTES,
     ParamSoc,
     FactComm,
     FactUtil,
     FactTOB,uEntCommun,UtilTOBPiece
      ;
Type
  TOF_REGLACOMPTE = Class (TOF_AFBASECODEAFFAIRE)
  private
    TOBTiers,TOBAffaire : tob;
    TOBDevis,TOBReglements,TOBPieces,TOBAcomptes_O: TOB;
    GSDEV,GSREGL,GSDAC : THGrid;
    TOTALREGL,TOTALAFFECT : THNumEDIt;
    BSuprRegl,BNewRegl,BValiderRegl,BAffectAuto,BFerme : TToolBarButton97;
    LesColonnesDEV,LesColonnesDAC,LesColonnesREGL : string;
    DEVISSEL : TOB;
    FCancel : boolean;
    //CurVal,CurRegl : string;
    GSDAC_REGL,GSREGL_REGL : integer;
    procedure EvenementsGSDEV(active : boolean);
    procedure EvenementsGSREGL(Active : boolean);
    procedure EvenementsBNewRegl(active : boolean);
    procedure EvenementsBValiderRegl(active : boolean);
    procedure EvenementsGSDAC(active : boolean);
    procedure GetLesCOntroles;
    procedure creelesTOBs;
    procedure DefinilesEvenements;
    procedure definiLignescolonnesGSDEV;
    procedure AffichageGSDEv;
    procedure completeTobDevis;
    procedure GSDEVRowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
    procedure GSDEVEnter(Sender: TObject);
    procedure RemplitDemAcomptes;
    procedure chargeDemAcomptes;
    procedure AffichageDac;
    procedure definiLignescolonnesGSDAC;
    procedure GSDEVBeforeFlip(Sender: TObject; ARow: Integer;var Cancel: Boolean);
    procedure GSDEVFlipSelection(Sender: TObject);
    procedure deselectionne(Arow: integer);
    function  ControleSelection : boolean;
    procedure GSDACRowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
    procedure GSDACRowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
    procedure GSDACCellEnter(Sender: TObject; var ACol, ARow: Integer;var Cancel: Boolean);
    procedure GSDACCellExit(Sender: TObject; var ACol, ARow: Integer;var Cancel: Boolean);
    procedure GSDACEnter(Sender: TObject);
    procedure ChargeLesDevis;
    procedure definiLignescolonnesGSREGL;
    procedure AffichageGSREGL;
    procedure ChargeLesAcomptes;
    procedure InitLesControles;
    procedure ActiveLesGrids;
    procedure DesactiveLesGrids;
    procedure GSREGLEnter(Sender: TObject);
    procedure GSREGLCellEnter(Sender: TObject; var ACol, ARow: Integer;var Cancel: Boolean);
    procedure GSREGLCellExit(Sender: TObject; var ACol, ARow: Integer;var Cancel: Boolean);
    procedure GSDACDblClick (Sender : Tobject);
    procedure GSREGLFlipSelection(Sender: TObject);
    procedure GSREGLBeforeFlip(Sender: TObject; ARow: Integer;var Cancel: Boolean);
    procedure AffichelaligneRegl(TOBL: TOB; Arow: integer);
    procedure CalculeLaSommeRegl;
    procedure GSDACBeforeFlip(Sender: TObject; ARow: Integer;var Cancel: Boolean);
    procedure GSDACFlipSelection(Sender: TObject);
    procedure AddLesSupDac(TOBP: TOB);
    procedure AddLesSupREGL(TOBA: TOB);
    procedure RechargeGrids;
    procedure CalculeLaSommeDAC;
    procedure selectionneREGL(Arow: integer);
    procedure AffichelaligneDac(TOBL: TOB; Arow: integer);
    procedure selectionneDAC(Arow: integer);
    procedure EvenementsBAffectAuto(Active: boolean);
    procedure BAffectAutoClick(Sender: TOBject);
    procedure BnewReglClick(Sender: TObject);
    procedure GSREGLDblClick (Sender : Tobject);
    procedure BSuprReglClick(Sender: TObject);
    procedure EvenementsBSuprRegl(active: boolean);
    procedure BvaliderReglClick(Sender: Tobject);
    procedure EvenementsBFerme(active :boolean);
    procedure BFermeClick(Sender: Tobject);
    procedure ActiveLesBoutons;
    procedure DesactiveLesBoutons;
    procedure ValidelesReglements;
    procedure ChargeLesAdresses(TOBPiece: TOB);
    Procedure GSDACGetCellCanvas ( ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState) ;

  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override ;
  end ;

Implementation

procedure TOF_REGLACOMPTE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_REGLACOMPTE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_REGLACOMPTE.ValidelesReglements;
VAR TOBAcomptes,TOBDAC,TOBACC,TOBA,TOBA_O,TOBAUsed,TOBPieceAcompte,TOBPA : TOB;
    Arow , Indice : integer;
    QQ : TQuery;
    ResteRegl : double;
begin
TOBAcomptes := TOB.create ('LES ACOMPTES',nil,-1);
TOBPieceAcompte := TOB.Create ('LES PIECES',nil,-1);
TRY
For Arow := 1 to GSDAC.rowcount -1 do
    begin
    if GSDAC.IsSelected (Arow) then
       BEGIN

       TOBDAC := TOB(GSDAC.Objects[0,Arow]);
       TOBDAC.putValue ('BST_MONTANTREGL',TOBDAC.GetValue('BST_MONTANTREGL')+TOBDAC.GetValue('AFFECTE'));
       TOBDAC.UpdateDB (true);

       TOBPA := TOB.Create ('PIECE',TOBPieceAcompte,-1);
       QQ := OpenSql ('SELECT * FROM PIECE WHERE GP_NATUREPIECEG="'+TOBDAC.GetValue('BST_NATUREPIECE')+'" AND '
             + ' GP_SOUCHE="'+TOBDAC.GetValue('BST_SOUCHE')+'" AND '
             + ' GP_NUMERO='+inttostr(TOBDAC.GetValue('BST_NUMEROFAC'))+ ' AND GP_INDICEG=0',false,-1,'',true);
       TOBPA.SelectDB ('',QQ);
       ferme (QQ);
       for Indice := 0 to TOBDac.detail.count -1 do
           begin
           TOBACC := TOBDAc.detail[Indice];
           TOBACC.putvalue('GAC_NATUREPIECEG',TOBDAC.GetValue('BST_NATUREPIECE'));
           TOBACC.putvalue('GAC_SOUCHE',TOBDAC.GetValue('BST_SOUCHE'));
           TOBACC.putvalue('GAC_NUMERO',TOBDAC.GetValue('BST_NUMEROFAC'));
           TOBACC.putvalue('GAC_INDICEG',0);
           MajMontants( TobPA,TobACC)  ;

           TOBA := TOBAcomptes.findfirst(['GAC_NATUREPIECEG','GAC_SOUCHE','GAC_NUMERO','GAC_INDICEG','GAC_JALECR','GAC_NUMECR'],
                                         [TOBACC.GetValue('GAC_NATUREPIECEG'),TOBACC.GetValue('GAC_SOUCHE'),
                                         TOBACC.GetValue('GAC_NUMERO'),TOBACC.GEtVAlue('GAC_INDICEG'),TOBACC.GetValue('GAC_JALECR'),
                                         TOBACC.GetValue('GAC_NUMECR')],
                                         true);
           if TOBA = nil then
              begin
              TOBA := TOB.Create ('ACOMPTES',TOBAcomptes,-1);
              TOBA.dupliquer (TOBACC,false,true);
              TOBPA.PutValue('GP_ACOMPTEDEV',TOBPA.GetValue('GP_ACOMPTEDEV')+TOBA.GetValue('GAC_MONTANTDEV'));
              TOBPA.PutValue('GP_ACOMPTE',TOBPA.GetValue('GP_ACOMPTE')+TOBA.GetValue('GAC_MONTANT'));
              end else
              begin
              TOBPA.PutValue('GP_ACOMPTE',TOBPA.GetValue('GP_ACOMPTE')+TOBACC.GetValue('GAC_MONTANT'));
              TOBPA.PutValue('GP_ACOMPTEDEV',TOBPA.GetValue('GP_ACOMPTEDEV')+TOBACC.GetValue('GAC_MONTANTDEV'));
              TOBA.PutValue('GAC_MONTANT',TOBA.GetValue('GAC_MONTANT')+TOBACC.GetValue('GAC_MONTANT'));
              TOBA.PutValue('GAC_MONTANTDEV',TOBA.GetValue('GAC_MONTANTDEV')+TOBACC.GetValue('GAC_MONTANTDEV'));
              end;
           end;
       END;
    end;
if TOBAcomptes.Detail.count  > 0 then
   BEGIN
   TOBAComptes.InsertOrUpdateDB(true);
   TOBPieceAcompte.UpdateDB (true);
   END;
FINALLY
TOBAcomptes.free;
TOBPIeceAcompte.free;
END;
end;

procedure TOF_REGLACOMPTE.OnUpdate ;
begin
  Inherited ;
ValidelesReglements;
end ;

procedure TOF_REGLACOMPTE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_REGLACOMPTE.ChargeLesDevis;
var Q : TQuery;
begin
Q := OpenSql ('SELECT * FROM PIECE WHERE GP_AFFAIRE="'+TOBAffaire.GetValue('AFF_AFFAIRE')
             +'" AND GP_NATUREPIECEG="'+VH_GC.AFnatAffaire
             +'"',true,-1,'',true);
if not Q.eof then TOBDevis.loaddetaildb('PIECE','','',Q,false,true);
ferme(Q);
completeTobDevis;
end;

procedure TOF_REGLACOMPTE.AddLesSupREGL(TOBA : TOB);
begin
TobA.addchampsupValeur('MONTANTINIT',0,False);
TOBA.addchampsupValeur('MONTANTDISPO',0,false) ;
TOBA.addchampsupValeur('NEWACC','',false) ;
end;

procedure TOF_REGLACOMPTE.ChargeLesAcomptes;
var TOBCOmpta,TOBE,TOBA : TOB;
    Achat :       boolean;
    stDebit, stCredit,stCouverture, WhereNatEcr,NotWhereNatEcr :string;
    MtAcompte,MtLettre,MontantDispo : double;
    SQL:String;
    Q:TQuery;
    i : integer;
begin
TOBReglements.clearDetail;
TOBCompta := TOB.Create ('LES ACOMPTES COMPT',nil,-1);
if DEVISSEL = nil then exit;
// champs devise suivant type de saisie de la pièce
if DEVISSEL.GetValue('GP_DEVISE')<> V_PGI.DevisePivot then
begin
  stDebit:='E_DEBITDEV';
  stCredit:='E_CREDITDEV';
  stCouverture:='E_COUVERTUREDEV';
end else if DEVISSEL.GetValue('GP_SAISIECONTRE')='-' then
begin
stDebit:='E_DEBITDEV';
stCredit:='E_CREDITDEV';
stCouverture:='E_COUVERTUREDEV';
end;
(*
else begin stDebit:='E_DEBITEURO';stCredit:='E_CREDITEURO'; stCouverture:='E_COUVERTUREEURO'; end ;
*)
if DEVISSEL.GetValue('GP_VENTEACHAT')='ACH' then Achat:=true else Achat:=False ;
if Achat then begin WhereNatEcr:='(E_NATUREPIECE="OF" OR E_NATUREPIECE="RF")';
                    NotWhereNatEcr:='(E_NATUREPIECE<>"OF" AND E_NATUREPIECE<>"RF")';
              end
         else begin WhereNatEcr:='(E_NATUREPIECE="OC" OR E_NATUREPIECE="RC")' ;
                    NotWhereNatEcr:='(E_NATUREPIECE<>"OC" AND E_NATUREPIECE<>"RC")' ;
              end;
SQL:='SELECT E_JOURNAL,E_NUMEROPIECE,E_AUXILIAIRE,E_NATUREPIECE,E_MODEPAIE,(E_CREDIT-E_DEBIT) as SOLDECR '
    +',('+stCredit+'-'+stDebit+') as SOLDECRDEV'
    //(E_CREDITEURO-E_DEBITEURO) as SOLDECREURO'
    +',(SELECT sum(GAC_MONTANTDEV) from ACOMPTES where GAC_JALECR=ECR.E_JOURNAL and GAC_NUMECR=ECR.E_NUMEROPIECE )'
    +' as MONTANTACOMPTE '
    +',(SELECT sum('+stDebit+'-'+stCredit+') FROM ECRITURE ECC WHERE ECC.E_AUXILIAIRE=ECR.E_AUXILIAIRE AND ECC.E_GENERAL=ECR.E_GENERAL '
    +'         AND ECC.E_QUALIFPIECE=ECR.E_QUALIFPIECE AND ECC.E_ECHE=ECR.E_ECHE AND '+ NotWhereNatEcr+' AND ECC.E_DEVISE=ECR.E_DEVISE  '
    //AND ECC.E_SAISIEEURO=ECR.E_SAISIEEURO '
    +'         AND ECC.E_REFGESCOM="" AND ECC.E_LETTRAGE<>"" AND ECC.E_LETTRAGE=ECR.E_LETTRAGE '
    +' ) as MONTANTLETTRE '
    +' ,E_LIBELLE,E_NUMTRAITECHQ '
    +' FROM ECRITURE ECR '
    +' WHERE E_AUXILIAIRE="'+TobTiers.GetValue('T_AUXILIAIRE')+'" AND E_GENERAL="'+TobTiers.GetValue('T_COLLECTIF')+'" '
    +'       AND E_QUALIFPIECE="N" AND E_ECHE="X" AND '+WhereNatEcr+' AND (E_ETATLETTRAGE="AL" OR E_ETATLETTRAGE="PL") '
    +'       AND E_DEVISE="'+DEVISSEL.GetValue('GP_DEVISE')+'"';
    // AND E_SAISIEEURO="'+DEVISSEL.GetValue('GP_SAISIECONTRE')+'" ';
Q:=OpenSQL(SQL,True,-1,'',true);
TRY
if Not Q.EOF then TOBCompta.LoadDetailDB('','','',Q,True,True) ;

if TOBCompta.Detail.Count >0 then
   begin
   for i:=0 to TOBCompta.Detail.Count-1 do
       BEGIN
       TOBE:=TOBCompta.Detail[i] ;
       mtAcompte:=VarToDouble(TOBE.GetValue('MONTANTACOMPTE')) ;
       mtLettre:= VarToDouble(TOBE.GetValue('MONTANTLETTRE')) ;
       MontantDispo:=TOBE.GetValue('SOLDECRDEV')-mtAcompte-mtLettre;
       If MontantDispo<0 then MontantDispo:=0;
       MontantDispo:=Arrondi(Montantdispo,6);
       TOBA:=TOBReglements.FindFirst(['GAC_JALECR','GAC_NUMECR'],[TOBE.GetValue('E_JOURNAL'),TOBE.GetValue('E_NUMEROPIECE')],False) ;
       if TOBA=Nil then
          BEGIN
          if (TOBE.GetValue('SOLDECRDEV')-mtAcompte-mtLettre)>0 then
             begin
             TOBA:=Tob.create( 'ACOMPTES',TOBReglements,-1) ;
             AddLesSupREGL(TOBA);
             TOBA.putValue('NEWACC','R');
             TOBA.PutValue('MONTANTINIT',TOBE.GetValue('SOLDECRDEV')) ;
             TobA.putValue('GAC_JALECR',TOBE.GetValue('E_JOURNAL'));
             TobA.putValue('GAC_NUMECR',TOBE.GetValue('E_NUMEROPIECE'));
             TobA.putValue('GAC_MONTANTDEV',0);
             TobA.putValue('MONTANTDISPO',MontantDispo);
             TobA.putValue('GAC_MODEPAIE',TOBE.GetValue('E_MODEPAIE'));
             TobA.putValue('GAC_CBLIBELLE','');
             TobA.putValue('GAC_CBINTERNET','');
             TobA.putValue('GAC_DATEEXPIRE','');
             TobA.putValue('GAC_TYPECARTE','');
             TobA.putValue('GAC_NUMCHEQUE','');
             // Modif BTP
             TobA.putValue('GAC_LIBELLE',TOBE.GetValue('E_LIBELLE'));
             TobA.putValue('GAC_NUMCHEQUE',TOBE.GetValue('E_NUMTRAITECHQ'));
             // --
             if copy(TOBE.GetValue('E_NATUREPIECE'),1,1)='R' then TobA.putValue('GAC_ISREGLEMENT','X') else TobA.putValue('GAC_ISREGLEMENT','-');
             end;
          end else
             begin
             TobA.addchampsup('MONTANTDISPO',false) ;
             TobA.addchampsupValeur('MONTANTINIT',TOBE.GetValue('SOLDECRDEV')) ;
             TobA.putValue('MONTANTDISPO',MontantDispo);
             TobA.putValue('GAC_MONTANTDEV',0);
             end;
       end;
   end;
FINALLY
Ferme(Q) ;
TOBCompta.free;
END;
end;

procedure TOF_REGLACOMPTE.OnArgument (S : String ) ;
begin
  Inherited ;
  GSDAC_REGL := -1;
  GSREGL_REGL := -1;
  FCancel := false;
  // Creation des tob utilisation
  creelesTOBs;
  // Rcuperation de tobTiers et TOBAffaire
  TOBAffaire := laTOB;
  TOBTiers := TOB(laTOB.data);
// Recuperation des controls
  GetLesCOntroles;
  // Mise en place des evenements sur les controles
  DefinilesEvenements;
  LesColonnesDEV := 'SELECT;GP_NUMERO;GP_AVENANT;ADRESSE;GP_DATEPIECE;GP_TOTALHTDEV;GP_TOTALTTCDEV;GP_DEVISE' ;
  LesColonnesDAC := 'SELECT;BST_NUMEROSIT;BST_DATESIT;BST_MONTANTTTC;RELIQUAT;AFFECTE' ;
  LesColonnesREGL := 'SELECT;GAC_MODEPAIE;GAC_LIBELLE;GAC_MONTANTDEV;MONTANTDISPO;GAC_ISREGLEMENT' ;

  // chargement de la liste des devis pour l'affaire
  ChargeLesDevis;
  // remplissage des soldes d'acomptes
  ChargeLesAcomptes;
  // Affichage
  TOBAffaire.putecran (ecran);
  TOBTIERS.PUTECRAN (ecran);
  definiLignescolonnesGSDEV;
  definiLignescolonnesGSDAC;
  definiLignescolonnesGSREGL;
  AffichageGSDEv;
  AffichageDac;
  AffichageGSREGL;
end ;

procedure TOF_REGLACOMPTE.OnClose ;
begin
  Inherited ;
  if (Fcancel) and (not ControleSelection) then BEGIN lasterror := 1; Exit; END;
  TOBDevis.free;
  TOBReglements.free;
  TOBPieces.free;
  TOBAComptes_O.FREE;
end ;

(*creation des tobs *)
procedure TOF_REGLACOMPTE.creelesTOBs;
begin
TOBAComptes_O := TOB.Create ('LES ACOMPTES ORIG',nil,-1);
TOBDevis := TOB.create ('Les DEVIS',nil,-1);
TOBReglements := TOB.Create ('LES ACOMPTES',nil,-1);
TOBPieces := TOB.Create ('LES DEMANDES',nil,-1);
end;

(* Recuperation des controles *)

procedure TOF_REGLACOMPTE.InitLesControles;
begin
DesactiveLesGrids;
DesactiveLesBoutons;
end;

procedure TOF_REGLACOMPTE.GetLesCOntroles;
begin
GSDEV := THGrid(GetCOntrol('GSDEV'));
GSREGL := THGrid(GetControl('GSREGL'));
GSDAC := THGrid (GetControl('GSDAC'));
TOTALREGL := THNumEdit(GetCOntrol('TOTALREGL'));
TOTALAFFECT := THNumEDIt(GetControl('TOTALAFFECT'));
BNewRegl := TToolbarButton97 (GetCOntrol('BNEWREGL'));
BSuprRegl := TToolbarButton97 (GetCOntrol('BSUPRREGL'));
BValiderRegl := TToolbarButton97 (GetCOntrol('BVALIDEREGL'));
BAffectAuto := TToolbarButton97 (GetCOntrol('BAFFECTAUTO'));
BFerme := TToolbarButton97 (GetCOntrol('BFERME'));
InitLesControles;
end;

(* EVENEMENTS SUR CONTROLES *)
procedure TOF_REGLACOMPTE.DefinilesEvenements;
begin
EvenementsGSDEV (true);
EvenementsGSREGL (true);
EvenementsGSDAC (true);
EvenementsBNewRegl (true);
EvenementsBsuprRegl (true);
EvenementsBValiderRegl (true);
EvenementsBAffectAuto (true);
EvenementsBFerme (true);
end;

(* GSDEV *)

procedure TOF_REGLACOMPTE.EvenementsGSDEV (active : boolean);
begin
if active then
   begin
   GSDEV.OnRowEnter  := GSDEVRowEnter;
   GSDEV.OnEnter := GSDEVEnter;
   GSDEV.OnBeforeFlip := GSDEVBeforeFlip;
   GSDEV.OnFlipSelection := GSDEVFlipSelection;
   end else
   begin
   GSDEV.OnRowEnter  := nil;
   GSDEV.OnEnter := nil;
   GSDEV.OnBeforeFlip := nil;
   GSDEV.OnFlipSelection := nil;
   end;
end;


procedure TOF_REGLACOMPTE.GSDEVFlipSelection(Sender: TObject);
var Indice : integer;
begin
if not GSDEV.IsSelected (GSDEV.Row) then
   BEGIN
   DEVISSEL := nil;
   RechargeGrids;
   DesactiveLesGrids;
   DesactiveLesBoutons;
   END else
   BEGIN
   //
   DEVISSEL := TOB(GSDEV.Objects[0,GSDEV.row]);
   ActiveLesBoutons;
   ActiveLesGrids;
   RechargeGrids;
   GSREGLEnter (self);
   GSDACEnter (self);
   END;
for Indice := 1 TO GSDEV.rowcount -1 do
    begin
    if Indice = GSDEV.row then continue;
    if GSDEV.isselected (Indice) then deselectionne (indice);
    end;
end;

procedure TOF_REGLACOMPTE.GSDEVBeforeFlip(Sender: TObject; ARow: Integer;var Cancel: Boolean);
begin
if TOBDevis.detail.count = 0 then BEGIN Cancel := true;exit;END;
if not ControleSelection then BEGIN Cancel:= true; exit; END;
end;

function TOF_REGLACOMPTE.ControleSelection : boolean;
begin
result := true;
//
// verification de la tob des acomptes pour demande si abandon des reglements des acomptes
//
if (GSDAC.nbSelected > 0)  then
   BEGIn
   if PGIask (TraduireMemoire('Abandon du pointage des règlements en cours ?'), ecran.caption) <> mryes then result := false;
   END;
end;

procedure TOF_REGLACOMPTE.DesactiveLesGrids;
begin
GSREGL.Enabled := false;
GSDAC.Enabled := false;
end;

procedure TOF_REGLACOMPTE.DesactiveLesBoutons;
begin
BSuprRegl.Enabled := false;
BNewRegl.Enabled := false;
BAffectAUto.Enabled := false;
BValiderRegl.Enabled := false;
end;

procedure TOF_REGLACOMPTE.ActiveLesBoutons;
begin
BSuprRegl.Enabled := true;
BNewRegl.Enabled := True;
BAffectAUto.Enabled := True;
BValiderRegl.Enabled := true;
end;

procedure TOF_REGLACOMPTE.ActiveLesGrids;
begin
GSREGL.Enabled := True;
GSDAC.Enabled := True;
end;

procedure TOF_REGLACOMPTE.RechargeGrids;
begin
TOTALREGL.Value := 0;
TOTALAFFECT.value := 0;
RemplitDemAcomptes;
definiLignescolonnesGSDAC;
AffichageDac;
ChargeLesAcomptes;
definiLignescolonnesGSREGL;
AffichageGSREGL;
end;

procedure TOF_REGLACOMPTE.deselectionne (Arow: integer);
begin
GSDEV.OnBeforeFlip := nil;
GSDEV.OnFlipSelection := nil;
ControleSelection;
GSDEV.FlipSelection (Arow);
GSDEV.onbeforeflip := GSDEVBeforeFlip;
GSDEV.OnFlipSelection := GSDEVFlipSelection;
end;


procedure TOF_REGLACOMPTE.definiLignescolonnesGSDEV;
var i:integer;
    Nam,st : string;
    FF : string;
    NbColonnes : integer;
begin
st := LesColonnesDev;
NbColonnes := 0;
repeat
   Nam:=ReadTokenSt(St) ;
   if nam = '' then break;
   inc(NbCOlonnes);
until nam = '';
GSDEV.ColCount := NbColonnes;
if TOBDevis.detail.count = 0 then GSDEV.RowCount := 2
                             else GSDEV.rowCount := TOBDevis.Detail.count+1;
FF:='#';
if V_PGI.OkDecV>0 then
   begin
    FF:='# ##0.';
    for i:=1 to V_PGI.OkDecV-1 do
       begin
       FF:=FF+'0';
       end;
    FF:=FF+'0';
   end;
st := LesColonnesDev;
i := 0;
repeat
   Nam:=ReadTokenSt(St) ;
   if nam = '' then break;
   if Nam='SELECT' then
      BEGIN
      GSDEV.Cells[i,0]:=' ' ;
      GSDEV.ColFormats[i]:='-1';
      GSDEV.ColWidths [i]:=10;
      END;
   if Nam='GP_NUMERO' then
      begin
      GSDEV.Cells[i,0]:='N° Devis';
      GSDEV.ColFormats[i]:='-1' ;
      GSDEV.ColAligns [i] := taLeftJustify;
      GSDEV.ColWidths [i]:=40;
      end;
   if Nam='GP_AVENANT' then
      begin
      GSDEV.Cells[i,0]:='Avenant';
      GSDEV.ColFormats[i]:='-1' ;
      GSDEV.ColAligns [i] := taCenter;
      GSDEV.ColWidths [i]:=20;
      end;
   if Nam='GP_DATEPIECE' then
      begin
      GSDEV.Cells[i,0]:='Date';
      GSDEV.ColTypes[i]:='D' ;
      GSDEV.ColFormats[i]:=ShortdateFormat ;
      GSDEV.ColAligns [i] := taCenter;
      GSDEV.ColWidths [i]:=50;
      end;
   if Nam='ADRESSE' then
      begin
      GSDEV.Cells[i,0]:='Adresse';
      GSDEV.ColFormats[i]:='-1' ;
      GSDEV.ColAligns [i] := taCenter;
      GSDEV.ColWidths [i]:=110;
      end;
   if Nam='GP_DEVISE' then
      BEGIN
      GSDEV.Cells[i,0]:='Dev.';
      //GSDEV.ColFormats[i]:='CB=BTDEVISE_SYMB' ;
      GSDEV.ColAligns [i] := taCenter;
      GSDEV.ColWidths [i]:=30;
      END;
   if (Nam='GP_TOTALHTDEV') or (Nam = 'GP_TOTALTTCDEV') then
      BEGIN
      if Nam = 'GP_TOTALHTDEV' then GSDEV.Cells[i,0]:='Montant HT';
      if Nam = 'GP_TOTALTTCDEV' then GSDEV.Cells[i,0]:='Montant TTC';
      GSDEV.ColAligns [i] := taRightJustify;
      GSDEV.ColTypes[i]:='R';
      GSDEV.ColFormats[i]:=FF ;
      GSDEV.ColWidths [i]:=80;
      END;
   inc (i);
until nam = '';
end;

procedure TOF_REGLACOMPTE.GSDEVRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
end;

procedure TOF_REGLACOMPTE.GSDEVEnter(Sender: TObject);
var cancel,chg : boolean;
begin
GSDEV.row := 1;
cancel := false;
chg := false;
GSDEVROWEnter (self,1,cancel,chg);
end;

(* GSREGL *)
procedure TOF_REGLACOMPTE.EvenementsGSREGL(Active : boolean);
begin
if active then
   BEGIN
   GSREGL.OnEnter := GSREGLEnter;
   GSREGL.OnBeforeFlip := GSREGLBeforeFlip;
   GSREGL.OnFlipSelection := GSREGLFlipSelection;
   GSREGL.OnCellEnter := GSREGLCellEnter;
   GSREGL.OnCellExit := GSREGLCellExit;
   GSREGL.OnDblClick := GSREGLDblClick;
   END ELSE
   BEGIN
   GSREGL.OnEnter := nil;
   GSREGL.OnBeforeFlip := nil;
   GSREGL.OnFlipSelection := nil;
   GSREGL.OnCellEnter := nil;
   GSREGL.OnCellExit := nil;
   END;
end;

procedure TOF_REGLACOMPTE.selectionneREGL (Arow : integer);
begin
GSREGL.OnBeforeFlip := nil;
GSREGL.FlipSelection (Arow);
GSREGL.OnBeforeFlip := GSREGLBeforeFlip;
end;

procedure TOF_REGLACOMPTE.GSREGLBeforeFlip(Sender: TObject; ARow: Integer;var Cancel: Boolean);
var Indice,AlRow,IndRegl : Integer;
    TOBR,TOBDAC,TOBLDAC,TOBLR : TOB;
    ResteRegl,ResteDac,ValAffecte : double;
begin
TOBR := TOB(GSRegl.objects[0,Arow]); if TOBR = nil then BEGIN Cancel:= true;exit; END;
if GSRegl.IsSelected (ARow) then
   begin
   TOBR.PutValue('GAC_MONTANTDEV',0);
   if TOBR.detail.count > 0 then
      BEGIN
      For Indice := 0 to TOBR.detail.count -1 do // Parcours des demandes d'acomptes associe au reglement
          begin
          TOBLDAC := TOBR.detail[Indice];
          for Alrow :=1 to GSDAC.RowCount -1 do // parcours des demandes d'acomptes pour enlever le montant de reglement
              begin
              TOBDAC := TOB(GSDAC.objects[0,Alrow]);
              if (TOBLDAC.GetValue('BST_NUMEROSIT') = TOBDAC.GetValue('BST_NUMEROSIT')) then
                  BEGIN
                  TOBDAC.PutValue('AFFECTE',TOBDAC.GetValue('AFFECTE')-TOBLDAC.GetValue('AFFECTE'));
                  IndRegl := 0;
                  repeat
                      TOBLR := TOBDAC.detail[IndRegl];
                      if (TOBR.GetValue('GAC_JALECR')=TOBLR.GetValue('GAC_JALECR')) and
                      (TOBR.GetValue('GAC_NUMECR')=TOBLR.GetValue('GAC_NUMECR')) then TOBLR.free
                                                                                 else inc(Indregl);
                  until Indregl >= TOBDAC.detaiL.Count;
                  AffichelaligneDac (TOBDAC,Alrow);
                  if TOBDAC.GetValue('AFFECTE') = 0 then selectionneDAC(Alrow)
                                                    else CalculeLaSommeDAC;
                  break;
                  END;
              end;
          end;
      TOBR.cleardetail;
      AffichelaligneRegl (TOBR,Arow);
      END;
   end else
   begin // Affectation de reglements à la demande d'acompte courante
   for Indice := 1 to GSDAC.RowCount -1 do
       begin
       ResteRegl :=TOBR.GetValue('MONTANTDISPO') - TOBR.GetValue('GAC_MONTANTDEV');
       if (ResteRegl <= 0)  then break;
       TOBDAC := TOB(GSDAC.objects[0,Indice]);
       if TOBDAC.GetValue('ACCREGLE')='X' then Continue;
       ResteDac := TOBDAC.GetValue('RELIQUAT') - TOBDAC.GetValue('AFFECTE');
       if (ResteDac > 0) then
          begin
          if ResteRegl > ResteDac then ValAffecte := ResteDac
                                  else ValAffecte := ResteRegl;
          // Sur la demande d'acompte on affecte le reglement
          TOBDAC.PutValue('AFFECTE',TOBDAC.GetValue('AFFECTE')+ValAffecte);
          // Affectation de la valeur au reglement et mis en place du pointeur sur demande acompte
          TOBR.PutValue('GAC_MONTANTDEV',TOBR.GetValue('GAC_MONTANTDEV')+ValAffecte);
          //
          TOBLDAC := TOBR.FIndFirst(['BST_NUMEROSIT'],[TOBDac.GetValue('BST_NUMEROSIT')],true);
          if TOBLDAC = nil then
             BEGIN
             TOBLDAC := TOB.Create ('BSITUATIONS',TOBR,-1);
             AddLesSupDac (TOBLDAC);
             TOBLDAC.dupliquer (TOBDAC,false,true);
             TOBLDAC.PutValue('AFFECTE',0);
             END;
          TOBLDAC.PutValue('AFFECTE',TOBLDAC.GetValue('AFFECTE')+ValAffecte); // on sauvegarde ainsi le montant affecte a la demande
          //
          TOBLR := TOBDAC.FIndFirst(['GAC_JALECR','GAC_NUMECR'],[TOBR.GetValue('GAC_JALECR'),TOBR.GetValue('GAC_NUMECR')],true);
          if TOBLR = nil then
             begin
             TOBLR := TOB.Create ('ACOMPTES',TOBDAC,-1);
             AddLesSupREGL(TOBLR);
             TOBLR.dupliquer (TOBR,false,true);
             TOBLR.Putvalue ('GAC_MONTANTDEV',0);
             end;
          TOBLR.Putvalue ('GAC_MONTANTDEV',TOBLR.Getvalue ('GAC_MONTANTDEV')+Valaffecte);
          //
          AffichelaligneDAC (TOBDAC,Indice);
          //
          if not GSDAC.ISselected (Indice) then SelectionneDAC (Indice)
                                           else CalculeLaSommeDAC;
          end;
       end;
   if TOBR.GetValue('GAC_MONTANTDEV') = 0 then BEGIN Cancel := true; Exit; END;
   AffichelaligneRegl (TOBR,Arow);
   end;
end;

procedure TOF_REGLACOMPTE.GSREGLFlipSelection(Sender: TObject);
begin
CalculeLaSommeRegl;
end;

procedure TOF_REGLACOMPTE.GSREGLEnter (Sender : TObject);
var Arow, Acol : integer;
    Cancel : boolean;
begin
Acol := GSREGL_REGL;
Arow := 1;
Cancel := false;
GSREGLCellEnter (self,Acol,Arow,cancel);
GSREGL.row := Arow;
GSREGL.Col := ACol;
end;

procedure TOF_REGLACOMPTE.CalculeLaSommeRegl;
var ARow : integer;
    TOBR : TOB;
    Valeur : double;
begin
TOTALRegl.Value := 0;
for ARow := 1 to GSREGL.RowCount-1 do
    begin
    if GSRegl.IsSelected(Arow) then
       begin
       TOBR := TOB(GSREGL.Objects [0,Arow]);
       if TOBR.GetValue('GAC_MONTANTDEV') <> 0 then Valeur := TOBR.GetValue('GAC_MONTANTDEV')
                                               else Valeur := TOBR.GetValue('MONTANTDISPO');
       TOTALREGL.Value := TOTALREGL.Value + Valeur;
       end;
    end;
end;

procedure TOF_REGLACOMPTE.GSREGLDblClick (Sender : Tobject);
var TOBLPIECE,TOBNAcpt,TOBREGL : TOB;
    Laction : string;
begin
if GSREGL.IsSelected (GSREGL.row) then
   BEGIN
   PGIBOx(TraduireMemoire('Impossible de modifier un règlement affecté'),ecran.caption);
   exit;
   END;
TOBREGL := TOB(GSREGL.Objects[0,GSREGL.Row]); if TOBRegl = nil then exit;
if TOBREGL.GetValue('MONTANTINIT') <> TOBREGL.GetValue('MONTANTDISPO') then Laction := 'CONSULTATION'
                                                                       else LAction := 'MODIFICATION';
TOBLPiece := TOB.Create ('PIECE',nil,-1);
TOBNAcpt := TOB.Create ('LES ACOMPTES',nil,-1);

//
TOBLPiece.Dupliquer (DEVISSEL,false,true);
TOBLPiece.PutValue ('GP_UTILISATEUR',V_PGI.User);
TOBLPiece.PutValue ('GP_DATECREATION',V_PGI.DateEntree);
TOBLPiece.PutValue ('GP_DATEPIECE',V_PGI.DateEntree);
TRY
if ModifAcompte (TOBTiers,TOBLPiece,TOBregl,Laction) then
   BEGIn
   TOBREGL.putvalue('GAC_MONTANTDEV',0);
   AffichageGSREGL;
   END;
FINALLY
TOBLPiece.free;TOBNACPt.free;
END;
end;

procedure TOF_REGLACOMPTE.GSREGLCellEnter (Sender: TObject; var ACol, ARow: Integer;var Cancel: Boolean);
begin
(*
if GSREGL.ColLengths [GSREGL.col] = -1 then exit;
CurRegl := GSREGL.Cells [GSREGL.Col,GSREGL.row];
*)
end;

procedure TOF_REGLACOMPTE.GSREGLCellExit (Sender: TObject; var ACol, ARow: Integer;var Cancel: Boolean);
//var TOBL : TOB;
begin
(*
if CurRegl = GSREGL.Cells [ACol,Arow] then exit;
if Acol = GSREGL_REGL then
   BEGIN
   TOBL := TOB(GSREGL.objects[0,Arow]);
   if valeur(GSREGL.cells[Acol,Arow]) > TOBL.GetValue('MONTANTDISPO') then
      begin
      PGIBox (TraduireMemoire ('Impossible : Montant saisi supérieur au montant disponible'),ecran.caption);
      GSREgl.Cells[Acol,Arow] := CurRegl;
      exit;
      end;
   TOBL.PutValue('GAC_MONTANTDEV',valeur(GSREGL.cells[Acol,Arow]));
   AffichelaligneRegl (TOBL,Arow);
   if not GSRegl.IsSelected(Arow) then GSREGLFlipSelection (self);
   CalculeLaSommeRegl;
   END;
*)
end;

procedure TOF_REGLACOMPTE.definiLignescolonnesGSREGL;
var i:integer;
    Nam,st : string;
    FF : string;
    NbColonnes : integer;
begin
EvenementsGSREGL (false);
GSREGL.AllSelected := false;
st := LesColonnesREGL;
NbColonnes := 0;
repeat
   Nam:=ReadTokenSt(St) ;
   if nam = '' then break;
   inc(NbCOlonnes);
until nam = '';
GSREGL.VidePile (false);
GSREGL.ColCount := NbColonnes;
if TOBReglements.Detail.count = 0 then GSREGL.rowcount := 2
                                  else GSREGL.rowCount := TOBReglements.Detail.count+1;
FF:='#';
if V_PGI.OkDecV>0 then
   begin
    FF:='# ##0.';
    for i:=1 to V_PGI.OkDecV-1 do
       begin
       FF:=FF+'#';
       end;
    FF:=FF+'0';
   end;
st := LesColonnesREGL;
i := 0;
repeat
   Nam:=ReadTokenSt(St) ;
   if nam = '' then break;
   if Nam='SELECT' then
      BEGIN
      GSREGL.Cells[i,0]:=' ' ;
      GSREGL.ColFormats[i]:='-1';
      GSREGL.ColWidths [i]:=10;
      GSREGL.ColLengths [i] := -1;
      END;
   if Nam='GAC_MODEPAIE' then
      begin
      GSREGL.Cells[i,0]:='Type';
      GSREGL.ColAligns [i] := taCenter;
      GSREGL.ColFormats[i]:='CB=TTMODEPAIE';
      GSREGL.ColWidths [i]:=40;
      GSREGL.ColLengths [i] := -1;
      end;
   if Nam='GAC_LIBELLE' then
      begin
      GSREGL.Cells[i,0]:='Libellé';
      GSREGL.ColWidths [i]:=80;
      GSREGL.ColLengths [i] := -1;
      end;
   if (Nam='GAC_MONTANTDEV') or (Nam = 'MONTANTDISPO')  then
      BEGIN
      if Nam = 'GAC_MONTANTDEV' then BEGIN GSREGL.Cells[i,0]:='Affecté'; GSREGL_REGL := i; END;
      if Nam = 'MONTANTDISPO' then BEGIN GSREGL.Cells[i,0]:='Reste'; GSREGL.ColLengths [i] := -1; END;
      GSREGL.ColAligns [i] := taRightJustify;
      GSREGL.ColTypes[i]:='R';
      GSDAC.ColFormats[i]:=FF ;
      GSREGL.ColWidths [i]:=60;
      END;
   if (Nam='GAC_ISREGLEMENT')  then
      BEGIN
      GSREGL.ColAligns [i] := taCenter;
      GSREGL.ColTypes[i]:='B';
      GSREGL.ColFormats[i]:=IntToStr(Integer(csCoche));
      GSREGL.ColWidths [i]:=10;
      GSREGL.ColLengths [i] := -1;
      END;
   inc (i);
until nam = '';
EvenementsGSREGL (true);
end;

procedure TOF_REGLACOMPTE.AffichelaligneRegl (TOBL : TOB; Arow : integer);
begin
TOBL.PutLigneGrid (GSREGL,Arow,false,false,LesColonnesREGL);
end;

procedure TOF_REGLACOMPTE.AffichageGSREGL;
var Indice : integer;
TF : TFVierge;
begin
if TOBReglements.detail.count > 0 then
   BEGIN
   GSREGL.RowCount := TOBReglements.detail.count +1;
   for Indice := 0 to TOBReglements.detail.count -1 do
       begin
       AffichelaligneRegl (TOBReglements.detail[Indice],Indice+1);
       GSREGL.Objects[0,Indice+1] := TOBReglements.Detail[Indice];
       end;
   END;
// resize de la grid
TF := TFvierge(ecran);
TF.HMTrad.ResizeGridColumns (GSREGL);
end;

(* GSDAC *)
procedure TOF_REGLACOMPTE.EvenementsGSDAC(active : boolean);
begin
if active then
   BEGIN
   GSDAC.OnEnter := GSDACEnter;
   GSDAC.OnRowEnter  := GSDACRowEnter;
   GSDAC.OnRowExit   := GSDACRowExit;
   GSDAC.OnCellEnter := GSDACCellEnter;
   GSDAC.OnCellExit  := GSDACCellExit;
   GSDAC.OnBeforeFlip := GSDACBeforeFlip;
   GSDAC.OnFlipSelection := GSDACFlipSelection;
   GSDAC.OnDblClick := GSDACDblClick;
   GSDAC.GetCellCanvas  := GSDACGetCellCanvas;
   END ELSE
   BEGIN
   GSDAC.OnEnter := nil;
   GSDAC.OnRowEnter  := nil;
   GSDAC.OnRowExit   := nil;
   GSDAC.OnCellEnter := nil;
   GSDAC.OnCellExit  := nil;
   GSDAC.OnBeforeFlip := nil;
   GSDAC.OnFlipSelection := nil;
   GSDAC.OnDblClick := nil;
   END;
end;

procedure TOF_REGLACOMPTE.GSDACGetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var TOBDAC : TOB;
begin
if ACol<GSDAC.FixedCols then Exit ;
if Arow < GSDAC.FixedRows then exit;
TOBDAC:=TOB(GSDAC.objects[0,Arow]); if TOBDAC=Nil then Exit ;
if TOBDAc.GetValue('ACCREGLE') <> 'X' then exit;
//Canvas.Brush.Color := clInactiveCaption;
Canvas.Font.Style:=Canvas.Font.Style+[fsbold,fsItalic];
Canvas.Font.Color:=clActiveCaption ;
end;

procedure TOF_REGLACOMPTE.GSDACDblClick (Sender : Tobject);
var TOBAcomptes,TOBPiece,TOBACC : TOB;
    TOBA,TOBDAC : TOB;
    QQ : Tquery;
    Sql : string;
begin
TOBDAC := TOB(GSDAC.objects[0,GSDAC.row]);
if TOBDAC = nil then Exit;
TOBAcomptes := TOB.Create ('LES ACOMPTES',nil,-1);
TOBPiece := TOB.create ('PIECE',nil,-1);
//
Sql := 'SELECT * FROM PIECE WHERE GP_NATUREPIECEG="'+TOBDAC.GetValue('BST_NATUREPIECE')+'" AND '+
       'GP_SOUCHE="'+TOBDac.GetValue('BST_SOUCHE')+'" AND '+
       'GP_NUMERO='+inttostr(TOBDAC.GetValue('BST_NUMEROFAC'))+' AND '+
       'GP_INDICEG=0';
QQ := OpenSql (Sql,true,-1,'',true);
TOBPiece.selectDb ('',QQ);
ferme (QQ);
TOBA := TOBAcomptes_O.findfirst(
         ['GAC_NATUREPIECEG','GAC_SOUCHE','GAC_NUMERO'],
         [TOBDAC.GetValue('BST_NATUREPIECE'),TOBDAC.GetValue('BST_SOUCHE'),TOBDAC.GetValue('BST_NUMEROFAC')],
         true);
while TOBA <> nil do
     begin
     TOBACC := TOB.Create ('ACOMPTES',TOBAcomptes,-1);
     TOBACC.dupliquer (TOBA,false,true);
     TobAcc.addchampsupValeur('MONTANTDISPO',0,false) ;
     TOBA := TOBAcomptes_O.findnext(
             ['GAC_NATUREPIECEG','GAC_SOUCHE','GAC_NUMERO'],
             [TOBDAC.GetValue('BST_NATUREPIECE'),TOBDAC.GetValue('BST_SOUCHE'),TOBDAC.GetValue('BST_NUMEROFAC')],
             true);
     end;
if TobAcomptes.detail.count > 0 then ConsultAcompte (TOBTiers,TOBPiece,TOBAcomptes,'CONSULTATION');
TOBAcomptes.free; TOBPiece.free;
end;

procedure TOF_REGLACOMPTE.selectionneDAC (Arow : integer);
begin
GSDAC.OnBeforeFlip := nil;
GSDAC.FlipSelection (Arow);
GSDAC.OnBeforeFlip := GSDACBeforeFlip;
end;


procedure TOF_REGLACOMPTE.AffichelaligneDac (TOBL : TOB; Arow : integer);
begin
TOBL.PutLigneGrid (GSDAC,Arow,false,false,LesColonnesDAC);
end;

procedure TOF_REGLACOMPTE.GSDACBeforeFlip (Sender: TObject; ARow: Integer;var Cancel: Boolean);
var Indice,AlRow,IndRegl : Integer;
    TOBR,TOBDAC,TOBLDAC,TOBLR : TOB;
    ResteRegl,ResteDac,ValAffecte : double;
begin
TOBDAC := TOB(GSDAC.objects[0,Arow]); if TOBDAC = nil then BEGIN Cancel := true; Exit; END;
(* NEW ONE *)
if TOBDAC.GetValue('ACCREGLE')='X' then BEGIN Cancel := true; exit; END;
// ---
if GSDAC.IsSelected (ARow) then
   begin
   TOBDAC.PutValue('AFFECTE',0);
   if TOBDAC.detail.count > 0 then
      BEGIN
      For Indice := 0 to TOBDAC.detail.count -1 do
          begin
          TOBLR := TOBDAC.detail[Indice];
          for Alrow :=1 to GSREGL.RowCount -1 do
              begin
              TOBR := TOB(GSREGL.objects[0,Alrow]); if TOBR = nil then BEGIN CAncel := true; exit; END;
              if (TOBR.GetValue('GAC_JALECR')=TOBLR.GetValue('GAC_JALECR')) and
                 (TOBR.GetValue('GAC_NUMECR')=TOBLR.GetValue('GAC_NUMECR')) then
                  BEGIN
                  TOBR.PutValue('GAC_MONTANTDEV',TOBR.GetValue('GAC_MONTANTDEV')-TOBLR.GetValue('GAC_MONTANTDEV'));
                  IndRegl := 0;
                  repeat
                      TOBLDAC := TOBR.detail[IndRegl];
                      if (TOBLDAC.GetValue('BST_NUMEROSIT') = TOBDAC.GetValue('BST_NUMEROSIT')) then TOBLDAC.free
                                                                                                else inc(Indregl);
                  until Indregl >= TOBR.detaiL.Count;
                  AffichelaligneRegl (TOBR,Alrow);
                  if TOBR.GetValue('GAC_MONTANTDEV') = 0 then selectionneREGL(Alrow)
                                                         else CalculeLaSommeRegl;
                  break;
                  END;
              end;
          end;
      TOBDAC.cleardetail;
      AffichelaligneDac (TOBDAC,Arow);
      END;
   end else
   begin // Affectation de reglements à la demande d'acompte courante
   for Indice := 1 to GSREGL.RowCount -1 do
       begin
       ResteDac := TOBDAC.GetValue('RELIQUAT') - TOBDAC.GetValue('AFFECTE');
       if ResteDac <= 0 then break;
       TOBR := TOB(GSRegl.objects[0,Indice]);  if TOBR = nil then BEGIN Cancel := true; Exit; END;
       ResteRegl :=TOBR.GetValue('MONTANTDISPO') - TOBR.GetValue('GAC_MONTANTDEV');
       if (ResteRegl > 0) then
          begin
          if ResteRegl > ResteDac then ValAffecte := ResteDac
                                  else ValAffecte := ResteRegl;
          // Sur la demande d'acompte on affecte le reglement
          TOBDAC.PutValue('AFFECTE',TOBDAC.GetValue('AFFECTE')+ValAffecte);
          // Affectation de la valeur au reglement et mis en place du pointeur sur demande acompte
          TOBR.PutValue('GAC_MONTANTDEV',TOBR.GetValue('GAC_MONTANTDEV')+ValAffecte);
          //
          TOBLDAC := TOB.Create ('BSITUATIONS',TOBR,-1);
          AddLesSupDac (TOBLDAC);
          TOBLDAC.dupliquer (TOBDAC,false,true);
          TOBLDAC.PutValue('AFFECTE',ValAffecte); // on sauvegarde ainsi le montant affecte a la demande
          //
          TOBLR := TOB.Create ('ACOMPTES',TOBDAC,-1);
          AddLesSupREGL(TOBLR);
          TOBLR.dupliquer (TOBR,false,true);
          TOBLR.Putvalue ('GAC_MONTANTDEV',Valaffecte);
          //
          AffichelaligneRegl (TOBR,Indice);
          //
          if not GSRegl.ISselected (Indice) then SelectionneRegl (Indice)
                                            else CalculeLaSommeRegl;

          end;
       end;
   AffichelaligneDac (TOBDAC,Arow);
   end;
end;

procedure TOF_REGLACOMPTE.GSDACFlipSelection (Sender: TObject);
begin
CalculeLaSommeDAC;
end;


procedure TOF_REGLACOMPTE.CalculeLaSommeDAC;
var ARow : integer;
    TOBDAC : TOB;
begin
TOTALAffect.Value := 0;
for ARow := 1 to GSDAC.RowCount-1 do
    begin
    if GSDAC.IsSelected(Arow) then
       begin
       TOBDAC := TOB(GSDAC.Objects [0,Arow]);
       TOTALAffect.Value := TOTALAffect.Value + TOBDAC.GetValue('AFFECTE');
       end;
    end;
end;

procedure TOF_REGLACOMPTE.GSDACEnter(Sender: TObject);
var Cancel: boolean;
    Arow,Acol : integer;
begin
cancel := false;
Arow := 1;
Acol := GSDAC_REGL;
GSDACRowENter (self,1,cancel,false);
GSDAcCellEnter (Self,Acol,Arow,cancel);
GSDac.row := Arow;
GSDAC.col := Acol;
end;

procedure TOF_REGLACOMPTE.GSDACRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
end;

procedure TOF_REGLACOMPTE.GSDACRowExit (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
end;

procedure TOF_REGLACOMPTE.GSDACCellEnter (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
(*
if GSDAC.ColLengths [GSDAC.col] = -1 then BEGIN cancel := true; EXit; END;
CurVal := GSDAC.Cells [GSDAC.col,GSDAC.row];
*)
end;

procedure TOF_REGLACOMPTE.GSDACCellExit (Sender: TObject; var ACol,ARow : Integer; var Cancel : Boolean);
begin
end;

procedure TOF_REGLACOMPTE.AffichageDac;
var Indice : integer;
TF : TFVierge;
begin
if TOBPieces.detail.count > 0 then
   BEGIN
   GSDAC.RowCount := TOBPIeces.detail.count +1;
   for Indice := 0 to TOBPieces.detail.count -1 do
       begin
       TOBPieces.Detail[Indice].PutLigneGrid (GSDAC,Indice+1,false,false,LesColonnesDAC);
       GSDAC.Objects[0,Indice+1] := TOBPieces.detail[Indice];
       end;
   END;
// resize de la grid
TF := TFVierge(ecran);
TF.HMTrad.ResizeGridColumns (GSDAC);
end;

procedure TOF_REGLACOMPTE.definiLignescolonnesGSDAC;
var i:integer;
    Nam,st : string;
    FF : string;
    NbColonnes : integer;
begin
EvenementsGSDAC (false);
GSDAC.AllSelected := false;
st := LesColonnesDAC;
NbColonnes := 0;
repeat
   Nam:=ReadTokenSt(St) ;
   if nam = '' then break;
   inc(NbCOlonnes);
until nam = '';
GSDAC.VidePile (false);
GSDAC.ColCount := NbColonnes;

if TOBPieces.Detail.count = 0 then GSDAC.rowcount := 2
                              else GSDAC.rowCount := TOBPieces.Detail.count+1;
FF:='#';
if V_PGI.OkDecV>0 then
   begin
    FF:='# ##0.';
    for i:=1 to V_PGI.OkDecV-1 do
       begin
       FF:=FF+'0';
       end;
    FF:=FF+'0';
   end;
st := LesColonnesDAC;
i := 0;
repeat
   Nam:=ReadTokenSt(St) ;
   if nam = '' then break;
   if Nam='SELECT' then
      BEGIN
      GSDAC.Cells[i,0]:=' ' ;
      GSDAC.ColFormats[i]:='-1';
      GSDAC.ColWidths [i]:=10;
      GSDAC.ColLengths [i] := -1;
      END;
   if Nam='BST_NUMEROSIT' then
      begin
      GSDAC.Cells[i,0]:='Num.';
      GSDAC.ColFormats[i]:='-1' ;
      GSDAC.ColAligns [i] := taLeftJustify;
      GSDAC.ColWidths [i]:=40;
      GSDAC.ColLengths [i] := -1;
      end;
   if Nam='BST_DATESIT' then
      begin
      GSDAC.Cells[i,0]:='Date';
      GSDAC.ColTypes[i]:='D' ;
      GSDAC.ColFormats[i]:=ShortdateFormat ;
      GSDAC.ColWidths [i]:=70;
      GSDAC.ColLengths [i] := -1;
      end;
   if (Nam='BST_MONTANTTTC') or (Nam = 'RELIQUAT') or (Nam = 'AFFECTE') then
      BEGIN
      if Nam = 'BST_MONTANTTTC' then BEGIN GSDAC.Cells[i,0]:='Montant TTC'; GSDAC.ColLengths [i] := -1; END;
      if Nam = 'RELIQUAT' then BEGIN GSDAC.Cells[i,0]:='Reste'; GSDAC.ColLengths [i] := -1; END;
      if Nam = 'AFFECTE' then BEGIN GSDAC.Cells[i,0]:='Règlement'; GSDAC_REGL := i; END;
      GSDAC.ColAligns [i] := taRightJustify;
      GSDAC.ColTypes[i]:='R';
      GSDAC.ColFormats[i]:=FF ;
      GSDAC.ColWidths [i]:=111;
      END;
   inc (i);
until nam = '';
EvenementsGSDAC (true);
end;

procedure TOF_REGLACOMPTE.AddLesSupDac (TOBP : TOB);
begin
TOBP.AddChampSupValeur('ACCREGLE','-',false);
TOBP.addchampsupValeur('RELIQUAT',0,false);
TOBP.addchampsupValeur('REGLE',0,false);
TOBP.addchampsupValeur('AFFECTE',0,false);
end;

procedure TOF_REGLACOMPTE.chargeDemAcomptes;
var Sql : string;
    QQ : Tquery;
    Indice : integer;
    TOBP,TOBS : TOB;
    TOBSOLDE : TOB;
    reliquat : double;
begin
TOBAComptes_O.clearDetail;
TOBPIeces.cleardetail;
TOBSolde := TOB.Create ('LES DEMANDES',nil,-1);
if DEVISSEL = nil then exit;
// recuperation des situations de type DAC ayant le meme code sous affaire que le devis
sql := 'SELECT SIT.*,'
       +'(select sum(GAC_MONTANTDEV) FROM ACOMPTES ACC WHERE ACC.GAC_NATUREPIECEG=SIT.BST_NATUREPIECE AND '
       +'ACC.GAC_SOUCHE=SIT.BST_SOUCHE AND ACC.GAC_NUMERO=SIT.BST_NUMEROFAC) as AFFECTE '
       +'FROM BSITUATIONS SIT WHERE BST_SSAFFAIRE="'+ DEVISSEL.GetValue('GP_AFFAIREDEVIS') +'" '
       +'AND BST_VIVANTE="X" ORDER BY BST_NUMEROSIT';
QQ := OpenSql (sql,true,-1,'',true);
TRY
if not QQ.eof then
   begin
   TOBSolde.LoadDetailDB ('','','',QQ,false,true);
   for Indice := 0 to TOBSolde.detail.count -1 do
      begin
      TOBS := TOBSolde.detail[Indice];
      reliquat := TOBS.GetValue ('BST_MONTANTTTC') - VarToDouble(TOBS.GetValue ('AFFECTE'));
      (*
      if (reliquat > 0)   then
         begin
      *)
         TOBP := TOB.Create ('BSITUATIONS',TOBPieces,-1);
         AddLesSupDac (TOBP);
         TOBP.PutValue ('RELIQUAT',reliquat);
         TOBP.Putvalue ('BST_NATUREPIECE',TOBS.GetValue('BST_NATUREPIECE'));
         TOBP.Putvalue ('BST_SOUCHE',TOBS.GetValue('BST_SOUCHE'));
         TOBP.Putvalue ('BST_NUMEROFAC',TOBS.GetValue('BST_NUMEROFAC'));
         TOBP.Putvalue ('BST_NUMEROSIT',TOBS.GetValue('BST_NUMEROSIT'));
         TOBP.Putvalue ('BST_DATESIT',TOBS.GetValue('BST_DATESIT'));
         TOBP.Putvalue ('BST_AFFAIRE',TOBS.GetValue('BST_AFFAIRE'));
         TOBP.Putvalue ('BST_SSAFFAIRE',TOBS.GetValue('BST_SSAFFAIRE'));
         TOBP.Putvalue ('BST_MONTANTHT',TOBS.GetValue('BST_MONTANTHT'));
         TOBP.Putvalue ('BST_MONTANTTVA',TOBS.GetValue('BST_MONTANTTVA'));
         TOBP.Putvalue ('BST_MONTANTTTC',TOBS.GetValue('BST_MONTANTTTC'));
         TOBP.Putvalue ('BST_MONTANTREGL',TOBS.GetValue('BST_MONTANTREGL'));
         if reliquat <= 0 then TOBP.Putvalue ('ACCREGLE','X');
         Sql := 'SELECT * FROM ACOMPTES WHERE GAC_NATUREPIECEG="'+TOBP.GetValue('BST_NATUREPIECE')+'" AND '
             + 'GAC_SOUCHE="'+TOBP.GetValue('BST_SOUCHE')+'" AND '
             + 'GAC_NUMERO='+inttostr(TOBP.GetValue('BST_NUMEROFAC'));
         TOBAcomptes_O.LoadDetailFromSQL (SQL,true,true);
      (*
         end;
      *)
      end;
   end;
FINALLY
ferme (QQ);
TOBSolde.free;
END;
end;

procedure TOF_REGLACOMPTE.RemplitDemAcomptes;
begin
chargeDemAcomptes;
end;

procedure TOF_REGLACOMPTE.EvenementsBSuprRegl(active : boolean);
begin
BsuprRegl.OnClick := BsuprReglClick;
end;

procedure TOF_REGLACOMPTE.BSuprReglClick (Sender : TObject);
var TOBRegL : TOB;
begin
if GSREGL.IsSelected (GSREGL.row) then
   BEGIN
   PgiBox (TraduireMemoire ( 'veuillez déselectionner le règlement'),ecran.caption);
   exit;
   END;
TOBREGl := TOB(GSRegl.objects[0,GSRegl.row]); if TOBRegl = nil then exit;
if TOBRegl.GetValue ('MONTANTINIT') <> TOBRegl.GetValue ('MONTANTDISPO') then
   begin
   PgiBox (TraduireMemoire ('Impossible : Règlement affecté'),ecran.caption);
   exit;
   end;
if PGiAsk (TraduireMemoire ('Etes-vous sur de vouloir supprimer ce règlement ?'),ecran.caption) <> mryes then exit;
SuprimeAcompte (DEVISSEL,TOBRegl);
TOBRegl.free;
AffichageGSREGL;
end;

procedure TOF_REGLACOMPTE.EvenementsBNewRegl(active : boolean);
begin
BnewRegl.OnClick := BnewReglClick;
end;

procedure TOF_REGLACOMPTE.BnewReglClick (Sender : TObject);
var TOBACC,TOBLPIece : TOB;
begin
TOBACC := TOB.Create ('DES ACOMPTES',nil,-1);
TOBLPiece := TOB.Create ('PIECE',nil,-1);
AddLesSupEntete (TOBLPIece);
TOBLPiece.Dupliquer (DEVISSEL,false,true);
TOBLPiece.PutValue ('GP_UTILISATEUR',V_PGI.User);
TOBLPiece.PutValue ('GP_DATECREATION',V_PGI.DateEntree);
TOBLPiece.PutValue ('GP_DATEPIECE',V_PGI.DateEntree);
TRY
if NewAcompte (TOBTiers,TOBLPiece,TOBAcc) then
   BEGIn
   TOBACC.detail[0].putvalue('GAC_MONTANTDEV',0);
   TOBACC.detail[0].ChangeParent (TOBReglements,-1);
   AffichageGSREGL;
   END;
FINALLY
TOBACC.free; TOBLPiece.free;
END;
end;

procedure TOF_REGLACOMPTE.BvaliderReglClick (Sender : Tobject);
begin
ValidelesReglements ;
GSDEV.OnBeforeFlip := nil;
GSDEV.FlipSelection (GSDEV.row);
GSDEV.OnBeforeFlip := GSDEVBeforeFlip;
end;

procedure TOF_REGLACOMPTE.EvenementsBValiderRegl(active : boolean);
begin
BValiderRegl.OnClick := BvaliderReglClick;
end;

procedure TOF_REGLACOMPTE.AffichageGSDEv;
var indice : integer;
    TF : TFvierge;
begin
if TOBDevis.detail.count = 0 then exit;
for Indice := 0 to TOBDevis.detail.count -1 do
    begin
    TOBDevis.Detail[Indice].PutLigneGrid (GSDEV,Indice+1,false,false,LesColonnesDev);
    GSDEV.Objects[0,Indice+1] := TOBDEVIS.detail[Indice];
    end;
// resize de la grid
TF := TFvierge(ecran);
TF.HMTrad.ResizeGridColumns (GSDEV);
end;

procedure TOF_REGLACOMPTE.ChargeLesAdresses (TOBPiece : TOB);
var Cledoc : R_CLEDOC;
    Q : Tquery;
begin
if GetParamSoc('SO_GCPIECEADRESSE') then
  BEGIN
  CleDoc:=TOB2CleDoc(TOBPiece) ; 
  Q:=OpenSQL('SELECT GPA_ADRESSE1 FROM PIECEADRESSE WHERE '+WherePiece(CleDoc,ttdPieceAdr,False)+' ORDER BY GPA_TYPEPIECEADR',True,-1,'',true);
  if not Q.eof then TOBPiece.putvalue('ADRESSE',Q.Findfield('GPA_ADRESSE1').AsString );
  Ferme(Q) ;
  END else
  BEGIN
  Q:=OpenSQL('SELECT ADR_ADRESSE1 FROM ADRESSES WHERE ADR_NUMEROADRESSE='+IntToStr(TOBPiece.GetValue('GP_NUMADRESSELIVR')),True,-1,'',true);
  if not Q.eof then TOBPiece.putvalue('ADRESSE',Q.Findfield('ADR_ADRESSE1').AsString );
  ferme (Q);
  END ;
end;

procedure TOF_REGLACOMPTE.completeTobDevis;
var Indice : integer;
    TOBD : TOB;
    TypeFacturation : String;
begin
Indice := 0;
repeat
    TOBD := TOBDEVIS.Detail [Indice];
    TypeFacturation := RenvoieTypeFact(TOBD.GetValue('GP_AFFAIREDEVIS'));
    if Typefacturation = 'DIR' then
       BEGIN
       TOBD.free
       END else
       BEGIN
       TOBD.AddChampSupValeur ('ADRESSE','',false);
       ChargeLesAdresses (TOBD);
       Inc (Indice);
       END;
until Indice >= TOBDevis.detail.count;
end;

(* BAFFECTAUTO *)
procedure TOF_REGLACOMPTE.EvenementsBAffectAuto (Active : boolean);
begin
BAffectAuto.onclick := BAffectAutoClick;
end;

procedure TOF_REGLACOMPTE.BAffectAutoClick (Sender : TOBject);
var ARow : integer;
begin
if GSREGL.nbSelected >  0 then
   BEGIN
   PGIBox (TraduireMemoire ('Des affectations ont déjà été effectuées'),ecran.caption);
   exit;
   END;
For Arow := 1 to GSREGL.RowCount -1 do
    BEGIN
    GSRegL.FlipSelection (Arow);
    END;
end;

procedure TOF_REGLACOMPTE.BFermeClick (Sender : Tobject);
begin
Fcancel := true;
end;

procedure TOF_REGLACOMPTE.EvenementsBFerme(active :boolean);
begin
BFerme.OnClick := BFermeClick;
end;

procedure TOF_REGLACOMPTE.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('AFF_AFFAIRE'));
Aff1:=THEdit(GetControl('AFF_AFFAIRE1')); Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3')); Aff4:=THEdit(GetControl('AFF_AVENANT'));
Tiers:=THEdit(GetControl('AFF_TIERS'));
end;

Initialization
 registerclasses ( [ TOF_REGLACOMPTE ] ) ;
end.

