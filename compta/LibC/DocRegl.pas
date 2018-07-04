{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 15/04/2003
Modifié le ... : 09/05/2007
Description .. : 15/04/2003 : Passage en eAGL
Suite ........ : 
Suite ........ : 27/06/2003 : Correction termporaire du pb avec champs 
Suite ........ : DR_NUMCHEQUE et avec Agl
Suite ........ : 
Suite ........ : JP 21/02/05 : Gestion de la banque prévisionnelle
Suite ........ : 
Suite ........ : JP 09/05/07 : FQ 20075 : pd sur les tablettes pour les états
Mots clefs ... : 
*****************************************************************}
unit DocRegl;

interface

uses
  Controls, ExtCtrls, StdCtrls, Buttons,
  Hctrls, HSysMenu, hmsgbox, HTB97, Mask,
  SysUtils,   // StrToInt, IntToStr, DateToStr, Trim, StrToDate
  Classes,
  Graphics,   // clBtnFace, clWindow
  Forms,
  EdtEtat,    // EditEtat
{$IFDEF EAGLCLIENT}
  UtileAGL,   // LanceEtat, LanceDocument, LanceDocumentSpool, TestDocument
{$ELSE}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  EdtREtat,   // LanceEtat
  EdtRDoc,    // LanceDocument, LanceDocumentSpool, TestDocument
  uPDFBatch,  // pour gestion du spooler
{$ENDIF}
  HPdfPrev,   // pour gestion du spooler
  Grids,
  HEnt1,      // V_PGI,
  SaisUtil,   // TSorteLettre, TSuiviMP, RDEVISE, TOBM, RMVT
  SaisComm,   // DecodeLC, WhereEcriture
  Ent1,       // VH
  NumChek,    // SaisieNumCheque
  Lettutil,   // FormatCheque
  ValideDR,   // ValideImpDR
  uTOB,       // TOB
  ed_Tools ;

Type tDevDocRegl=(onLocale,OnEuro,OnDevise) ;

function LanceDocRegl(tt : TList ; SorteLettre : TSorteLettre ; ModeleTraite,GroupeEncadeca : String ; SuiviMP : Boolean ;
                      Var MSED : tMSEncaDeca ; OkPrintDialog : Boolean = TRUE ; OnDev : tDevDocRegl = OnLocale ;
                      smp : TSuiviMP = smpAucun ; TOBParamEsc : TOB = nil ; vBoReEdit : Boolean = False ) : Integer ;

function LanceDocReglTID(tt : TList ; SorteLettre : TSorteLettre ; ModeleTraite,GroupeEncadeca : String ; SuiviMP : Boolean ;
                         Var MSED : tMSEncaDeca ; OkPrintDialog : Boolean = TRUE ; OnDev : tDevDocRegl = OnLocale ;
                         smp : TSuiviMP = smpAucun ; TOBParamEsc : TOB = nil ; vBoReEdit : Boolean = False ) : Integer ;

Const colTiers = 0 ;
      colLibelle = 1 ;
      colMontantPIVOT = 2 ;
      colNumCheque = 3 ;
      colOrigines = 4 ;
      colPrevi = 5 ;
      colMontantDEV = 6 ;
      colDateEche = 7 ;
      colCollectif = 8 ;

type
  TFDocRegl = class(TForm)
    pBoutons: TPanel;
    bAnnuler: THBitBtn;
    BAide: THBitBtn;
    BImprimer: THBitBtn;
    HMTrad: THSystemMenu;
    BValider: THBitBtn;
    bApercu: THBitBtn;
    bTest: THBitBtn;
    Msg: THMsgBox;
    P: TPanel;
    GE: THGrid;
    P2: TPanel;
    G: THGrid;
    Panel2: TPanel;
    Panel1: TPanel;
    bDown: TToolbarButton97;
    bUp: TToolbarButton97;
    bGroupTiers: TToolbarButton97;
    bCopyDetail: TToolbarButton97;
    HLabel4: THLabel;
    HLabel5: THLabel;
    HLabel6: THLabel;
    HCompte: THLabel;
    HBanque: TPanel;
    HModele: TPanel;
    Panel7: TPanel;
    HLabel1: THLabel;
    HTitreImprime: THLabel;
    bModele: TToolbarButton97;
    Action: TLabel;
    E_MODELE: THValComboBox;
    CompteBanque: THCritMaskEdit;
    HLabel2: THLabel;
    bGroupTiersEche: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure bApercuClick(Sender: TObject);
    procedure bTestClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure bModeleClick(Sender: TObject);
    procedure bDownClick(Sender: TObject);
    procedure bUpClick(Sender: TObject);
    procedure bCopyDetailClick(Sender: TObject);
    procedure bGroupTiersEcheClick(Sender: TObject);
    procedure bGroupTiersClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure bAnnulerClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure E_JOURNALChange(Sender: TObject);
    procedure E_MODELEChange(Sender: TObject);
    procedure CompteBanqueDblClick(Sender: TObject);
    procedure CompteBanqueExit(Sender: TObject);
    procedure GSorted(Sender: TObject);
  private
    { Déclarations privées }
    tt : TList ;
    Modele,General,GeneRIB,GeneDevise,BqDevise : string ;
    ChqPrinted,ForcerGroupe : boolean ;
    LastOrigine : Integer ;
    GeneCoef    : Double ;
    DEV         : RDEVISE ;
    OkPrintDialog : Boolean ;
    ModeleMultiSession : String ;
    MSED : TMSEncaDeca ;
    TIDTIC : Boolean ;
    OnDev    : tDevDocRegl ;
    ModeleOk : Boolean ;
    bEtatAGL : Boolean ;
    procedure RazTemp ;
    procedure GenererTablesTemp ;
    procedure AfficheGridEcheances ;
    procedure AjouteEcheGrid(O : TOBM ; lig : Integer) ;
    function  LibelleTiers(Aux : string) : string ;
    procedure PutDR_NUMCHEQUE(snChq,sRef : String) ;
    procedure ImprimeDocRegl(Selection,Test,Apercu : boolean ; AvecImpr : boolean = True) ;
    function  RequeteDocRegl(l : integer) : string ;
    procedure AttributNumCheque(Ligne : Integer ; sRef,sNumCheque : String ; Affiche : boolean) ;
    procedure ChercheModele ( Charge : boolean ) ;
    procedure EcheToCheque( iEche, iCheq : Integer ; VoirErreurs, ControleNeg : Boolean ) ;
    procedure ChequeToEche(iCheq : Integer) ;
    procedure InitCheques ;
    function  NoMoreEche : Boolean ;
    procedure RenseigneRIB ;
    procedure CompteToBanque ( Charge : boolean ) ;
    Function  InitE_MODELE : Boolean ;
//    function GetNbPages(Const SQL : String) : Integer;
    procedure SetNumChqDansEche( vStNumChq : String ; vInChq : Integer; MajBqPrevi : Boolean = False) ;
  public
    SorteLettre  : TSorteLettre ;
    ModeleTraite,GroupeEncadeca : String ;
    SuiviMP                     : Boolean ;
    LaColMontant                : Integer ;
    TOBParamEsc									: TOB ;
    smp                         : TSuiviMP ;
    FBoReEdit                   : Boolean ;
    FLargeur                    : Integer; {JP 05/6/07 : FQ 18664}
  end;

implementation


{$R *.DFM}

uses
  {$IFNDEF EAGLCLIENT}
  UProcGen, {CopierFichier -> JP 04/06/07 : FQ 19507}
  {$ENDIF}
  UtilPGI, {EncodeRIB}
  ULibWindows, 
  ParamSoc;


{***********A.G.L.***********************************************
Auteur  ...... : Compta
Créé le ...... : 05/11/2004
Modifié le ... : 05/11/2004
Description .. : Y'a longtemps.... : Création du module...
Suite ........ : 
Suite ........ : 05/11/2004 : Ajout d'un paramètre spécifiant la réédition 
Suite ........ : uniquement des documents ( FQ14857 SBO)
Mots clefs ... : 
*****************************************************************}
function LanceDocRegl(tt : TList ; SorteLettre : TSorteLettre ; ModeleTraite,GroupeEncadeca : String ; SuiviMP : Boolean ;
                      Var MSED : tMSEncaDeca ; OkPrintDialog : Boolean = TRUE ; OnDev : tDevDocRegl = OnLocale ;
                      smp : TSuiviMP = smpAucun ; TOBParamEsc : TOB = nil ; vBoReEdit : Boolean = False ) : Integer ;
var FDocRegl : TFDocRegl ;
    i : integer ;
    Q : TQuery ;
begin
result:=-1 ;
FDocRegl := TFDocRegl.Create(Application) ;
   try
   FDocRegl.tt:=tt ;
   FDocRegl.OnDev:=OnDev ;
   FDocRegl.SorteLettre:=SorteLettre ;
   FDocRegl.ModeleTraite:=ModeleTraite ;
   FDocRegl.GroupeEncadeca:=GroupeEncadeca ;
   FDocRegl.SuiviMP:=SuiviMP ;
   FDocRegl.OkPrintDialog:=OkPrintDialog ;
   FDocRegl.MSED:=MSED ;
   FDocRegl.TIDTIC:=FALSE ;
   FDocRegl.TOBParamEsc := TOBParamEsc ;
   FDocRegl.smp := smp ;
   FDocRegl.FBoReEdit := vBoReEdit ;  // Mode réédition // FQ 14857 SBO 05/11/2004
   if FDocRegl.ShowModal=mrOK then
      begin
      i:=0 ;
      Q:=OpenSQL('SELECT COUNT(DR_AUXILIAIRE) FROM DOCREGLE WHERE DR_USER="'+V_PGI.User+'" AND DR_EDITE="X"',True) ;
      if not Q.EOF then i:=Q.Fields[0].AsInteger ;
      Ferme(Q) ;
      result:=i ;
      end ;
   finally
   MSED:=FDocRegl.MSED ;
   FDocRegl.Free ;
   end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Compta
Créé le ...... : 05/11/2004
Modifié le ... : 05/11/2004
Description .. : Y'a longtemps.... : Création du module...
Suite ........ : 
Suite ........ : 05/11/2004 : Ajout d'un paramètre spécifiant la réédition 
Suite ........ : uniquement des documents ( FQ14857 SBO)
Mots clefs ... : 
*****************************************************************}
function LanceDocReglTID(tt : TList ; SorteLettre : TSorteLettre ; ModeleTraite,GroupeEncadeca : String ; SuiviMP : Boolean ;
                      Var MSED : tMSEncaDeca ; OkPrintDialog : Boolean = TRUE ; OnDev : tDevDocRegl = OnLocale ;
                      smp : TSuiviMP = smpAucun ; TOBParamEsc : TOB = nil ; vBoReEdit : Boolean = False ) : Integer ;
var
  FDocRegl : TFDocRegl ;
    i : integer ;
    Q : TQuery ;
begin
result:=-1 ;
FDocRegl := TFDocRegl.Create(Application) ;
   try
   FDocRegl.tt:=tt ;
   FDocRegl.OnDev:=OnDev ;
   FDocRegl.SorteLettre:=SorteLettre ;
   FDocRegl.ModeleTraite:=ModeleTraite ;
   FDocRegl.GroupeEncadeca:=GroupeEncadeca ;
   FDocRegl.SuiviMP:=SuiviMP ;
   FDocRegl.OkPrintDialog:=OkPrintDialog ;
   FDocRegl.MSED:=MSED ;
   FDocRegl.TIDTIC:=TRUE ;
   FDocRegl.TOBParamEsc := TOBParamEsc ;
   FDocRegl.smp := smp ;
   FDocRegl.FBoReEdit := vBoReEdit ;    // Mode réédition // FQ 14857 SBO 05/11/2004
   if FDocRegl.ShowModal=mrOK then
      begin
      i:=0 ;
      Q:=OpenSQL('SELECT COUNT(DR_AUXILIAIRE) FROM DOCREGLE WHERE DR_USER="'+V_PGI.User+'" AND DR_EDITE="X"',True) ;
      if not Q.EOF then i:=Q.Fields[0].AsInteger ;
      Ferme(Q) ;
      result:=i ;
      end ;
   finally
   MSED:=FDocRegl.MSED ;
   FDocRegl.Free ;
   end ;
end ;

Function TFDocRegl.InitE_MODELE : Boolean ;
var O : TOBM ;
    Journal : String ;
    Q : TQuery ;
    LibMod,CodeMod,Dev,St,St1 : String ;
    NatModele : String ;
    TypModele : String ;
    OkOk : Boolean ;
    TOBB,TOBL : Tob ;
    i : Integer ;
BEGIN
  Result := FALSE ;

  If OnDev=OnLocale Then DEV:=V_PGI.DevisePivot
    Else If OnDev=OnDevise Then DEV:='DEV'
    Else DEV:=V_PGI.DeviseFongible;

// Gestion des éditions états / documents : Récupération des modèles paramétrés dans les banques
//                                          uniquement si utilisation du générateur de document
  if bEtatAGL then
    begin
    TypModele := 'E' ;
    If SorteLettre=tslCheque Then NatModele:='CLC'
      Else If SorteLettre=tslVir Then NatModele:='CLV'
      Else If SorteLettre=tslPre Then NatModele:='CLP'
      Else Exit ;
    end
  else
    begin
    TypModele := 'L' ;
    If SorteLettre=tslCheque Then NatModele:='LCH'
      Else If SorteLettre=tslVir Then NatModele:='LVI'
      Else If SorteLettre=tslPre Then NatModele:='LPR'
      Else Exit ;
    end ;

O:=TOBM(tt[0]) ; Journal:=O.GetMvt('E_JOURNAL') ;
St:=Journal+';'+Dev+';' ;
Q:=OpenSQL('SELECT MO_CODE,MO_LIBELLE FROM MODELES WHERE MO_TYPE="' + TypModele + '" AND MO_NATURE="'+NatModele+'" AND MO_LIBELLE LIKE "'+Journal+';%" ',TRUE) ;
TOBB:=TOB.Create('',Nil,-1) ;
TOBB.LoadDetailDB('MODELES','','',Q,False,True) ;
Ferme(Q) ;
OkOk:=TOBB.Detail.Count>0 ;
If Not OkOk Then
  BEGIN
  Q:=OpenSQL('SELECT MO_CODE,MO_LIBELLE FROM MODELES WHERE MO_TYPE="' + TypModele + '" AND MO_NATURE="'+NatModele+'" ',TRUE) ;
  TOBB.LoadDetailDB('MODELES','','',Q,False,True) ;
  Ferme(Q) ;
  END ;
OkOk:=TOBB.Detail.Count>0 ;
If Not OkOk Then
  BEGIN
  TOBB.Free ; Exit ;
  END ;

E_MODELE.DataType:='' ; E_MODELE.Items.Clear ; E_MODELE.Values.Clear ; E_MODELE.Plus:='';
St1:='' ;
For i:=0 To TOBB.Detail.Count-1 Do
  BEGIN
  TOBL:=TOBB.Detail[i] ; LibMod:=TOBL.GetValue('MO_LIBELLE') ; CodeMod:=TOBL.GetValue('MO_CODE') ;
  E_MODELE.Values.Add(CodeMod) ; E_MODELE.Items.Add(LibMod) ;
  If Copy(LibMod,1,Length(St))=St Then St1:=CodeMod ;
  END ;
If St1<>'' Then E_MODELE.Value:=St1 Else E_MODELE.Value:=E_MODELE.Values[0] ;
TOBB.Free ;
Result:=TRUE ;
END ;

procedure TFDocRegl.CompteToBanque ( Charge : boolean ) ;
Var Q : TQuery;
begin

  if CompteBanque.Text<>'' then
    begin

    Q := OpenSQL('SELECT BQ_CODE,BQ_LETTRECHQ,BQ_LETTREPRELV,BQ_LETTREVIR,BQ_LETTRELCR,BQ_LIBELLE,BQ_DEVISE'
                +' FROM BANQUECP WHERE BQ_GENERAL="'+CompteBanque.Text+'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"',True); // 19/10/2006 YMO Multisociétés

    if (not Q.EOF) then
      begin
      HBanque.Caption := Q.FindField('BQ_CODE').AsString + ' - ' + Q.FindField('BQ_LIBELLE').AsString;
      BQDevise := Q.FindField('BQ_DEVISE').AsString;

      // Gestion des éditions états / documents : Récupération des modèles paramétrés dans les banques
      //                                          uniquement si utilisation du générateur de document
      {JP 07/06/07 : FQ 19591 : Activation pour le générateur d'état
      if not bEtatAGL then
        begin}

        if (SorteLettre = tslCheque) then
          begin
          // BPY le 08/01/2004 => correction a la demande de la FFF
          Modele := Q.FindField('BQ_LETTRECHQ').AsString;
          If (Not ModeleOk)
            Then HModele.Caption := Modele + ' - ' + RechDom('ttModeleLettreCHQ',Modele,False)
            else E_MODELE.Value := Modele;
          // fin BPY
          end
        // BPY le 21/01/2004 => point 88 de la FFF
        else If (SorteLettre = tslPre) then
          begin
          Modele := Q.FindField('BQ_LETTREPRELV').AsString;
          If (Not ModeleOk)
            Then HModele.Caption := Modele + ' - ' + RechDom('TTMODELELETTREPRE',Modele,False)
            else E_MODELE.Value := Modele;
          end
        else if (SorteLettre = tslVir) Then
          begin
          Modele := Q.FindField('BQ_LETTREVIR').AsString;
          If (Not ModeleOk)
            Then HModele.Caption := Modele + ' - ' + RechDom('TTMODELELETTREVIR',Modele,False)
            else E_MODELE.Value := Modele;
          end
        else if (SorteLettre = tslTraite(*???*)) Then
          begin
          Modele := Q.FindField('BQ_LETTRELCR').AsString;
          If (Not ModeleOk)
            Then HModele.Caption := Modele + ' - ' + RechDom('TTMODELELETTRETRA',Modele,False)
            else E_MODELE.Value := Modele;
          end
        // Fin BPY
        else
          begin
          If (SorteLettre<>tslPre) And (SorteLettre<>tslVir) and Charge then
            begin
            Modele := ModeleTraite;
            E_MODELE.Value := Modele;
            end;
          end;

        //end;
      end ;
      Ferme(Q) ;
    end
  else
    begin
    HBanque.Caption:='' ;
    HCompte.Caption:='' ;
    end;

end;

procedure TFDocRegl.ChercheModele ( Charge : boolean ) ;
var O : TOBM ;
    Q : TQuery ;
    Journal,CptGene : String ;
    Trouv : boolean ;
BEGIN
  Trouv:=False ;
  if (SorteLettre=tslCheque) Or (SorteLettre=tslVir) Or (SorteLettre=tslPre) then
  begin
    if Charge And (Not ModeleOk) then Modele:= '' ;
    if tt.Count=0 then exit ;
    O:=TOBM(tt[0]) ;
    Journal:=O.GetMvt('E_JOURNAL') ;
    Q:=OpenSQL('SELECT J_CONTREPARTIE,J_LIBELLE,G_LIBELLE FROM JOURNAL LEFT JOIN GENERAUX ON G_GENERAL=J_CONTREPARTIE WHERE J_JOURNAL="'+journal+'" AND J_NATUREJAL="BQE"',True) ;
    if not Q.EOF then
    begin
      CompteBanque.Text:=Q.FindField('J_CONTREPARTIE').AsString ;
      HCompte.Caption:=Q.FindField('G_LIBELLE').AsString ;
      Trouv:=True ;
    end ;
    Ferme(Q) ;
    if Not Trouv then
    begin
      CptGene:=O.GetMvt('E_CONTREPARTIEGEN') ;
      Q:=OpenSQL('Select G_LIBELLE from GENERAUX Where G_GENERAL="'+Cptgene+'" AND G_NATUREGENE="BQE"',True) ;
      if Not Q.EOF then
         begin
         CompteBanque.Text:=CptGene ;
         HCompte.Caption:=Q.FindField('G_LIBELLE').AsString ;
         end ;
      Ferme(Q) ;
      end ;
  end

  else begin
    //YMO 28/11/2005 Pour les pièces en devise, on va chercher la devise de la banque
    //(compte de contrepartie du journal).
    if tt.Count=0 then exit ;
    O:=TOBM(tt[0]) ;
    Journal:=O.GetMvt('E_JOURNAL') ;
    {JP 20/12/05 : FQ 10457 : si on n'a pas choisi de compte d'effets et que les écritures n'ont pas leur banque
                   prévisionnelle, on reprend le compte de contrepartie du journal. S'il s'agit d'un compte
                   d'effets, Q sera vide et on n'affectera donc pas CompteBanque et HCompte}
    Q := OpenSQL('SELECT BQ_DEVISE, BQ_GENERAL, BQ_LIBELLE FROM BANQUECP LEFT JOIN JOURNAL ON BQ_GENERAL = J_CONTREPARTIE ' +
                 'WHERE J_JOURNAL = "' + journal + '" AND J_NATUREJAL = "BQE"'
                 +' AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"', True); // 24/10/2006 YMO Multisociétés

    if not Q.EOF then begin
      BqDevise := Q.FindField('BQ_DEVISE').AsString;
      {JP 20/12/05 : FQ 10457 : S'il y a une banque prévisionnelle, elle a été affectée dans le FormShow,
                    juste avant l'appel de cette procedure}
      if (CompteBanque.Text = '') and (SorteLettre = tslBOR) then begin
        CompteBanque.Text := Q.FindField('BQ_GENERAL').AsString;
        HCompte.Caption   := Q.FindField('BQ_LIBELLE').AsString;
        CompteBanque.Enabled := True;
        CompteBanque.Color   := clWindow;
      end;
    end;

    Ferme(Q) ;

    if Charge then begin
      Modele:=ModeleTraite ;
      E_MODELE.Value:=Modele ;
    end ;
  end ;
  CompteToBanque(Charge) ;
end ;

procedure TFDocRegl.GenererTablesTemp ;
// Génère les enregistrements dans les tables temporaires DOCREGLE et DOCFACT
var i,ie,j,NumPG,kk : integer ;
    O : TOBM ;
    m : RMVT ;
    Montant, MontantDev, MontantEsc, TauxEsc  : Double ;
    Q3,QTemp : TQuery ;
    s,Origines,SQL3,SQLTemp,Rib,Lib,s1,ModePaie,RefTire,BanquePrevi,RefI : String ;
    DateEche,DateComptable : TDateTime ;
    sTrace : String ;
    YY : RMVT ;
    Okok : Boolean ;
    TOBE : TOB ;
    Jal,Exo  : String ;
    SQL  : String ;
    NumP,NumL,NumE : integer ;
    //SG6 06/01/05 Gestion sous forme de Tob la génération de la table DocFact
    vTobDocF : TOB;
    vTobDocFFille : TOB;
    lNumEncaDeca : String ; // SBO 01/04/2005 : FQ 15432
begin
// Suppression préalable des chèques non imprimés
ExecuteSQL('DELETE FROM DOCFACT WHERE DF_USER="'+V_PGI.User+'" AND EXISTS(SELECT * FROM DOCREGLE WHERE DR_ORIGINE=DOCFACT.DF_ORIGINE AND DR_EDITE="-")') ;
ExecuteSQL('DELETE FROM DOCREGLE WHERE DR_USER="'+V_PGI.User+'" AND DR_EDITE="-"') ;

// Création des enregistrements
for i:=1 to G.RowCount-2 do //VT 03/01/2002 // VL 06/01/03
   begin

   Origines := G.Cells[colOrigines,i] ;

   // On cherche les infos de la première échéance
   s             := Origines ;
   s1            := ReadTokenSt(s) ;
   j             := StrToInt(s1) ;
   O             := TOBM(tt[j-1]) ;
   Rib           := O.GetMvt('E_RIB') ;
   DateEche      := O.GetMvt('E_DATEECHEANCE') ;
   DateComptable := O.GetMvt('E_DATECOMPTABLE') ;
   NumPG         := O.GetMvt('E_NUMEROPIECE') ;
   ModePaie      := O.GetMvt('E_MODEPAIE') ;
   RefTire       := O.GetMvt('E_REFEXTERNE') ;
   Lib           := O.GetMvt('E_LIBELLE') ;
   RefI          := O.GetMvt('E_REFINTERNE') ;
   lNumEncaDeca  := O.GetMvt('E_NUMENCADECA') ; // SBO 01/04/2005 : FQ 15432

   if SuiviMP then
   begin
     {JP 21/02/05 : Pour les BOR, on utilise la valeur de CompteBanque qui est soit saisie soit
                    renseignée automatiquement si E_BANQUEPREVI est renseigné }
     if SorteLettre = tslBOR then BanquePrevi := CompteBanque.Text
                             else BanquePrevi := O.GetMvt('E_BANQUEPREVI') ;
   end else
    BanquePrevi   := O.GetMvt('E_BANQUEPREVI') ;

   if Not SuiviMP then
     begin
     if Lib='' then
       if ((s='') or (SorteLettre=tslBOR)) then
         if ((O.LC<>Nil) and (O.LC.Count>0)) then
         begin
         SQLTemp := 'SELECT * FROM ECRITURE WHERE ' ;
         m:=DecodeLC(O.lc[0]) ; // record de type rmvt
         SQLTemp := SQLTemp + WhereEcriture(tsGene,m,True) ; // après Where
         QTemp := OpenSQL(SQLTemp,True) ;
         if not QTemp.EOF then
            begin
            Lib:=QTemp.FindField('E_LIBELLE').AsString
            end ;
         Ferme(QTemp) ;
         end ;
      end ;
   if Lib='' then Lib:=G.Cells[colLibelle,i] ;

   //RR 16032003
   SQL := 'INSERT INTO DOCREGLE (' ;
   SQL := SQL+ 'DR_USER,' ;
   SQL := SQL + 'DR_AUXILIAIRE,' ;
   SQL := SQL + 'DR_DATEECHEANCE,' ;
   SQL := SQL + 'DR_DATECOMPTABLE,';
   SQL := SQL + 'DR_MODEPAIE,';
   SQL := SQL + 'DR_MONTANT,';
   SQL := SQL + 'DR_MONTANTDEV,';
   SQL := SQL + 'DR_ORIGINE,';
   SQL := SQL + 'DR_LIBELLE,';
   SQL := SQL + 'DR_NUMCHEQUE,';
   SQL := SQL + 'DR_EDITE,';
   SQL := SQL + 'DR_RIB,';
   SQL := SQL + 'DR_RIBEMETTEUR,';
   SQL := SQL + 'DR_CPTBANQUE,';
   SQL := SQL + 'DR_GENERAL,';
   SQL := SQL + 'DR_REFTIRE,';
   SQL := SQL + 'DR_NUMEROPIECE,';
   SQL := SQL + 'DR_BANQUEPREVI,';
   SQL := SQL + 'DR_NUMENCADECA,';                    // SBO 01/04/2005 : FQ 15432
   SQL := SQL + 'DR_DEVISE';                          // YMO 22/02/2006 : FQ 17532
   SQL := SQL + ') VALUES ("' ;
   SQL := SQL + V_PGI.User + '","' ;
   SQL := SQL + G.Cells[colTiers,i] + '","' ;
   SQL := SQL + usdatetime(DateEche)  + '","' ;
   SQL := SQL + usdatetime(DateComptable) + '","' ;
   SQL := SQL + ModePaie  + '",' ;
   SQL := SQL + StrFPoint(valeur(G.Cells[colMontantPIVOT,i])) + ',' ;
   SQL := SQL + StrFPoint(valeur(G.Cells[colMontantDEV,i]))+ ',' ;
   SQL := SQL + inttostr(LastOrigine+i) + ',"' ;
   SQL := SQL + CheckdblQuote(Lib)  + '","' ;
   // SBO 21/01/2005 FQ 15252 Reprise du numéro de traite en mode réédition
   if FBoReEdit
     then SQL := SQL + Copy( O.GetMvt('E_NUMTRAITECHQ') ,1,17)  + '","'
     else SQL := SQL + Copy(RefI,1,17)  + '","' ;
   // Fin SBO 21/01/2005 FQ 15252
   SQL := SQL + '-'  + '","' ;
   SQL := SQL + Rib  + '","' ;
   SQL := SQL + GeneRIB  + '","' ;
   SQL := SQL + CompteBanque.Text  + '","' ;
   SQL := SQL + G.Cells[colCollectif,i]  + '","' ;
   SQL := SQL + RefTire  + '",' ;
   SQL := SQL + inttostr(NumPG) + ',"' ;
   SQL := SQL + BanquePrevi + '","' ;
   SQL := SQL + lNumEncaDeca  + '","' ;           // SBO 01/04/2005 : FQ 15432
   SQL := SQL + Genedevise;                       // YMO 22/02/2006 : FQ 17532
   SQL := SQL + '")';

   ExecuteSQL(SQL);

   ie:=0 ;
   while Length(Origines)>0 do
      begin
      s:=ReadTokenSt(Origines) ; j:=StrToInt(s) ;
      O:=TOBM(tt[j-1]) ;
      SQL3 := 'SELECT * FROM ECRITURE WHERE (' ;
      if Not SuiviMP then
         BEGIN
         Okok:=False ;
         for kk:=0 to O.LC.Count-1 do
             BEGIN
             m:=DecodeLC(O.LC[kk]) ;
             SQL3:=SQL3+WhereEcriture(tsGene,m,True)+') OR (' ;
             Okok:=True ;
             END ;
         System.Delete(SQL3,Length(SQL3)-4,5) ;
         END else
         BEGIN
         sTrace:=O.GetMvt('E_TRACE') ; Okok:=False ;
         if sTrace<>'' then
            BEGIN
            YY:=DecodeLC(sTrace) ; Okok:=True ;
            SQL3:=SQL3+WhereEcriture(tsGene,YY,True)+')' ;
            END ;
         END ;
      if Okok then
         BEGIN
         Q3:=OpenSQL(SQL3,True) ;
         //SG6 06/01/05 Insertion dans table DocFact a l'aide d'une TOB
         //             Ajout de nouveaux champs
         vTobDocF:=TOB.Create('$DOCFACT',nil,-1);

         While Not Q3.EOF do
            begin
            vTobDocFFille := TOB.Create('DOCFACT',vTobDocF,-1);

            Inc(ie) ;

            Montant:=Q3.FindField('E_CREDIT').AsFloat-Q3.FindField('E_DEBIT').AsFloat;

            Montantdev:=Q3.FindField('E_CREDITDEV').AsFloat-Q3.FindField('E_DEBITDEV').AsFloat;
            MontantEsc  := 0 ;
            TauxEsc     := 0 ;
   	    if (TOBParamEsc <> nil) and ( TOBParamEsc.Detail.Count > 0 ) then
            begin
              Jal	:= Q3.FindField('E_JOURNAL').asString ;
              Exo	:= Q3.FindField('E_EXERCICE').AsString ;
              NumP	:= Q3.FindField('E_NUMEROPIECE').AsInteger ;
              NumL	:= Q3.FindField('E_NUMLIGNE').AsInteger ;
              NumE	:= Q3.FindField('E_NUMECHE').AsInteger ;
              TOBE	:= TOBParamEsc.FindFirst(['E_JOURNAL','E_EXERCICE','E_NUMEROPIECE','E_NUMLIGNE','E_NUMECHE'],[Jal,Exo,NumP,NumL,NumE],False) ;
	      if (TOBE <> nil) and (TOBE.FieldExists('SANSESCOMPTE')) and (TOBE.GetValue('SANSESCOMPTE')<>'X') then
              begin
                MontantEsc := TOBE.GetValue('MONTANTESC') ;
  		TauxEsc := TOBE.GetValue('TAUXESC') ;
              end;
            end;

            vTobDocFFille.PutValue('DF_USER',V_PGI.User);
            vTobDocFFille.PutValue('DF_DATEECHEANCE',Q3.FindField('E_DATEECHEANCE').AsDateTime);
            vTobDocFFille.PutValue('DF_MONTANT',Montant*GeneCoef);
            vTobDocFFille.PutValue('DF_MONTANTDEV',Montantdev*GeneCoef);
            vTobDocFFille.PutValue('DF_LIBELLE',Q3.FindField('E_LIBELLE').AsString);
            vTobDocFFille.PutValue('DF_REFERENCE',Q3.FindField('E_REFINTERNE').AsString);
            vTobDocFFille.PutValue('DF_REFEXTERNE',Q3.FindField('E_REFEXTERNE').AsString);
            vTobDocFFille.PutValue('DF_ORIGINE',LastOrigine+i);
            vTobDocFFille.PutValue('DF_DATECOMPTABLE',Q3.FindField('E_DATECOMPTABLE').AsDateTime);
            vTobDocFFille.PutValue('DF_COUVERTURE',Q3.FindField('E_COUVERTURE').AsFloat);
            vTobDocFFille.PutValue('DF_COUVERTUREDEV',Q3.FindField('E_COUVERTUREDEV').AsFloat);
            vTobDocFFille.PutValue('DF_INDICE',ie);
            vTobDocFFille.PutValue('DF_DATEREFEXTERNE',Q3.FindField('E_DATEREFEXTERNE').AsDateTime);
            vTobDocFFille.PutValue('DF_NUMEROPIECE',Q3.FindField('E_NUMEROPIECE').AsInteger);
            vTobDocFFille.PutValue('DF_NUMLIGNE',Q3.FindField('E_NUMLIGNE').AsInteger);
            vTobDocFFille.PutValue('DF_CONTREPARTIEAUX',Q3.FindField('E_CONTREPARTIEAUX').AsString);
            vTobDocFFille.PutValue('DF_TAUXESC',TauxEsc);
            vTobDocFFille.PutValue('DF_MONTANTESC',MontantEsc);
            //SG6 06/01/05 Nouveau Champ
            vTobDocFFille.PutValue('DF_JOURNAL',Q3.FindField('E_JOURNAL').AsString);
            vTobDocFFille.PutValue('DF_NATUREPIECE',Q3.FindField('E_NATUREPIECE').AsString);
            vTobDocFFille.PutValue('DF_SOCIETE',Q3.FindField('E_SOCIETE').AsString);
            vTobDocFFille.PutValue('DF_ETABLISSEMENT',Q3.FindField('E_ETABLISSEMENT').AsString);
            vTobDocFFille.PutValue('DF_REFLIBRE',Q3.FindField('E_REFLIBRE').AsString);
            vTobDocFFille.PutValue('DF_AFFAIRE',Q3.FindField('E_AFFAIRE').AsString);
            vTobDocFFille.PutValue('DF_MODEPAIE',Q3.FindField('E_MODEPAIE').AsString);
            vTobDocFFille.PutValue('DF_DATERELANCE',Q3.FindField('E_DATERELANCE').AsDateTime);
            vTobDocFFille.PutValue('DF_NIVEAURELANCE',Q3.FindField('E_NIVEAURELANCE').AsInteger);
            vTobDocFFille.PutValue('DF_TIERSPAYEUR',Q3.FindField('E_TIERSPAYEUR').AsString);
            vTobDocFFille.PutValue('DF_QTE1',Q3.FindField('E_QTE1').AsFloat);
            vTobDocFFille.PutValue('DF_QTE2',Q3.FindField('E_QTE2').AsFloat);
            vTobDocFFille.PutValue('DF_NOMLOT',Q3.FindField('E_NOMLOT').AsString);
            vTobDocFFille.PutValue('DF_LIBRETEXTE0',Q3.FindField('E_LIBRETEXTE0').AsString);
            vTobDocFFille.PutValue('DF_LIBRETEXTE1',Q3.FindField('E_LIBRETEXTE1').AsString);
            vTobDocFFille.PutValue('DF_LIBRETEXTE2',Q3.FindField('E_LIBRETEXTE2').AsString);
            vTobDocFFille.PutValue('DF_LIBRETEXTE3',Q3.FindField('E_LIBRETEXTE3').AsString);
            vTobDocFFille.PutValue('DF_LIBRETEXTE4',Q3.FindField('E_LIBRETEXTE4').AsString);
            vTobDocFFille.PutValue('DF_LIBRETEXTE5',Q3.FindField('E_LIBRETEXTE5').AsString);
            vTobDocFFille.PutValue('DF_LIBRETEXTE6',Q3.FindField('E_LIBRETEXTE6').AsString);
            vTobDocFFille.PutValue('DF_LIBRETEXTE7',Q3.FindField('E_LIBRETEXTE7').AsString);
            vTobDocFFille.PutValue('DF_LIBRETEXTE8',Q3.FindField('E_LIBRETEXTE8').AsString);
            vTobDocFFille.PutValue('DF_LIBRETEXTE9',Q3.FindField('E_LIBRETEXTE9').AsString);
            vTobDocFFille.PutValue('DF_TABLE0',Q3.FindField('E_TABLE0').AsString);
            vTobDocFFille.PutValue('DF_TABLE1',Q3.FindField('E_TABLE1').AsString);
            vTobDocFFille.PutValue('DF_TABLE2',Q3.FindField('E_TABLE2').AsString);
            vTobDocFFille.PutValue('DF_TABLE3',Q3.FindField('E_TABLE3').AsString);
            vTobDocFFille.PutValue('DF_LIBREMONTANT0',Q3.FindField('E_LIBREMONTANT0').AsFloat);
            vTobDocFFille.PutValue('DF_LIBREMONTANT1',Q3.FindField('E_LIBREMONTANT1').AsFloat);
            vTobDocFFille.PutValue('DF_LIBREMONTANT2',Q3.FindField('E_LIBREMONTANT2').AsFloat);
            vTobDocFFille.PutValue('DF_LIBREMONTANT3',Q3.FindField('E_LIBREMONTANT3').AsFloat);
            vTobDocFFille.PutValue('DF_NUMTRAITECHQ',Q3.FindField('E_NUMTRAITECHQ').AsString);
            vTobDocFFille.PutValue('DF_NUMENCADECA', lNumEncaDeca );                          // SBO 01/04/2005 : FQ 15432
            vTobDocFFille.PutValue('DF_REFGESCOM',Q3.FindField('E_REFGESCOM').AsString);
            vTobDocFFille.PutValue('DF_LIBREDATE',Q3.FindField('E_LIBREDATE').AsDateTime);
            vTobDocFFille.PutValue('DF_SAISIMP',Q3.FindField('E_SAISIMP').AsFloat);
            // YMO 22/02/2006 : FQ 17532              
            vTobDocFFille.PutValue('DF_DEVISE',Q3.FindField('E_DEVISE').AsString);
            Q3.Next ;

            end ;

         Ferme(Q3) ;
         //SG6 06/01/05 Insertion ds table DocFact
         vTobDocF.InsertDB(nil);
         FreeAndNil(vTobDocF);
         END ;
      end ;
   end;
end ;

procedure TFDocRegl.RazTemp ;
begin
ExecuteSQL('DELETE FROM DOCREGLE WHERE DR_USER="'+V_PGI.User+'"') ;
ExecuteSQL('DELETE FROM DOCFACT WHERE DF_USER="'+V_PGI.User+'"') ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 21/09/2004
Modifié le ... :   /  /    
Description .. : Remplit la grille des échéances à partir de la liste des 
Suite ........ : échéances "tt"
Mots clefs ... : 
*****************************************************************}
procedure TFDocRegl.AfficheGridEcheances ;
var i : integer ;
    O : TOBM ;
begin
  // Vide la grille
  GE.VidePile(True) ;
  GE.RowCount := tt.Count + 1 ; //VT 03/01/2002

  // Parcours la liste des échances
  for i:=0 to tt.Count-1 do
    begin
    O := TOBM( tt[i] ) ;
    // on vide la flag de traitement de l'écriture
    O.PutMvt( 'E_FLAGECR', '' ) ;
    // remplit la ligne
    AjouteEcheGrid(O, i+1 ) ;
    end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 21/09/2004
Modifié le ... :   /  /    
Description .. : Remplit la ligne "lig" de la grille des échéance avec les 
Suite ........ : données de l'écriture contenu dans le TOBM "O"
Mots clefs ... :
*****************************************************************}
procedure TFDocRegl.AjouteEcheGrid(O : TOBM ; lig : Integer) ;
var m,md : Double ;
begin
  if TIDTIC
    then GE.Cells[0,lig] := O.GetMvt('E_GENERAL')
    else GE.Cells[0,lig] := O.GetMvt('E_AUXILIAIRE') ;
  GE.Cells[1,lig] := IntToStr(O.GetMvt('E_NUMEROPIECE')) ;
  GE.Cells[2,lig] := DateToStr(O.GetMvt('E_DATEECHEANCE')) ;
  m  := ( O.GetMvt('E_DEBIT')    - O.GetMvt('E_CREDIT')    ) * GeneCoef ;
  md := ( O.GetMvt('E_DEBITDEV') - O.GetMvt('E_CREDITDEV') ) * GeneCoef ;
  GE.Cells[3,lig] := StrFMontant( m,  15, V_PGI.OkDecV, '', True ) ;
  GE.Cells[4,lig] := StrFMontant( md, 15, V_PGI.OkDecV, '', True ) ;
  GE.Cells[5,lig] := O.GetMvt('E_GENERAL') ;
end ;

function TFDocRegl.LibelleTiers(Aux : string) : string ;
var QTiers : TQuery ;
    St : String ;
begin
result := '' ;
If TIDTIC Then St:='SELECT G_LIBELLE FROM GENERAUX WHERE G_GENERAL="' + Aux + '"'
          Else St:='SELECT T_LIBELLE FROM TIERS WHERE T_AUXILIAIRE="' + Aux + '"' ;
QTiers := OpenSQL(St,True) ;
if not QTiers.EOF then result := QTiers.Fields[0].AsString ;
Ferme(QTiers) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 09/05/2003
Modifié le ... :   /  /    
Description .. : 
Suite ........ : 09/05/2003 : Virer les accolades dans la requête...
Mots clefs ... : 
*****************************************************************}
function TFDocRegl.RequeteDocRegl(l : integer) : string ;
var sql : string ;
begin

  if bEtatAGL then
    begin
    // Pour les lanceEtat, la requête peut être paramétrée par les utilisateurs
    //  , on ne fixe que le WHERE sur le document à éditer
    sql := 'DR_USER="' + V_PGI.User + '" AND DR_ORIGINE=' + IntToStr( LastOrigine + l ) ;
    end
  else
    begin
    // Pour les lanceDocument, le traitement reste inchangé
    sql := 'SELECT * FROM DOCREGLE' ;
    sql := sql + ' LEFT JOIN DOCFACT ON DF_USER=DR_USER AND DF_ORIGINE=DR_ORIGINE' ;
    if TIDTIC
      then sql := sql + ' LEFT JOIN GENERAUX ON G_GENERAL=DR_AUXILIAIRE'
      else sql := sql + ' LEFT JOIN TIERS ON T_AUXILIAIRE=DR_AUXILIAIRE'
                      + ' LEFT OUTER JOIN CONTACT ON C_AUXILIAIRE=T_AUXILIAIRE AND C_PRINCIPAL="X"' ;
    //YMO 27/02/2006 FQ12928 Jointure ajoutée
    sql := sql + ' LEFT JOIN RESSOURCE ON ARS_UTILASSOCIE=DR_USER' ;
    sql := sql + ' LEFT OUTER JOIN BANQUECP ON BQ_GENERAL=DR_CPTBANQUE' ;
    sql := sql + ' AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"' ; // 19/10/2006 YMO Multisociétés
    sql := sql + ' WHERE DR_USER="' + V_PGI.User + '"' ;
    sql := sql + ' AND DR_ORIGINE=' + IntToStr( LastOrigine + l ) ;
    end ;

  Result := sql ;

end ;

function TFDocRegl.NoMoreEche : Boolean ;
var i : integer ;
begin
result:=True ;
for i:=1 to GE.RowCount-1 do
   begin
   if GE.RowHeights[i]<>-1 then begin Result:=False ; exit ; end ;
   end ;
end ;

procedure TFDocRegl.PutDR_NUMCHEQUE(snChq,sRef : String) ;
Var i,NoChq : Integer ;
    STransit : String ;
BEGIN
for i:=1 to G.RowCount-1 do //VT 03/01/2003
   begin
   if (Trim(G.Cells[colTiers,i])='') then continue ;
   noChq:=StrToInt(snChq) ;
   if i=1 then noChq:=noChq else noChq:=noChq+1 ;
   sTransit:=IntToStr(noChq) ; While Length(sTransit)<Length(snChq) do sTransit:='0'+sTransit ;
   snChq:=sTransit ;
   AttributNumCheque(i,sRef,snChq,True) ;
   end ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/01/2003
Modifié le ... : 21/09/2004
Description .. : 28/01/2003 : Ajout du paramètre AvecImpr (boolean à vrai
Suite ........ : par défaut), qui permet de générer l'édition dans les tables
Suite ........ : temporaires DocRegl et DocFact sans l'imprimer (si passer à
Suite ........ : Faux) (utilisé dans le cas d'un traitement d'escompte).
Suite ........ :
Suite ........ : 21/09/2004 : SBO : Remplacement des lancedocument par
Suite ........ : des lanceetat
Mots clefs ... :
*****************************************************************}
procedure TFDocRegl.ImprimeDocRegl(Selection,Test,Apercu : boolean ; AvecImpr : boolean = True) ;
var sql        : string ;
    TTI        : TList ;
    LL         : TStrings ;
    noChq      : integer ;
    Incr       : integer ;
    i          : integer ;
    NbCheques  : integer ;
    NbLignes   : integer ;
    NbPages    : integer ;
    NbPagesTot : integer ;
    sRef       : string ;
    snChq      : string ;
    sNumChq    : string ;
    sTransit   : string ;
    lStSpooler : string ;
    Okok       : Boolean ;
    lBoQRPDF   : Boolean ;
    lBoNoPrintDialog : Boolean ;
    NatEtat    : String ;
begin

  // =========================
  // == Test des paramètres ==
  // =========================
  // Pour les BOR
  if (SorteLettre = tslBor) then
    begin
    // Compte bancaire obligatoire
    if CompteBanque.Text = '' then
      begin
      Msg.Execute(10,Caption,'') ;
      Exit ;
      end ;
    // Vérification de la présence du compte de banque renseigné
    if Not PresenceComplexe( 'GENERAUX', ['G_GENERAL','G_NATUREGENE'],
                                         ['=','='],
                                         [CompteBanque.Text,'BQE'],
                                         ['S','S'] ) then
      begin
      Msg.Execute(10,Caption,'') ;
      Exit ;
      end ;
    end ;

  lBoQRPDF := False ;

  // Test modèle / devise pour les BOR et les Traites
  if (SorteLettre = tslBor) or (SorteLettre = tslTraite) then
    begin
    if E_MODELE.Value='' then
      begin
      Msg.Execute(9,Caption,'') ;
      Exit ;
      end ;
    if GeneDevise<>BQDevise then
      begin
      Msg.Execute(20,caption,'') ;
      Exit ;
      end ;
    end ;

  // Pour l'édition unitaire, on test la validité de la ligne sélectionnée
  if Selection and (G.Cells[colTiers,G.Row]='') then exit ;


  // ====================================================
  // == Détermination du numéro et de la référence doc ==
  // ====================================================
  sRef  := '' ;
  snChq := '0' ;
  // Pour les chèque, saisie du numéro
  if SorteLettre=tslCheque then
    begin
    if Not SaisieNumCheque(sRef,snChq,False) then exit ;
    end
  // Pour les BOR, récupération du dernier numéro généré
  else if SorteLettre=tslBOR then
    begin
    sRef  := GetParamSoc('SO_CPSOUCHETRAFOU') ;
    snChq := IntToStr(GetParamSoc('SO_CPNUMTRAFOU')) ;
    end
  // Pour les BOR, récupération du dernier numéro généré
  else if SorteLettre=tslTraite then
    begin
    sRef  := GetParamSoc('SO_CPSOUCHETRACLI') ;
    snChq := IntToStr(GetParamSoc('SO_CPNUMTRACLI')) ;
    end ;

  SourisSablier ;

  // ========================================
  // == Remplissage des tables temporaires ==
  // ========================================
  if Not (Transactions(GenererTablesTemp,3)=oeOk) then
    begin
    SourisNormale ;
    exit ;
    end ;
  TTI := Tlist.Create ;

  // =====================================================
  // == Constitution de la liste des documents à éditer ==
  // =====================================================
  // Tous les documents
  if ((Selection) or (V_PGI.SAV)) then
    begin
    for i:=1 to G.RowCount-2 do //VT 03/01/2002
      begin
      sql := RequeteDocRegl(i) ;
      LL  := TStringList.Create ;
      LL.Add(sql) ;
      TTI.Add(LL) ;
      end ;
     end
  else // Uniquement la ligne en cours
    begin
    AttributNumCheque(G.Row,sRef,snChq,True) ;
    sql := RequeteDocRegl(G.Row) ;
    LL  := TStringList.Create ;
    LL.Add(sql) ;
    TTI.Add(LL) ;
    end ;

  NbPagesTot := 0 ;

  // Détermination de la nature des états : Diffère selon utilisation document OU générateur d'état
  // Pour la version avec générateur d'état
  Case SorteLettre Of
    tslCheque : if bEtatAGL
                  then NatEtat:='CLC'
                  else NatEtat:='LCH' ;
    tslPre    : if bEtatAGL
                  then NatEtat:='CLP'
                  else NatEtat:='LPR' ;
    tslVir    : if bEtatAGL
                  then NatEtat:='CLV'
                  else NatEtat:='LVI' ;
    Else if bEtatAGL
           then NatEtat:='CLT'
           else NatEtat:='LTR' ;
  end ;

  // Update des références pour les chèques / traites / BOR
  if ( SorteLettre in [tslCheque] )
      // FQ 14857 SBO 05/11/2004 pas de renumérotation en mode réédition pour les traites / BOR
      or ( ( SorteLettre in [tslTraite,tslBOR] ) and not FBoReEdit ) then
    PutDR_NUMCHEQUE(snChq,sRef) ;

  // =============
  // == EDITION ==
  // =============
  // Les sortie sur imprimante est paramètrable (cf. entête)
  if AvecImpr then
    begin

    // -----------------------------------
    // -- Edition via générateur d'état --
    // -----------------------------------
    if bEtatAGL then
      begin

      // Sauvegarde du paramétrage d'ouverture des options d'imprimante
      lBoNoPrintDialog := V_PGI.NoPrintDialog ;
      V_PGI.NoPrintDialog := not OkPrintDialog ;

      // Gestion du spooler
      if MSED.Spooler then
        begin
        // Initialisation
        {JP 04/06/07 : FQ 19507 : j'ai l'impression que RacineSpooler est toujours vide}
        if Trim(MSED.RacineSpooler) = '' then
          Case SorteLettre Of
            tslCheque : MSED.RacineSpooler := 'CHQ';
            tslPre    : MSED.RacineSpooler := 'PRE';
            tslVir    : MSED.RacineSpooler := 'VIR';
            tslTraite : MSED.RacineSpooler := 'TRA';
            tslBOR    : MSED.RacineSpooler := 'BOR';
          else
            MSED.RacineSpooler := 'CCMP';
          end;


        lStSpooler := Trim(MSED.RepSpooler) + '\' + Trim(MSED.RacineSpooler) + '.PDF' ;
        lBoQRPDF            := V_PGI.QRPDF ;
        V_PGI.QRPDF         := True;
        V_PGI.QRPDFQueue    := lStSpooler;
        V_PGI.QRPDFMerge    := lStSpooler;
        // En édition spooler, pas d'ouverture des options d'imprimantes
        V_PGI.NoPrintDialog := True;
        // Début du spooler
        StartPDFBatch( lStSpooler );
        // MSED.XFichierSpooler,
        // MSED.SoucheSpooler
        end ;

      // Boucle sur les documents à éditer
      for i:=0 to TTI.Count-1 do
        begin
        NbPages := LanceEtat( 'E', NatEtat, Modele,           // Type, Nature , Modèle d'état
                              Apercu or (MSED.Spooler),       // Aperçu
                              False,
                              False,
                              Nil,
                              TStringList(TTI[i]).Text,
                              '',
                              {JP 20/07/05 : FQ 15544 : ce paramétre ne concerne pas la boite de dialogue, mais le Duplicata
                               OkPrintDialog and (i = 0)       // PrintDialog la première fois seulement}
                              False
                              ) ;

        {JP 02/08/05 : FQ 16131 : Comme on lance plusieurs fois le LanceEtat, mais qu'au-delà de première fois on
                       n'affiche plus la boîte de dialoque(cf ci-dessous), je récupère le nombre de copies demandées
                       sur la première impression et je l'affecte au nombre de copies par défaut}
        if i = 0 then
          V_PGI.DefaultDocCopies := V_PGI.NbDocCopies;
            // Après la 2ème édition, plus d'ouverture des options d'imprimantes
        V_PGI.NoPrintDialog := True ;

        TStringList(TTI[i]).Text := IntToStr( NbPages ) ;
        NbPagesTot := NbPagesTot + NbPages ;

        end ;

      // Gestion du spooler
      if MSED.Spooler then
        begin
        // Fin du spooler
        CancelPDFBatch;
        {$IFNDEF EAGLCLIENT}
        lStSpooler := Trim(MSED.RepSpooler) + '\' + Trim(MSED.RacineSpooler) + '.PDF' ;
        {JP 04/06/07 : FQ 19507 : en 2/3, après le PreviewPDFFile, le fichier est supprimé : on en fait une sauvegarde}
        CopierFichier(lStSpooler, MSED.RepSpooler, Trim(MSED.RacineSpooler), '.PDF');
        PreviewPDFFile('', lStSpooler );
        {$ENDIF}
        V_PGI.QRPDF       := lBoQRPDF;
        V_PGI.QRPDFQueue  := '';
        V_PGI.QRPDFMerge  := '';
        end ;

      // Restauration du paramétrage d'ouverture des options d'imprimante
      V_PGI.NoPrintDialog := lBoNoPrintDialog ;
      end
    else
      // ----------------------------------------
      // -- Edition via générateur de document --
      // ----------------------------------------
      begin
        if Test then
        {$IFDEF EAGLCLIENT}
        {$ELSE}
          TestDocument('L',NatEtat,Modele,TTI,nil,Apercu)
        {$ENDIF}
        else begin
          try
            {$IFNDEF EAGLCLIENT}
            if MSED.Spooler then
              NbPagesTot:=LanceDocumentSpool('L',NatEtat,Modele,TTI,nil,Apercu,OkPrintDialog,FALSE, MSED.RepSpooler,MSED.RacineSpooler,MSED.XFichierSpooler,MSED.SoucheSpooler)
            else
            {$ENDIF}
              NbPagesTot:=LanceDocument('L',NatEtat,Modele,TTI,nil,Apercu,OkPrintDialog) ;
          except
            on E : Exception do begin
              {JP 03/10/07 : FQ 21504 qui fait référence à la FQ 14882 : plutôt que de laisser trainer la violation
                             d'accès, il me semble mieux de mettre un petit message explicatif ...}
              if TIDTIC then PGIBox(TraduireMemoire('Le document n''a pu être créé : veuillez vérifier que vous avez') + #13 +
                                    TraduireMemoire('bien sélectionner le document pour les TIC et TID.') + #13#13 +
                                    TraduireMemoire('Rappel : Le document ne doit pas contenir de champs ') + #13 +
                                    TraduireMemoire('de la table TIERS (Préfixe "T", ex : T_LIBELLE, T_EAN ...).'), Caption)
                        else PGIBox(TraduireMemoire('Le document n''a pu être créé : veuillez vérifier votre état.'), Caption);
            end;
          end;
        end ;
      end


    end
  // Simule l'impression
  else
    begin
    ChqPrinted := True ;
    NbPagesTot    := 1 ;
    end ;

  // ==============================
  // == TRAITEMENT POST EDITIONS ==
  // ==============================
  if (not Test) and (NbPagesTot>0) then
    if ((Not Apercu) or (V_PGI.SAV)) then
      begin
      Okok:=True ;

      // ---------------------------------------------------------------------
      // -- Bor / Traite / Chèque : Gestion de l'incrémentation des numéros --
      // ---------------------------------------------------------------------
      if ( SorteLettre in [tslCheque] )
          // FQ 14857 SBO 05/11/2004 pas de renumérotation en mode réédition pour les traites / BOR
          or ( ( SorteLettre in [tslTraite,tslBOR] ) and not FBoReEdit ) then
        begin
        
        // ==> MAJ des numéros et références des tables temporaires
        for i:=1 to G.RowCount-1 do
          begin
          // Test tiers
          if (Trim(G.Cells[colTiers,i])='') then continue ;
          // Détermination du numéro affecté
          // FQ14893 SBO 03/11/2004 Remise en place du gestion de saut de chèque en 2-tiers.
          // Ne fonctionne pas en eAGL car LanceDocument ne modifie la TList des requêtes.
          {$IFDEF EAGLCLIENT}
          Incr := 1 ;
          {$ELSE}
          if (SorteLettre in [tslCheque]) and (AvecImpr)
            then Incr := StrToInt( TStrings( TTI[i-1] )[0] )
            else Incr := 1 ;
          {$ENDIF EAGLCLIENT}
          noChq := StrToInt( snChq ) ;
          if i = 1
            then noChq := noChq + Incr - 1
            else noChq := noChq + Incr ;
          // Complétion du numéro
          sTransit := IntToStr(noChq) ;
          While Length(sTransit)<Length(snChq) do
            sTransit:='0'+sTransit ;
          // Affectation
          snChq:=sTransit ;
          // Maj des tables temporaires
          AttributNumCheque(i,sRef,snChq,True) ;
          end ; // fin du For

        // ==>  Validation des numéros de chèques / Maj ParamSoc
        if SorteLettre = tslCheque
          then Okok := ValideImpDR( G, LaColMontant )
        else if SorteLettre=tslBOR
          then SetParamSoc('SO_CPNUMTRAFOU',ValeurI(snChq)+1)
        else if SorteLettre=tslTraite
          then SetParamSoc('SO_CPNUMTRACLI',ValeurI(snChq)+1) ;

        end ;

      // -----------------------------------------------------------------------
      // -- Edition OK, MAJ des numéros et références chèques / traites / bor --
      // -----------------------------------------------------------------------
      if Okok then
        begin
        NbCheques := 0 ;
        NbLignes  := G.RowCount-1 ;

        // ==> Boucle sur les documents pour udpate table temporaire
        for i:=1 to G.RowCount-1 do
          begin

          if (Trim(G.Cells[colTiers,i])='') then continue ;

          // Bor / Traite / Chèque -> MAJ indicateur édition + Référence chèque
          if SorteLettre in [tslCheque,tslTraite,tslBOR] then
            begin
            if (G.Cells[colNumCheque,i]<>'') then
              begin
              sNumChq:=G.Cells[colNumCheque,i] ;
              ExecuteSQL('UPDATE DOCREGLE SET DR_EDITE="X", DR_NUMCHEQUE="'+sNumChq+'" WHERE DR_USER="'+V_PGI.User+'" AND DR_ORIGINE='+IntToStr(LastOrigine+i)) ;

              if SuiviMP then
              begin
                {Mise à jour de la banque prévisionnelle}
                if SorteLettre = tslBOR then
                  SetNumChqDansEche('', i, True);
              end;

              Inc(NbCheques) ;
              end ;
            end
          // Autres Docs -> MAJ indicateur édition seulement
          else
            ExecuteSQL('UPDATE DOCREGLE SET DR_EDITE="X" WHERE DR_USER="'+V_PGI.User+'" AND DR_ORIGINE='+IntToStr(LastOrigine+i)) ;

          end ;

        // ==> MAJ indicateur d'impression d'au moins une lettre-chèque
        if ( NbCheques > 0 ) or ( SorteLettre <> tslCheque ) then
          ChqPrinted:=True ;

        // ==> Boucle pour remonter les chèques annulés dans les échéances d'origine
        if SorteLettre = tslCheque then
          for i := G.RowCount - 1 downto 1 do
            begin
            if (Trim(G.Cells[colTiers,i])='') then continue ;
            if (G.Cells[colNumCheque,i]='') then
              ChequeToEche(i) ;
            end ;

        // ==> MAJ grille + indicateur pour continuer édition sur un autre modèle
        G.VidePile(True) ;
        G.RowCount  := 2 ; // BPY le 10/09/2004 : fiche n° 12446 => permettre l'edition en plusieur fois !
        LastOrigine := LastOrigine + NbLignes ;

        // Si plus rien à éditer -> Validation automatique
        if (ChqPrinted) and (NoMoreEche) and (AvecImpr) then // Uniquement si reellement imprimer
          bValiderClick(nil) ;

      end
    else
      // -------------------------------------------------
      // -- Boucle d'annulation des n° de chèque (grid) --
      // -------------------------------------------------
      for i:=1 to G.RowCount-1 do
        begin
        if (Trim(G.Cells[colTiers,i])='') then continue ;
        G.Cells[colNumCheque,i]:='' ;
        end ;

   end ;

  // FQ 10700
  BValider.Enabled := ChqPrinted;
  SourisNormale ;
  VideListe(TTI) ;
  TTI.Free ;

end ;

procedure TFDocRegl.AttributNumCheque(Ligne : Integer ; sRef,sNumCheque : String ; Affiche : boolean) ;
begin

  if ( SorteLettre in [tslVir, tslPre] )
      // FQ 14857 SBO 05/11/2004 pas de renumérotation en mode réédition pour les traites / BOR
      or ( ( SorteLettre in [tslTraite,tslBOR] ) and FBoReEdit ) then Exit;

  //VT 03/01/2002 (pour CWAS)
  ExecuteSQL('UPDATE DOCREGLE SET DR_NUMCHEQUE="' + FormatCheque(sRef,sNumCheque)
                     + '" WHERE DR_USER="' + V_PGI.User + '" AND DR_ORIGINE=' + IntToStr(LastOrigine+Ligne)) ;
  if Affiche then
    G.Cells[colNumCheque,Ligne] := FormatCheque(sRef,sNumCheque) ;

    SetNumChqDansEche( sNumCheque , Ligne ) ; // add SBO maj n°chq depuis CCMP // Fiche 12017

end ;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 21/09/2004
Modifié le ... :   /  /
Description .. : Transfert l'échéance de la ligne iEche de la grille des
Suite ........ : échéances dans la grille des documents à la ligne iCheq
Mots clefs ... :
*****************************************************************}
procedure TFDocRegl.EcheToCheque( iEche, iCheq : Integer ; VoirErreurs, ControleNeg : Boolean) ;
var MontantChequePIVOT : Double ;
    MontantEchePIVOT   : Double ;
    MontantChequeDEV   : Double ;
    MontantEcheDEV     : Double ;
    O                  : TOBM ;
    Origines           : string ;
    Pb                 : Boolean ;
begin

  // Contrôles de validité des indices de grille
  if ( iCheq <= 0 ) or ( iCheq > G.RowCount ) then exit ;
  if ( iEche <= 0 ) or ( iEche > GE.RowCount) then exit ;

  O:=TOBM(tt[iEche-1]) ;
  if O.GetMvt('E_FLAGECR')='X' then exit ;

  // Test du tiers de destination correspond bien au tiers de l'échéance
  if TIDTIC
    then Pb := O.GetMvt('E_GENERAL')    <> G.Cells[ colTiers, iCheq ]
    {JP 28/09/01 : FQ 16431 : Un même auxiliaire peut avoir deux généraux différents :
                   il faut donc rajouter le test sur le général}
    else Pb := (O.GetMvt('E_GENERAL')    <> G.Cells[ colCollectif, iCheq ]) or
               (O.GetMvt('E_AUXILIAIRE') <> G.Cells[ colTiers, iCheq ]);

  if (G.Cells[colTiers,iCheq]<>'') and Pb then
  begin
    if VoirErreurs then begin
      {JP 28/09/01 : FQ 16431 : Un même auxiliaire peut avoir deux généraux différents :
                     il faut donc rajouter le test sur le général}
      if O.GetMvt('E_AUXILIAIRE') <> G.Cells[ colTiers, iCheq ] then
        Msg.Execute(3,caption,'')
      else
        HShowMessage('0;' + Caption + ';Les généraux doivent être identiques.;W;O;O;O;', '', '');
    end;
    Exit ;
  end ;

  // Récupération des montants
  MontantEchePIVOT   := ( O.GetMvt('E_DEBIT')    - O.GetMvt('E_CREDIT')    ) * GeneCoef ;
  MontantEcheDEV     := ( O.GetMvt('E_DEBITDEV') - O.GetMvt('E_CREDITDEV') ) * GeneCoef ;
  MontantChequePIVOT := Valeur( G.Cells[ colMontantPIVOT, iCheq] ) ;
  MontantChequeDEV   := Valeur( G.Cells[ colMontantDEV,   iCheq] ) ;

  // Controle si les montants négatifs sont autorisé
  if ControleNEG and ( MontantChequePIVOT + MontantEchePIVOT <= 0 )
                 and not ( ( smp in [smpEncTraEdtNC] ) and ( SorteLettre in [tslTraite] ) ) then
     begin
     if VoirErreurs then Msg.Execute(5,caption,'') ;
     Exit ;
     end ;

  // Placement des grilles sur les lignes à traiter
  GE.Row := iEche ;
  G.Row  := iCheq ;

  // Affectation d'une nouveau document si besoin
  if G.Cells[colTiers,iCheq]='' then
    begin
    // Nouvelle ligne
    G.RowCount := G.RowCount + 1 ;
    MontantChequePIVOT := 0 ;
    MontantChequeDEV   := 0 ;
    // Code + Libelle compte
    if TIDTIC then
      begin
      G.Cells[colTiers,iCheq]   := O.GetMvt('E_GENERAL') ;
      G.Cells[colLibelle,iCheq] := LibelleTiers(O.GetMvt('E_GENERAL')) ;
      end
    else
      begin
      G.Cells[colTiers,iCheq]:=O.GetMvt('E_AUXILIAIRE') ;
      G.Cells[colLibelle,iCheq]:=LibelleTiers(O.GetMvt('E_AUXILIAIRE')) ;
      end ;
    // banque
    G.Cells[colPrevi,iCheq]:=O.GetMvt('E_BANQUEPREVI') ;
    // date d'échéance pour CCMP (???)
    if SuiviMP then
      G.Cells[ColDateEche,iCheq]:=DateToStr(O.GetMvt('E_DATEECHEANCE')) ;
    // Compte collectif
    G.Cells[colCollectif,iCheq]:=O.GetMvt('E_GENERAL') ; // GP le 02/02/2002
    end ;

  // Calcul des nouveaux montant du document suite à l'ajout de l'échéance
  MontantChequePIVOT := MontantChequePIVOT + MontantEchePIVOT ;
  G.Cells[ colMontantPIVOT, iCheq ] := StrFMontant( MontantChequePIVOT, 15, V_PGI.OkDecV, '' , True ) ;
  MontantChequeDEV   := MontantChequeDEV + MontantEcheDEV ;
  G.Cells[ colMontantDEV, iCheq ]   := StrFMontant( MontantChequeDEV, 15, DEV.Decimale, '', True ) ;

  // Maj à jour de la liste des lignes d'échéances d'origines sous la forme 'a;b;c;...')
  Origines := G.Cells[ colOrigines, iCheq ] ;
  if Origines <> ''
    then Origines := Origines + ';' + IntToStr(iEche)
    else Origines := IntToStr(iEche) ;
  G.Cells[ colOrigines, iCheq ] := Origines ;

  // Indicateur traitement
  O.PutMvt('E_FLAGECR','X') ;

  // on cache la ligne d'échéance traitée
  GE.RowHeights[ iEche ] := -1 ;

end ;

procedure TFDocRegl.ChequeToEche(iCheq : Integer) ;
var i          : Integer ;
    Origines   : string ;
    s          : string ;
    O          : TOBM ;
begin
// Contrôles
if (iCheq<=0) or (iCheq>G.RowCount) then exit ; //VT 03/01/2002
Origines:=G.Cells[colOrigines,iCheq] ;
while Length(Origines)>0 do
   begin
   s:=ReadTokenSt(Origines) ;
   i:=StrToInt(s) ;
   O:=TOBM(tt[i-1]) ;
   O.PutMvt('E_FLAGECR','') ;
   GE.RowHeights[i]:=GE.DefaultRowHeight ;
   end ;
G.DeleteRow(iCheq) ;
G.Row:=1 ;
end ;

procedure TFDocRegl.InitCheques ;
Var i : integer ;
begin
ForcerGroupe:=True ;
if SuiviMP then
   BEGIN
   if GroupeEncadeca='DET' then BCopyDetailClick(Nil) else
    if ((GroupeEncadeca='ECH') or (GroupeEncadeca='ECT')) then BGroupTiersEcheClick(Nil) else
     BGroupTiersClick(Nil) ;
   END else
   BEGIN
   bGroupTiersClick(nil) ; // Générer un chèque par tiers automatiquement
   END ;
ForcerGroupe:=False ;
for i:=G.RowCount-1 downto 1 do //VT 03/01/2002
    BEGIN
    if ((G.Cells[ColTiers,i]<>'') and (Valeur(G.Cells[colMontantPIVOT,i])<=0)) then ChequeToEche(i) ;
    END ;
end ;

procedure TFDocRegl.FormShow(Sender: TObject);
var O : TOBM ;
begin
  // En attendant d'être sur de la présence du PARAMSOC dans la SocRef, on fait le test
  if ExisteSQL('SELECT SOC_NOM FROM PARAMSOC WHERE SOC_NOM = "SO_CPDOCAVECETAT"')
    then bEtatAGL := GetParamSoc('SO_CPDOCAVECETAT', True)
    else bEtatAGL := False ;

   //SG6 12/03/2004
  {$IFDEF EAGLCLIENT}
     //on cache le bouton d'edition du modèle de document
     if not(bEtatAGL) then bModele.Visible:=false;
  {$ENDIF}

  if SuiviMP then
    bGroupTiersEche.Visible:=True ;
  If MSED.Spooler Then
    BApercu.Enabled:=FALSE ;

  ModeleOk := InitE_MODELE ;
  If ModeleMultiSession<>'' Then
    E_Modele.Value := ModeleMultiSession ;

  SourisSablier ;
  RazTemp ;

  // Initialisation de la grille des documents
  LaColMontant               := colMontantPivot ;
  G.ColWidths[ColOrigines]   := FLargeur; {JP 05/06/07 : FQ 18664}
  G.ColWidths[colPrevi]      := FLargeur; {JP 05/06/07 : FQ 18664}
  G.ColWidths[ColMontantDEV] := FLargeur; {JP 05/06/07 : FQ 18664}
  G.ColWidths[colDateEche]   := FLargeur; {JP 05/06/07 : FQ 18664}
  G.ColLengths[ColOrigines]   := -1 ;
  G.ColLengths[colPrevi]      := -1 ;
  G.ColLengths[ColMontantDEV] := -1 ;
  G.ColLengths[colDateEche]   := -1 ;

  {JP 28/09/01 : FQ 16431 : Un même auxiliaire peut avoir deux généraux différents :
                 il faut donc rajouter le test sur le général si l'on est pas en TIC
                 ou TID => dans ce cas il est mieux d'afficher le général}
  if TIDTIC then begin
    G.ColWidths[colCollectif]  := FLargeur; {JP 05/06/07 : FQ 18664}
    G.ColLengths[colCollectif] := -1 ;
  end;


  If ModeleOk Then
    BEGIN
      If OnDev=OnDevise Then
      BEGIN
      G.ColWidths[colMontantDev]   := G.ColWidths[colMontantPivot] ;
      G.ColWidths[colMontantPivot] := FLargeur; {JP 05/06/07 : FQ 18664}
      G.ColLengths[colMontantPivot] := -1 ;
      LaColMontant                 := colMontantDev ;
      END ;
    END ;

  General  := '' ;
  GeneCoef := 1.0 ;

  if tt.Count >= 1 then
    begin
    O          := TOBM(tt[0]) ;
    General    := O.GetMvt('E_GENERAL') ;
    GeneDevise := O.GetMvt('E_DEVISE') ;
    DEV.Code   := GeneDevise ;
    GetInfosDevise(DEV) ;
    end
  else
    begin
    GeneDevise := V_PGI.DevisePivot ;
    DEV.Code   := GeneDevise ;
    GetInfosDevise(DEV) ;
    end ;

Case SorteLettre of
   tslCheque : begin
               // Gestion des éditions états / documents
               {JP 09/05/07 : FQ 20075 : c'est géré dans InitE_MODELE
               if bEtatAGL then
                 E_MODELE.Plus:='CLC'
               else
                 begin
                 E_MODELE.Datatype:='TTMODELELETTRECHQ' ;
                 E_MODELE.Plus:='';
                 end ;
               }
               ChqPrinted            := False ;
               LastOrigine           := 0 ;

               CompteBanque.Enabled  := False ;
               CompteBanque.Color    := clBtnFace ;

               E_MODELE.Visible      := ModeleOk  ;
               HModele.Visible       := Not ModeleOk ;

               BDown.Hint            := Msg.Mess[31] ;
               BUp.Hint              := Msg.Mess[35] ;
               bGroupTiers.Hint      := Msg.Mess[39] ;
               bCopyDetail.Hint      := Msg.Mess[43] ;
               bGroupTiersEche.Hint  := Msg.Mess[47] ;
               BImprimer.Hint        := Msg.Mess[50] ;
               BImprimer.Caption     := Msg.Mess[50] ;
               HTitreImprime.Caption := Msg.Mess[29] ;

               Caption               := Msg.Mess[51] ;

               end ;
   tslVir    : begin
               // Gestion des éditions états / documents
               {JP 09/05/07 : FQ 20075 : c'est géré dans InitE_MODELE
               if bEtatAGL then
                 E_MODELE.Plus       := 'CLV'
               else
                 begin
                 E_MODELE.Datatype   := 'TTMODELELETTREVIR' ;
                 E_MODELE.Plus       := '';
                 end ;}

               CompteBanque.Enabled  := False ;
               CompteBanque.Color    := clBtnFace ;

               E_MODELE.Visible      := True  ;
               HModele.Visible       := False ;
               If Not ModeleOK Then
                 BEGIN
                 Modele              := '' ;
                 E_MODELE.Value      := Modele ;
                 END ;

               BDown.Hint            := Msg.Mess[32] ;
               BUp.Hint              := Msg.Mess[36] ;
               bGroupTiers.Hint      := Msg.Mess[40] ;
               bCopyDetail.Hint      := Msg.Mess[44] ;
               bGroupTiersEche.Hint  := Msg.Mess[48] ;
               BImprimer.Hint        := Msg.Mess[25] ;
               BImprimer.Caption     := Msg.Mess[25] ;
               HTitreImprime.Caption := Msg.Mess[23] ;

               G.Titres[3]           := ' ;C;S' ;
               G.ColLengths[colNumCheque] := -1 ;
               G.ColWidths[colNumCheque]  := FLargeur; {JP 05/06/07 : FQ 18664}
               Caption               := Msg.Mess[27] ;
               GeneCoef              := 1.0 ;
               HelpContext           := 7595300 ;
               END ;
   tslPre : BEGIN
               // Gestion des éditions états / documents
               {JP 09/05/07 : FQ 20075 : c'est géré dans InitE_MODELE
               if bEtatAGL then
                 E_MODELE.Plus     := 'CLP'
               else
                 begin
                 E_MODELE.Datatype := 'TTMODELELETTREPRE' ;
                 E_MODELE.Plus     := '';
                 end ;}

               If Not ModeleOK Then
                 BEGIN
                 CompteBanque.Enabled := False ;
                 CompteBanque.Color   := clBtnFace ;
                 END ;

               E_MODELE.Visible       := True  ;
               HModele.Visible        := False ;
               If Not ModeleOk Then
                 BEGIN
                 Modele               := '' ;
                 E_MODELE.Value       := Modele ;
                 END ;

               HTitreImprime.Caption  := Msg.Mess[24] ;

               BDown.Hint             := Msg.Mess[33] ;
               BUp.Hint               := Msg.Mess[37] ;
               bGroupTiers.Hint       := Msg.Mess[41] ;
               bCopyDetail.Hint       := Msg.Mess[45] ;
               bGroupTiersEche.Hint   := Msg.Mess[49] ;
               BImprimer.Hint         := Msg.Mess[26] ;
               BImprimer.Caption      := Msg.Mess[26] ;

               G.Titres[3]            := ' ;C;S' ;
               G.ColLengths[colNumCheque] := -1 ;
               G.ColWidths[colNumCheque]  := FLargeur; {JP 05/06/07 : FQ 18664}
               Caption                := Msg.Mess[28] ;
               GeneCoef               := -1.0 ;
               HelpContext            := 7595300 ;
               END ;
   tslTraite : BEGIN
               // Gestion des éditions états / documents
               if bEtatAGL then
                 E_MODELE.Plus        := 'CLT'
               else
                 begin
                 E_MODELE.Datatype    := 'TTMODELELETTRETRA' ;
                 E_MODELE.Plus        := '';
                 end ;
               CompteBanque.Enabled   := False ;
               CompteBanque.Color     := clBtnFace ;

               E_MODELE.Visible       := True  ;
               HModele.Visible        := False ;
               Modele                 := ModeleTraite ;
               E_MODELE.Value         := Modele ;

               HTitreImprime.Caption  := Msg.Mess[8] ;

               BDown.Hint             := Msg.Mess[30] ;
               BUp.Hint               := Msg.Mess[34] ;
               bGroupTiers.Hint       := Msg.Mess[38] ;
               bCopyDetail.Hint       := Msg.Mess[42] ;
               bGroupTiersEche.Hint   := Msg.Mess[46] ;
               BImprimer.Hint         := Msg.Mess[11] ;
               BImprimer.Caption      := Msg.Mess[11] ;

               G.Titres[3]            := ' ;C;S' ;
               G.ColLengths[colNumCheque] := -1 ;
               G.ColWidths[colNumCheque]  := FLargeur; {JP 05/06/07 : FQ 18664}
               Caption                := Msg.Mess[12] ;
               GeneCoef               := -1.0 ;
               HelpContext            := 7595300 ;
               END ;
   tslBOR    : BEGIN
               // Gestion des éditions états / documents
               if bEtatAGL then
                 E_MODELE.Plus          := 'CLT'
               else
                 begin
                 E_MODELE.Datatype    := 'TTMODELELETTRETRA' ;
                 E_MODELE.Plus        := '';
                 end ;
               {JP 21/02/05 : si la banque prévisionnelle est connue on désactive la zone}
               if suiviMP and (not EstSerie(S3) and Assigned(TT) and (TOB(tt[0]).GetString('E_BANQUEPREVI') <> '')) then begin
                 CompteBanque.Enabled := False;
                 CompteBanque.Color   := clBtnFace;
                 CompteBanque.Text    := TOB(tt[0]).GetString('E_BANQUEPREVI');
                 HCompte.Caption      := RechDom('TZGENERAL', CompteBanque.Text, False);
               end
               else
               begin
                 CompteBanque.Enabled := True;
                 CompteBanque.Color   := clWindow;
               end;

               E_MODELE.Visible       := True ;
               HModele.Visible        := False ;
               Modele                 := ModeleTraite ;
               E_MODELE.Value         := Modele ;

               HTitreImprime.Caption  := Msg.Mess[13] ;


               BDown.Hint             := Msg.Mess[16] ;
               BUp.Hint               := Msg.Mess[17] ;
               bGroupTiers.Hint       := Msg.Mess[18] ;
               bCopyDetail.Hint       := Msg.Mess[19] ;
               bGroupTiersEche.Hint   := Msg.Mess[22] ;
               BImprimer.Hint         := Msg.Mess[14] ;
               BImprimer.Caption      := Msg.Mess[14] ;

               G.Titres[3]            := ' ;C;S' ;
               G.ColLengths[colNumCheque] := -1 ;
               G.ColWidths[colNumCheque]  := FLargeur; {JP 05/06/07 : FQ 18664}
               Caption                := Msg.Mess[15] ;
               HelpContext            := 7595500 ;
               END ;
   END ;

  AfficheGridEcheances ;
  ChercheModele(True) ;
  InitCheques ;

  G.Row            := 1 ;
  Action.Caption   := '' ;
  // Bouton valider accessible si traitement de l'escompte
  BValider.Enabled := smp In [smpDecChqEdt,smpDecVirEdt,smpDecVirInEdt] ;  //( TOBParamEsc <> nil ) and ( TOBParamEsc.Detail.Count > 0 ) ;
  // Bouton Test invisible depuis DOC -> ETAT
  BTest.Visible    := not bEtatAgl ;

  SourisNormale ;
end;

procedure TFDocRegl.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  GE.VidePile(False) ;
  G.VidePile(False) ;
end;

procedure TFDocRegl.bApercuClick(Sender: TObject);
begin
ImprimeDocRegl(False,False,True) ;
end;

procedure TFDocRegl.bTestClick(Sender: TObject);
begin
  // Pas d'équivalent en LanceEtat
  if not bEtatAGL then
    ImprimeDocRegl(True,True,False) ;
end;

procedure TFDocRegl.BImprimerClick(Sender: TObject);
begin
ImprimeDocRegl(True,False,False) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 03/12/2004
Modifié le ... :   /  /    
Description .. : SG6 12/03/2004 On rend possible l'edition dans le cas des 
Suite ........ : etats en mode eAglClient
Mots clefs ... : 
*****************************************************************}
procedure TFDocRegl.bModeleClick(Sender: TObject);
begin
  if bEtatAGL then
    begin
    if (SorteLettre=tslCheque) then EditEtat('E','CLC',Modele,True,nil,'',Caption) Else
    if (SorteLettre=tslVir)    then EditEtat('E','CLV',Modele,True,nil,'',Caption) Else
    if (SorteLettre=tslPre)    then EditEtat('E','CLP',Modele,True,nil,'',Caption) Else
                                    EditEtat('E','CLT',Modele,True,nil,'',Caption) ;
    end;

{$IFNDEF EAGLCLIENT}
  // Gestion des éditions états / documents
  if not(bEtatAGL) then
    begin
    if (SorteLettre=tslCheque) then EditDocumentS5S7('L','LCH',Modele,True) Else
    if (SorteLettre=tslVir)    then EditDocumentS5S7('L','LVI',Modele,True) Else
    if (SorteLettre=tslPre)    then EditDocumentS5S7('L','LPR',Modele,True) Else
                                    EditDocumentS5S7('L','LTR',Modele,True) ;
    end
{$ENDIF}
end;

procedure TFDocRegl.bDownClick(Sender: TObject);
begin

  if GE.Cells[0,GE.Row]<>'' then // FFF 118
    if (smp in [smpEncPreEdtNC, smpEncTraEdtNC, smpDecVirEdtNC, smpDecChqEdtNC, smpDecBorEdtNC])
      then EcheToCheque(GE.Row,G.Row,True,False)
      else EcheToCheque(GE.Row,G.Row,True,True) ;

  Action.Caption := '' ;

end;

procedure TFDocRegl.bUpClick(Sender: TObject);
var
  iRow : integer;   //SG6 06/01/05 Positionnement curseur dans grid FQ 13725
begin
  //SG6 06/01/05 Positionnement curseur dans grid FQ 13725
  iRow := G.Row;

  if G.Cells[colTiers,G.Row]<>'' then
    ChequeToEche(G.Row) ;

  Action.Caption := '' ;
  //SG6 06/01/05 Positionnement curseur dans grid FQ 13725
  if (G.Cells[colTiers,iRow]='') and (iRow > 1) then Dec(iRow);
  G.row := iRow;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 21/09/2004
Modifié le ... :   /  /
Description .. : Création de la liste des documents à imprimer en proposant
Suite ........ : un document par échéance
Mots clefs ... :
*****************************************************************}
procedure TFDocRegl.bCopyDetailClick(Sender: TObject);
var O     : TOBM ;
    i,j,k : Integer ;
begin

  // Parcours des échéances
  for i:=1 to GE.RowCount-1 do
    begin

    O := TOBM(tt[i-1]) ;
    if O.GetMvt('E_FLAGECR')='X' then continue ;

    // Détermination de la ligne de destination dans la grille des documents (1ère ligne vide)
    k := 0 ;
    for j:=1 to G.RowCount-1 do
      if G.Cells[colTiers,j]='' then
        begin
        k:=j ;
        break ;
        end ;
    if k=0 then exit ;

    // Passage de l'échéance ligne i dans la document ligne k
    EcheToCheque(i,k,False,True) ;

    end ;

  Action.Caption := Msg.Mess[7] ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 21/09/2004
Modifié le ... : 21/09/2004
Description .. : Création de la liste des documents à imprimer en 
Suite ........ : regroupeant par document les couples tiers  / banque 
Suite ........ : prévisionnelle
Mots clefs ... : 
*****************************************************************}
procedure TFDocRegl.bGroupTiersClick(Sender: TObject);
var O     : TOBM ;
    i,j,k : Integer ;
    Tiers,BanquePrevi : string ;
begin

  // Parcours des échéances
  for i:=1 to GE.RowCount-1 do
    begin

    O:=TOBM(tt[i-1]) ;
    if O.GetMvt('E_FLAGECR')='X' then continue ;

    // Détermination du couple Auxiliaire / Banque à traiter
    if TIDTIC
      then Tiers:=O.GetMvt('E_GENERAL')
      else Tiers:=O.GetMvt('E_AUXILIAIRE') ;
    BanquePrevi:=O.GetMvt('E_BANQUEPREVI') ;

    // -------------------------------------------------------------------------------------
    // -- Détermination de la ligne de destination dans le tableau des documents à éditer --
    // -------------------------------------------------------------------------------------
    // recherche de la 1ère échéance Tiers / Banque dans la liste des documents à éditer
    k:=0 ;
    for j := 1 to G.RowCount - 1 do
       begin
       if ( G.Cells[colTiers,j] = Tiers ) and ( G.Cells[colPrevi,j] = BanquePrevi ) then
         begin
         k:=j ;
         break ;
         end ;
       end ;
    // Si aucun document pour le couple existe, on recherche la 1è_re ligne vide
    if k=0 then
     for j:=1 to G.RowCount-1 do
       if G.Cells[colTiers,j]='' then
         begin
         k:=j ;
         break ;
         end ;
    // Si aucune ligne destination trouver on sort
    if k=0 then exit ;

    // ------------------------------------------------------------
    // -- Passage de l'échéance ligne i dans la document ligne k --
    // ------------------------------------------------------------
    EcheToCheque( i, k, False, Not ForcerGroupe ) ;

    end ;

  Action.Caption:=Msg.Mess[6] ;

end ;

procedure TFDocRegl.BValiderClick(Sender: TObject);
begin
  if not ChqPrinted then
    begin
    // L'utilisateur peut dorénavant valider sans imprimer dans les modules
    // gérant l'escompte
    if (smp In [smpDecChqEdt,smpDecVirEdt,smpDecVirInEdt]) then
      ImprimeDocRegl(True,False,False,False)
    else
      begin
      Msg.Execute(4,caption,'') ;
      exit ;
      end ;
    end ;
ModalResult:=mrOk ;
end;

procedure TFDocRegl.bAnnulerClick(Sender: TObject);
begin
if ChqPrinted then if Msg.Execute(1,caption,'')<>mrYes then exit ;
ModalResult:=mrCancel ;
end;

procedure TFDocRegl.FormCreate(Sender: TObject);
begin
  PopupMenu    := AddMenuPop(PopupMenu,'','') ;
  GeneRIB      := '' ;
  GeneDevise   := V_PGI.DevisePivot ;
  BQDevise     := V_PGI.DevisePivot ;

  ForcerGroupe := False ;
  {JP 05/06/07 : FQ 18664 : lar largeur des colonnes doit être à -1 si en NT par contre à 0 autrement}
  FLargeur := GetWidthColFromOS;
end;

procedure TFDocRegl.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFDocRegl.RenseigneRIB ;
Var Q : TQuery ;
BEGIN
GeneRIB:='' ; if CompteBanque.Text='' then Exit ;
// Recherche RIB
Q:=OpenSQL('SELECT BQ_ETABBQ, BQ_GUICHET, BQ_NUMEROCOMPTE, BQ_CLERIB, BQ_DOMICILIATION, BQ_PAYS'
          +' FROM BANQUECP WHERE BQ_GENERAL="'+Comptebanque.Text+'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"',True) ; // 19/10/2006 YMO Multisociétés
if Not Q.EOF then
   BEGIN
   GeneRIB:=EncodeRIB(Q.FindField('BQ_ETABBQ').AsString,Q.FindField('BQ_GUICHET').AsString,
                      Q.FindField('BQ_NUMEROCOMPTE').AsString,Q.FindField('BQ_CLERIB').AsString,
                      Q.FindField('BQ_DOMICILIATION').AsString,CodeISOduPays(Q.FindField('BQ_PAYS').AsString)) ;
   END ;
Ferme(Q) ;
END ;

procedure TFDocRegl.E_JOURNALChange(Sender: TObject);
begin
ChercheModele(False) ;
RenseigneRIB ;
end;

procedure TFDocRegl.E_MODELEChange(Sender: TObject);
begin
Modele:=E_MODELE.Value ;
end;

procedure TFDocRegl.CompteBanqueDblClick(Sender: TObject);
begin
CompteToBanque(False) ;
end;

procedure TFDocRegl.CompteBanqueExit(Sender: TObject);
begin
CompteToBanque(False) ;
end;

procedure TFDocRegl.bGroupTiersEcheClick(Sender: TObject);
var O     : TOBM ;
    i,j,k : Integer ;
    Tiers,BanquePrevi : string ;
    DateEche : TDateTime ;
begin

  // Parcours des échéances
  for i:=1 to GE.RowCount-1 do
    begin

    O:=TOBM(tt[i-1]) ;
    if O.GetMvt('E_FLAGECR')='X' then continue ;

    // Détermination du trio Auxiliaire / Banque / Date d'échéance à traiter
    if TIDTIC
      then Tiers:=O.GetMvt('E_GENERAL')
      else Tiers:=O.GetMvt('E_AUXILIAIRE') ;
   BanquePrevi:=O.GetMvt('E_BANQUEPREVI') ;
   DateEche:=O.GetMvt('E_DATEECHEANCE') ;

    // -------------------------------------------------------------------------------------
    // -- Détermination de la ligne de destination dans le tableau des documents à éditer --
    // -------------------------------------------------------------------------------------
    // recherche de la 1ère échéance Tiers / Banque / Date échéance dans la liste des documents à éditer
    k:=0 ;
    for j:=1 to G.RowCount-1 do
      begin
      if ( G.Cells[colTiers,j] = Tiers ) and ( G.Cells[colPrevi,j] = BanquePrevi )
                                         and ( StrToDate( G.Cells[ColDateEche,j] ) = DateEche ) then
        begin
        k := j ;
        break ;
        end ;
      end ;
    // Si aucun document pour le couple existe, on recherche la 1è_re ligne vide
    if k=0 then
      for j:=1 to G.RowCount-1 do
        if G.Cells[colTiers,j]='' then
          begin
          k := j ;
          break ;
          end ;
    // Si aucune ligne destination trouver on sort
    if k=0 then exit ;

    // ------------------------------------------------------------
    // -- Passage de l'échéance ligne i dans la document ligne k --
    // ------------------------------------------------------------
    EcheToCheque( i, k, False, Not ForcerGroupe ) ;

    end ;

  Action.Caption:=Msg.Mess[21] ;

end;

procedure TFDocRegl.GSorted(Sender: TObject);
var i : integer ;
begin
  for i:=G.RowCount-1 downto 1 do
    if G.Cells[0,i]='' then G.DeleteRow(i) ;
  G.RowCount:=G.RowCount+1 ;
end;

procedure TFDocRegl.SetNumChqDansEche( vStNumChq : String ; vInChq : Integer; MajBqPrevi : Boolean = False) ;
var sOrig  : String ;
    iEche  : Integer ;
    O      : TOBM ;
begin
  {On ne gère pas la banque prévisionnelle en S3}
  if EstSerie(S3) and MajBqPrevi then Exit;

  // Récupération des réfèrence aux lignes d'échéances
  sOrig := G.Cells[colOrigines, vInChq] ;
  // Pour chacune, maj du num de traite/chq
  while sOrig <> '' do
  begin
    // Recup ligne échéance
    iEche := StrToInt( (ReadTokenSt( sOrig ) ) ) ;
    if iEche > 0 then
    begin
      // MAJ OBM
      O     := TOBM( tt[iEche-1] ) ;

      {JP 21/02/05 : Mise à jour de la banque prévisionnelle dans l'échéance}
      if MajBqPrevi then begin
        if SuiviMP then
        begin
          if not O.FieldExists('BANQUEPREVI') then
            O.AddChampSupValeur('BANQUEPREVI', CompteBanque.Text)
          else
            O.PutValue('BANQUEPREVI', CompteBanque.Text);
        end;
      end

      {Mise à jour du numéro de chèque}
      else begin
        if not O.FieldExists('NUMTRAITECHQ') then
          O.AddChampSupValeur('NUMTRAITECHQ', vStNumChq )
        else
          O.PutValue('NUMTRAITECHQ', vStNumChq );
      end;{if else MajBqPrevi}

    end;{if iEche}
  end;{while sOrig}
end;
//{$ENDIF}

// 13581
{function TFDocRegl.GetNbPages(Const SQL : String) : Integer;
var
  q : TQuery;
begin
  Result := 0;

  if SQL = '' then exit;

  if IsNumeric(Trim(SQL)) then begin
    Result := StrToInt(SQL);
    Exit;
  end;

  Q := OpenSQL(SQL, True);
  while not Q.Eof do begin
    Inc(Result);
    Q.Next;
  end;

  Ferme(Q);
end;}

end.
