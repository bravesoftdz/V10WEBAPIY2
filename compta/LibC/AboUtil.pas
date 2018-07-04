 unit AboUtil;

interface

uses
  Windows,
  SysUtils,
  Classes,
  Hctrls,
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ELSE}
  UTOB,
{$ENDIF}
  Hqry,
  Ent1,
  HEnt1,
  Formule,
  LetBatch,
  ULibanalytique,
  SaisComm,
  SaisUtil,
  Echeance,
  HCompte,
  ed_tools,
  UtilSais,
  UtilPGI ,
  Ulibecriture;

Procedure EcrituresAbo ( CodeG : String3 ; M : RMVT ; DEV : RDEVISE ; TPIECE : TList ;
                         ContratG : String = '' ; LibContratG : String = '') ;
//Procedure MajSoldesPiece ( M : RMVT ) ;
                                           
implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  CPVersion,
  CPProcGen,
  {$ENDIF MODENT1}
  uLibTrSynchro; {FQ 18081}

Type Les_Totaux = RECORD
                  E_Debit,E_Credit,Y_Debit,Y_Credit : Array[T_DPE] of Double ;
                  Pourc : double ;
                  END ;

Var OGene : TOBM ;
    MGene : RMVT ;
    LastEche,LastVal : TDateTime ;
    LastMP,LastReg : String ;

function Load_Ecr ( St : hString ) : Variant ;
Var V    : Variant ;
BEGIN
V:=#0 ; Result:=V ;
St:=uppercase(Trim(St)) ; if St='' then Exit ;
{Zones entête}
if St='JOURNAL' then V:=MGene.Jal else
if St='DATECOMPTABLE' then V:=MGene.DateC else
if St='NATUREPIECE' then V:=MGene.Nature else
if St='NUMEROPIECE' then V:=MGene.Num else
if St='DEVISE' then V:=MGene.CodeD else
if St='ETABLISSEMENT' then V:=MGene.Etabl else
{Zones Général}
if Copy(St,1,2)='G_' then
   BEGIN
   System.Delete(St,1,2) ;
   if St='GENERAL' then V:=OGene.GetMvt('E_GENERAL') else V:=RechercheLente('G_'+St,OGene.GetMVt('E_GENERAL')) ;
   END else
{Zones Auxiliaire}
if Copy(St,1,2)='T_' then
   BEGIN
   System.Delete(St,1,2) ;
   if St='AUXILIAIRE' then V:=OGene.GetMvt('E_AUXILIAIRE') else V:=RechercheLente('T_'+St,OGene.GetMvt('E_AUXILIAIRE')) ;
   END else
{Zones Journal}
if Copy(St,1,2)='J_' then
   BEGIN
   V:=RechercheLente('J_'+St,MGene.JAL) ;
   END else
{Zones Ecriture}
   BEGIN
   if Copy(St,1,2)='E_' then System.Delete(St,1,2) ;
   V:=OGene.GetMvt('E_'+St) ;
   END ;
Load_Ecr:=V ;
END ;

Procedure Abo_RenseigneBase ( OEcr,OEg : TOBM ; M : RMVT ; DEV : RDEVISE ; Dernier : boolean ; Var LesTot : Les_Totaux ) ;
Var XD,XC,DP,CP : Double ;
    St                : String ;
BEGIN
{Entête}
OEcr.PutMvt('E_ETABLISSEMENT',M.Etabl)  ; OEcr.PutMvt('E_EXERCICE',M.Exo) ;
OEcr.PutMvt('E_DEVISE',M.CodeD)         ; OEcr.PutMvt('E_QUALIFPIECE',M.Simul) ;
OEcr.PutMvt('E_NATUREPIECE',M.Nature)   ; OEcr.PutMvt('E_DATECOMPTABLE',M.DateC) ;
{$IFNDEF SPEC302}
OEcr.PutMvt('E_PERIODE',GetPeriode(M.DateC)) ; OEcr.PutMvt('E_SEMAINE',NumSemaine(M.DateC)) ;
{$ENDIF}
OEcr.PutMvt('E_NUMEROPIECE',M.Num)      ; OEcr.PutMvt('E_JOURNAL',M.Jal) ;
OEcr.PutMvt('E_DATETAUXDEV',M.DateTaux) ; OEcr.PutMvt('E_TAUXDEV',M.TauxD) ;
OEcr.PutMvt('E_CREERPAR','GEN')         ; OEcr.PutMvt('E_QUALIFORIGINE','ABO') ;
if M.Valide then OEcr.PutMvt('E_VALIDE','X') else OEcr.PutMvt('E_VALIDE','-') ;
{Détail}
OEcr.PutMvt('E_NUMLIGNE',OEg.GetMvt('EG_NUMLIGNE')) ;
OEcr.PutMvt('E_GENERAL',OEg.GetMvt('EG_GENERAL'))   ; OEcr.PutMvt('E_AUXILIAIRE',OEg.GetMvt('EG_AUXILIAIRE')) ;
OEcr.PutMvt('E_MODEPAIE',OEg.GetMvt('EG_MODEPAIE')) ;
CSupprimerInfoLettrage( OEcr ) ; // FQ 17419
// FQ 12090
OEcr.PutMvt('E_RIB',OEg.GetMvt('EG_RIB'));
if ((M.CodeD<>V_PGI.DevisePivot) and (M.CodeD<>'')) then
   BEGIN
   XD:=Valeur(OEg.GetMvt('EG_DEBITDEV')) ; XC:=Valeur(OEg.GetMvt('EG_CREDITDEV')) ;
   OEcr.PutMvt('E_DEBITDEV',XD) ; OEcr.PutMvt('E_CREDITDEV',XC) ;
   DP:=DeviseToPivot(XD,DEV.Taux,DEV.Quotite) ; CP:=DeviseToPivot(XC,DEV.Taux,DEV.Quotite) ;
   LesTot.E_Debit[tdPivot]:=LesTot.E_Debit[tdPivot]+DP ; LesTot.E_Credit[tdPivot]:=LesTot.E_Credit[tdPivot]+CP ;
   if Dernier then
      BEGIN
      if DP<>0 then DP:=DP-(LesTot.E_Debit[tdPivot]-LesTot.E_Credit[tdPivot])
               else CP:=CP+(LesTot.E_Debit[tdPivot]-LesTot.E_Credit[tdPivot]) ;
      END ;
    OEcr.PutMvt('E_DEBIT',DP) ; OEcr.PutMvt('E_CREDIT',CP) ;
   END else
   BEGIN
   XD:=Valeur(OEg.GetMvt('EG_DEBITDEV')) ; XC:=Valeur(OEg.GetMvt('EG_CREDITDEV')) ;
   DP:=XD ; CP:=XC ;

   LesTot.E_Debit[tdPivot]:=LesTot.E_Debit[tdPivot]+DP   ; LesTot.E_Credit[tdPivot]:=LesTot.E_Credit[tdPivot]+CP ;
   LesTot.E_Debit[tdDevise]:=LesTot.E_Debit[tdDevise]+XD ; LesTot.E_Credit[tdDevise]:=LesTot.E_Credit[tdDevise]+XC ;
   if Dernier then
      BEGIN
      if DP<>0 then DP:=DP-(LesTot.E_Debit[tdPivot]-LesTot.E_Credit[tdPivot])
               else CP:=CP+(LesTot.E_Debit[tdPivot]-LesTot.E_Credit[tdPivot]) ;
      if XD<>0 then XD:=XD-(LesTot.E_Debit[tdDevise]-LesTot.E_Credit[tdDevise])
               else XC:=XC+(LesTot.E_Debit[tdDevise]-LesTot.E_Credit[tdDevise]) ;
      END ;
   OEcr.PutMvt('E_DEBITDEV',XD)  ; OEcr.PutMvt('E_CREDITDEV',XC) ;
   OEcr.PutMvt('E_DEBIT',DP)     ; OEcr.PutMvt('E_CREDIT',CP) ;
   END ;
OGene:=OEcr ; MGene:=M ;
OEcr.PutMvt('E_REFINTERNE',GFormule(OEg.GetMvt('EG_REFINTERNE'),Nil,Load_Ecr,2)) ;
OEcr.PutMvt('E_LIBELLE',GFormule(OEg.GetMvt('EG_LIBELLE'),Nil,Load_Ecr,2)) ;
{#TVAENC}
if VH^.OuiTvaEnc then
   BEGIN
   St:=OEg.GetMvt('EG_TVAENCAIS') ; if St<>'' then OEcr.PutMvt('E_TVAENCAISSEMENT',St) ;
   St:=OEg.GetMvt('EG_TVA') ; if St<>'' then OEcr.PutMvt('E_TVA',St) ;
   END ;
END ;

Function Abo_RenseigneRegle ( OEcr : TOBM ; Var ModeR : T_MODEREGL ; M : RMVT ; DEV : RDEVISE ) : boolean ;
Var StM : String ;
    CGen  : TGGeneral ;
    Gene  : String ;
    OkMultiEche : boolean ;
BEGIN
Result:=False ;
StM:=ModeR.ModeInitial ; FillChar(ModeR,Sizeof(ModeR),#0) ; if STM='' then Exit ;
Gene:=OEcr.GetMvt('E_GENERAL') ; CGen:=TGGeneral.Create(Gene) ;
if ((Not CGen.Collectif) and (Lettrable(CGen)=NonEche)) then Exit ;
OkMultiEche:=True ; if ((CGen.Collectif) and (Ventilable(CGen,0))) then OkMultiEche:=False ;
if LienS1S3 then OkMultiEche:=False ; 
With ModeR do
   BEGIN
   Action:=taCreat ; ModeInitial:=StM ;
   TotalAPayerP:=OEcr.GetMvt('E_DEBIT')+OEcr.GetMvt('E_CREDIT') ;
   TotalAPayerD:=OEcr.GetMvt('E_DEBITDEV')+OEcr.GetMvt('E_CREDITDEV') ;
   CodeDevise:=M.CodeD ; Symbole:=DEV.Symbole ; Quotite:=DEV.Quotite ;
   TauxDevise:=M.TauxD ; Decimale:=DEV.Decimale ;
   DateFact:=M.DateC ; DateBL:=M.DateC ; 
// GP REGL
   DateFactExt:=DateFact ;
   Aux:=OEcr.GetMvt('E_AUXILIAIRE') ; if Aux='' then Aux:=OEcr.GetMvt('E_GENERAL') ;
   END ;
CalculModeregle(ModeR,OkMultiEche) ;
Result:=True ;
END ;

Procedure Abo_RecupEche( OEcr : TOBM ; R : T_ModeRegl ; NumEche : integer ; M : RMVT ) ;
Var Deb  : boolean ;
BEGIN
OEcr.PutValue('E_NUMECHE',NumEche) ;
OEcr.PutValue('E_ECHE','X') ;
OEcr.PutValue('E_NIVEAURELANCE',R.TabEche[NumEche].NiveauRelance) ;
OEcr.PutValue('E_MODEPAIE',R.TabEche[NumEche].ModePaie) ;
OEcr.PutValue('E_DATEECHEANCE',R.TabEche[NumEche].DateEche) ;
OEcr.PutValue('E_DATEVALEUR',R.TabEche[NumEche].DateEche) ;
LastMP:=R.TabEche[NumEche].ModePaie ; ;
LastEche:=R.TabEche[NumEche].DateEche ;
LastVal:=R.TabEche[NumEche].DateEche ;
OEcr.PutValue('E_COUVERTURE',R.TabEche[NumEche].Couverture) ;
OEcr.PutValue('E_COUVERTUREDEV',R.TabEche[NumEche].CouvertureDev) ;
OEcr.PutValue('E_ETATLETTRAGE',R.TabEche[NumEche].EtatLettrage) ;
OEcr.PutValue('E_LETTRAGE',R.TabEche[NumEche].CodeLettre) ;
OEcr.PutValue('E_LETTRAGEDEV',R.TabEche[NumEche].LettrageDev) ;
OEcr.PutValue('E_DATEPAQUETMAX',R.TabEche[NumEche].DatePaquetMax) ;
OEcr.PutValue('E_DATEPAQUETMIN',R.TabEche[NumEche].DatePaquetMin) ;
OEcr.PutValue('E_DATERELANCE',R.TabEche[NumEche].DateRelance) ;
Deb:=(OEcr.GetMvt('E_DEBIT')<>0) ;
OEcr.PutValue('E_DEBIT',0) ; OEcr.PutValue('E_CREDIT',0) ;
//QEcr.FindField('E_TYPEMVT').AsString:=QuelTypeMvt(OEcr.GetMvt('E_NUMLIGNE')=1,NumEche=1,True) ;
if Deb then
   BEGIN
   OEcr.PutValue('E_DEBIT',R.TabEche[NumEche].MontantP) ;
   OEcr.PutValue('E_DEBITDEV',R.TabEche[NumEche].MontantD) ;
   if R.TabEche[NumEche].MontantP>0 then OEcr.PutValue('E_ENCAISSEMENT','ENC')
                                    else OEcr.PutValue('E_ENCAISSEMENT','DEC') ;
   END else
   BEGIN
   OEcr.PutValue('E_CREDIT',R.TabEche[NumEche].MontantP );
   OEcr.PutValue('E_CREDITDEV',R.TabEche[NumEche].MontantD );
   if R.TabEche[NumEche].MontantP>0 then OEcr.PutValue('E_ENCAISSEMENT','DEC' )
                                    else OEcr.PutValue('E_ENCAISSEMENT','ENC' );
   END ;
OEcr.PutValue('E_ENCAISSEMENT',SensEnc(OEcr.GetMvt('E_DEBIT'),OEcr.GetMvt('E_CREDIT'))) ;
OEcr.PutValue('E_ORIGINEPAIEMENT',R.TabEche[NumEche].DateEche) ;
{#TVAENC}
if VH^.OuiTvaEnc then
   BEGIN
   if ((M.Nature='FC') or (M.Nature='AC') or (M.Nature='FF') or (M.Nature='AF')) then
    if EstJalFact(M.Jal) then OEcr.PutValue('E_EMETTEURTVA','X') ;
   END ;
END ;

Function TrouveUnMP : String ;
Var QQ : TQuery ;
BEGIN
Result:='' ;
QQ:=OpenSQL('SELECT MP_MODEPAIE FROM MODEPAIE',True) ;
if Not QQ.EOF then Result:=QQ.Fields[0].AsString ;
Ferme(QQ) ;
if Result<>'' then LastMP:=Result ;
END ;

Procedure Abo_RecupTronc ( OEcr : TOBM ; M : RMVT ; CodeAbo,LibAbo : String ; vTSA : TList ) ;
Var CGen    : TGGeneral ;
    CAux    : TGTiers ;
    NatG,NatT : String3 ;
BEGIN
{Création CGen, CAux}
CAux:=Nil ; NatG:='' ; NatT:='' ;
CGen:=TGGeneral.Create(OEcr.GetMvt('E_GENERAL')) ; NatG:=CGen.NatureGene ;
if OEcr.GetMvt('E_AUXILIAIRE')<>'' then
   BEGIN
   CAux:=TGTiers.Create(OEcr.GetMvt('E_AUXILIAIRE')) ; NatT:=CAux.NatureAux ;
   END ;
{Remplissage Champs}
OEcr.SetCotation(0) ;
OEcr.SetMPACC ;
//OEcr.EgalChamps(QEcr) ;
OEcr.PutValue('E_QUALIFPIECE',M.Simul) ;
if NatT<>'' then OEcr.PutValue('E_TYPEMVT',QuelTypeMvt(CAux.Auxi,NatT,M.Nature) )
            else OEcr.PutValue('E_TYPEMVT',QuelTypeMvt(CGen.General,NatG,M.Nature) );
if CAux<>Nil then OEcr.PutValue('E_REGIMETVA',CAux.RegimeTVA) ;
OEcr.PutValue('E_ECRANOUVEAU','N') ;
OEcr.PutValue('E_BUDGET','') ; OEcr.PutValue('E_ECHE','-') ;
OEcr.PutValue('E_LIBRETEXTE8',CodeAbo) ;
OEcr.PutValue('E_LIBRETEXTE9',LibAbo) ;
// GG COM
OEcr.PutValue('E_IO','X') ;
if CAux<>Nil then BEGIN OEcr.PutValue('E_REGIMETVA',CAux.RegimeTVA) ; LastReg:=CAux.RegimeTVA ; END ;
if CGen<>Nil then
   BEGIN
   if ((OEcr.GetValue('E_TVA')='') or (Not VH^.OuiTvaEnc)) then OEcr.PutValue('E_TVA',CGen.Tva) ;
   OEcr.PutValue('E_TPF',CGen.Tpf) ;
   OEcr.PutValue('E_BUDGET',CGen.Budgene) ;
   if Lettrable(CGen)=MonoEche then
      BEGIN
      OEcr.PutValue('E_NUMECHE',1) ;
      OEcr.PutValue('E_ECHE','X') ;
      OEcr.PutValue('E_ENCAISSEMENT',SensEnc(OEcr.GetMvt('E_DEBIT'),OEcr.GetMvt('E_CREDIT'))) ;
      OEcr.PutValue('E_ORIGINEPAIEMENT',OEcr.GetValue('E_DATEECHEANCE')) ;
      if LastMP<>'' then OEcr.PutValue('E_MODEPAIE',LastMP)
                    else OEcr.PutValue('E_MODEPAIE',TrouveUnMP) ;
      if LastEche>0 then OEcr.PutValue('E_DATEECHEANCE',LastEche)
                    else OEcr.PutValue('E_DATEECHEANCE',OEcr.GetValue('E_DATECOMPTABLE')) ;
      if LastVal>0 then OEcr.PutValue('E_DATEVALEUR',LastVal)
                   else OEcr.PutValue('E_DATEVALEUR',OEcr.GetValue('E_DATECOMPTABLE')) ;
      if OEcr.GetValue('E_REGIMETVA')='' then
         BEGIN
         if LastReg<>'' then OEcr.PutValue('E_REGIMETVA',LastReg)
                        else OEcr.PutValue('E_REGIMETVA',VH^.RegimeDefaut) ;
         END ;
      {#TVAENC}
      if VH^.OuiTvaEnc then
         BEGIN
         if ((M.Nature='FC') or (M.Nature='AC') or (M.Nature='FF') or (M.Nature='AF')) then
          if EstJalFact(M.Jal) then OEcr.PutValue('E_EMETTEURTVA','X') ;
         END ;
      END ;
   if Ventilable(CGen,0) then OEcr.PutValue('E_ANA','X') ;
   END ;
 AjouteAno(vTSA,OEcr,NatG,false) ;
END ;

procedure Abo_RecupAna ( OAna , OEcr : TOBM ; DEV : RDEVISE ; RY : R_YV ; Arrond,Premier : boolean ; Var LesTot : Les_Totaux ) ;
Var
 Ax    : String3 ;
 Pourc,DD,CD,DP,CP,TotPivot,TotDevise : Double ;
BEGIN
Ax:=RY.Axe ;
TotPivot:=OEcr.GetMvt('E_DEBIT')+OEcr.GetMvt('E_CREDIT') ;
TotDevise:=OEcr.GetMvt('E_DEBITDEV')+OEcr.GetMvt('E_CREDITDEV') ;
EcrToAna(OEcr,OAna) ;
OAna.PutMvt('Y_SECTION',RY.Section)     ; OAna.PutMvt('Y_NUMVENTIL',RY.NumVentil) ;
OAna.PutMvt('Y_TOTALECRITURE',TotPivot) ; OAna.PutMvt('Y_TOTALDEVISE',TotDevise) ;
if OAna.GetMvt('Y_NUMVENTIL')=1 then OAna.PutMvt('Y_TYPEMVT','AE') else  OAna.PutMvt('Y_TYPEMVT','AL') ;
Pourc:=RY.Pourcentage  ; OAna.PutMvt('Y_POURCENTAGE',Pourc) ;
LesTot.Pourc:=LesTot.Pourc+Pourc ;
DD:=OEcr.GetMvt('E_DEBITDEV')  ; CD:=OEcr.GetMvt('E_CREDITDEV') ;
DP:=OEcr.GetMvt('E_DEBIT')     ; CP:=OEcr.GetMvt('E_CREDIT') ;
if ((Arrond) and (Not Premier) and (Arrondi(LesTot.Pourc-100.0,ADecimP)=0)) then
   BEGIN
   if DP<>0 then
      BEGIN
      DD:=TotDevise-LesTot.Y_Debit[tdDevise] ; CD:=0 ;
      DP:=TotPivot-LesTot.Y_Debit[tdPivot] ; CP:=0 ;
      END else
      BEGIN
      CD:=TotDevise-LesTot.Y_Credit[tdDevise] ; DD:=0 ;
      CP:=TotPivot-LesTot.Y_Credit[tdPivot] ; DP:=0 ;
      END ;
   END else
   BEGIN
   DP:=Arrondi(DP*Pourc/100.0,V_PGI.OkDecV)    ; CP:=Arrondi(CP*Pourc/100.0,V_PGI.OkDecV) ;
   DD:=Arrondi(DD*Pourc/100.0,DEV.Decimale) ; CD:=Arrondi(CD*Pourc/100.0,DEV.Decimale) ;

   LesTot.Y_Debit[tdDevise]:=LesTot.Y_Debit[tdDevise]+DD ;
   LesTot.Y_Debit[tdPivot]:=LesTot.Y_Debit[tdPivot]+DP ;
   LesTot.Y_Credit[tdDevise]:=LesTot.Y_Credit[tdDevise]+CD ;
   LesTot.Y_Credit[tdPivot]:=LesTot.Y_Credit[tdPivot]+CP ;
   END ;
OAna.PutMvt('Y_DEBIT',DP)     ; OAna.PutMvt('Y_CREDIT',CP) ;
OAna.PutMvt('Y_DEBITDEV',DD)  ; OAna.PutMvt('Y_CREDITDEV',CD) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 02/06/2006
Modifié le ... :   /  /    
Description .. : - FB 10678 - LG - 02/06/2006 - gestion des devises ds les 
Suite ........ : ano dynamiques
Mots clefs ... : 
*****************************************************************}
Procedure EcrituresAbo ( CodeG : String3 ; M : RMVT ; DEV : RDEVISE ; TPIECE : TList ;
                         ContratG : String = '' ; LibContratG : String = '') ;
Var QEG,QAG,QGene : TQuery ;
    OEG,OEcr,OAna,OHisto : TOBM ;
    i,NumL,NbL,k,NumAxe : integer ;
    OldAxe,CpteGene,NewAxe : String ;
    ModeR    : T_MODEREGL ;
    RY       : R_YV ;
    Premier  : boolean ;
    TVentAna : TList ;
    OkAx     : Array[1..MaxAxe] of Boolean ;
    LesTot   : Les_Totaux ;
    lTSA     : TList ;
BEGIN
lTSA := TList.Create ;
QEG:=OpenSQL('Select Count(EG_TYPE) from ECRGUI Where EG_TYPE="ABO" and EG_GUIDE="'+CodeG+'"',True) ;
NbL:=QEG.Fields[0].AsInteger ;
Ferme(QEG) ;
QEG:=OpenSQL('Select * from ECRGUI Where EG_TYPE="ABO" and EG_GUIDE="'+CodeG+'" Order by EG_TYPE,EG_GUIDE,EG_NUMLIGNE',True) ;
OEG:=TOBM.Create(EcrGui,'',True) ; OEcr:=TOBM.Create(EcrGen,'',True) ;
TVentAna:=TList.Create ; LastMP:='' ; LastEche:=0 ; LastVal:=0 ; LastReg:='' ;
FillChar(LesTot,Sizeof(LesTot),#0) ;
While Not QEG.EOF do
   BEGIN
   {Boucle sur les écritures guidées générales}
   OEG.PutMvt('EG_DEBITDEV','');		// Modif VL V585.5
   OEG.PutMvt('EG_CREDITDEV','');		// Modif VL V585.5
   OEG.PutMvt('EG_AUXILIAIRE','');  // Fiche 11606
   OEG.ChargeMvt(QEG) ; NumL:=OEG.GetMvt('EG_NUMLIGNE') ; M.NumLigne:=NumL ;
   Abo_RenseigneBase(OEcr,OEG,M,DEV,NumL=NbL,LesTot) ;
   ModeR.ModeInitial:=OEg.GetMvt('EG_MODEREGLE') ;
   if Abo_RenseigneRegle(OEcr,ModeR,M,DEV) then
      BEGIN
      for i:=1 to ModeR.NbEche do
          BEGIN
          Abo_RecupTronc(OEcr,M, ContratG, LibContratG,lTSA) ;
          Abo_RecupEche(OEcr,ModeR,i,M) ;

          {JP 23/05/06 : FQ 18081 : Mise à jour de E_TRESOSYNCHRO pour les écritures d'abonnement}
          if EstComptaTreso then
            MajTresoEcritureTOB(OEcr, taCreat);

          OEcr.InsertDB(nil) ;
          MajSoldeCompteTOB(OEcr, True) ;         {FP 03/11/2005 FQ16492}
          if ((NumL=1) and (i=1)) then if V_PGI.IoError=oeOk then
             BEGIN
             OHisto:=TOBM.Create(EcrGen,'',True) ;
             OHisto.Dupliquer(OEcr,true,true) ;
            // OHisto.ChargeMvt(OEcr) ;
             TPIECE.Add(OHisto) ;
             END ;
          END ;
      END else
      BEGIN
      Abo_RecupTronc(OEcr,M,ContratG, LibContratG,lTSA) ;

      {JP 23/05/06 : FQ 18081 : Mise à jour de E_TRESOSYNCHRO pour les écritures d'abonnement}
      if EstComptaTreso then
        MajTresoEcritureTOB(OEcr, taCreat);

      OEcr.InsertDB(nil) ;
      MajSoldeCompteTOB(OEcr, True) ;    {FP 03/11/2005 FQ16492}
      if NumL=1 then if V_PGI.IoError=oeOk then
         BEGIN
         OHisto:=TOBM.Create(EcrGen,'',True) ;
         OHisto.Dupliquer(OEcr,true,true) ;
         //OHisto.ChargeMvt(OEcr) ;
         TPIECE.Add(OHisto) ;
         END ;
      END ;
   {Recherche des ventilations Guide / Pre-ventil / Attente}
   FillChar(OkAx,Sizeof(OkAx),#0) ;
   QAG:=OpenSQL('Select * from ANAGUI Where AG_TYPE="ABO" and AG_GUIDE="'+CodeG+'" AND AG_NUMLIGNE='+IntToStr(NumL)
               +' Order By AG_TYPE, AG_GUIDE, AG_NUMLIGNE, AG_AXE, AG_NUMVENTIL',True) ;
   if Not QAG.EOF then
      BEGIN
      FillChar(LesTot.Y_Debit,Sizeof(LesTot.Y_Debit),#0) ;
      FillChar(LesTot.Y_Credit,Sizeof(LesTot.Y_Credit),#0) ;
      LesTot.Pourc:=0 ;
      Premier:=True ;
      While Not QAG.EOF do
         BEGIN
         {Guide analytique}
        // QAna.Insert ;
         RY.Section:=QAG.FindField('AG_SECTION').AsString ; RY.Axe:=QAG.FindField('AG_AXE').AsString ;
         if Length(RY.Axe)>=2 then NumAxe:=Ord(RY.Axe[2])-48 else NumAxe:=1 ; OkAx[NumAxe]:=True ;
         RY.NumVentil:=QAG.FindField('AG_NUMVENTIL').AsInteger  ; RY.Pourcentage:=Valeur(QAG.FindField('AG_POURCENTAGE').AsString) ;
         if Premier then OldAxe:=RY.Axe else if ((OldAxe<>RY.Axe)) then
            BEGIN
            FillChar(LesTot.Y_Debit,Sizeof(LesTot.Y_Debit),#0) ;
            FillChar(LesTot.Y_Credit,Sizeof(LesTot.Y_Credit),#0) ;
            LesTot.Pourc:=0 ;
            END ;
         OAna:=TOBM.Create(EcrAna,RY.Axe,True) ;
         Abo_RecupAna(OAna,OEcr,DEV,RY,True,Premier,LesTot) ;
         OAna.InsertDB(nil) ;
         MajSoldeSectionTOB(OAna, True) ;   {FP 03/11/2005 FQ16492}
         OAna.Free ;
         QAG.Next ; Premier:=False ;
         END ;
      END else
      BEGIN
      {Autres ventilations possibles}
      CpteGene:=OEcr.GetMvt('E_GENERAL') ;
      QGene:=OpenSQL('Select G_VENTILABLE1, G_VENTILABLE2, G_VENTILABLE3, G_VENTILABLE4, G_VENTILABLE5 from GENERAUX Where G_GENERAL="'+CpteGene+'"',True) ;
      if Not QGene.EOF then
         BEGIN
         for i:=0 to 4 do if ((QGene.Fields[i].AsString='X') and (Not OkAx[i+1])) then
             BEGIN
             VideListe(TVentAna) ; NewAxe:='A'+IntToStr(i+1) ;
             VentileGenerale(NewAxe,OEcr,DEV,TVentAna,True) ;
             for k:=0 to TVentAna.Count-1 do
                 BEGIN
                 OAna:=TOBM(TVentAna[k]) ;
                 OAna.InsertDB(nil);
                 MajSoldeSectionTOB(OAna, True) ;   {FP 03/11/2005 FQ16492}
                 END ;
             END ;
         END ;
      Ferme(QGene) ;
      END ;
   Ferme(QAG) ;
   QEG.Next ;
   END ;
Ferme(QEG) ;
OEG.Free ; OEcr.Free ; TVentAna.Free ;
if lTSA <> nil then
begin
 if not ExecReqMAJAno(lTSA) then
   begin
     V_PGI.IoError := oeSaisie ;
     exit ;
   end ;
end ;
if lTSA <> nil then
begin
for i:=0 to lTSA.Count-1 do
 if assigned(lTSA[i]) then Dispose(lTSA[i]);
lTSA.Free ;
end ;
END ;
(*
Procedure MajSoldesPiece ( M : RMVT ) ;
Var Q : TQuery ;
BEGIN
{Q:=OpenSQL('Select * from ECRITURE Where '+WhereEcriture(tsGene,M,False),True) ;
MajDesSoldesTOB ( Q, False ) ;   // à mettre en standard
While Not Q.EOF do BEGIN MajSoldeCompte(Q) ; Q.Next ; END ;
Ferme(Q) ;
Q:=OpenSQL('Select * from ANALYTIQ Where '+WhereEcriture(tsAnal,M,False),True) ;

While Not Q.EOF do BEGIN MajSoldeSection(Q,True) ; Q.Next ; END ;
Ferme(Q) ;  }
END ;
*)
end.
