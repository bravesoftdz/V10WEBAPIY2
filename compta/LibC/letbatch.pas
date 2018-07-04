unit LetBatch;

interface

uses Windows, Messages, SysUtils, Classes, Graphics, Controls, Hqry, Ent1,HEnt1,
     SaisUtil, HCtrls, SaisComm,UtilSais,ParamSoc,
{$IFDEF EAGLCLIENT}
{$ELSE EAGLCLIENT}
   DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
   DBGrids,
  {$IFNDEF SANSCOMPTA}
     UtilSoc, HCompte,
  {$ENDIF SANSCOMPTA}
{$ENDIF EAGLCLIENT}
  {$IFDEF MODENT1}
  CPProcMetier,
  CPTypeCons,
  CPProcGen,
  {$ENDIF MODENT1}
  Utob,
  ULibAnalytique,
  Forms, StdCtrls, LettUtil,ed_tools, UtilPGI,HMsgBox,uEntCommun ;


{$IFNDEF EAGLCLIENT}
  {$IFNDEF EAGLSERVER}
  {$IFNDEF ERADIO}
    procedure MajSoldeCompte ( TEcr : TDataSet ) ; // en eagl voir MajSoldeCompteTOB
  {$ENDIF !ERADIO}
  {$ENDIF EAGLSERVER}
  {$IFNDEF SANSCOMPTA}
    procedure MajSoldeSection ( TAna : TDataSet ; Plus : boolean ) ; // en eagl voir MajSoldeSectionTOB
  {$ENDIF !SANSCOMPTA}
{$ENDIF !EAGLCLIENT}

Procedure VentileGenerale ( Ax : String3 ; OEcr : TOBM ; DEV : RDevise ; TVentAna : TList ; Preventil : boolean; aDossier : string = '') ;
{ déplacement dans Ulibanalytique
Procedure EcrToAna ( OEcr,OAna : TOBM ) ;
Procedure FiniMontantsAna ( OAna : TOBM ; RY : R_YV ; DEV : RDEVISE ; UnSeul : boolean; aDossier : string = '') ;
}
Procedure EquilibreDerniere ( TVentAna : TList ; OEcr : TOBM ; RY : R_YV ; DEV : RDevise ; Force : boolean; aDossier : string = '') ;
{$IFDEF EAGLCLIENT}
procedure RemplirOFromM ( O : TOBM ;  M : RMVT ; Totale : boolean ) ;
{$ENDIF}
{$IFNDEF SANSCOMPTA}
procedure Lett_RempliBase ( TEcr : TQUERY ; OL : TOBM ; Client : Boolean ; Regul : PG_Lettrage ; ForceSaisieEuro : String = '') ;
procedure Lett_NewLigne ( TEcr : TQUERY ; M : RMVT ) ;
procedure CalculeMontants( XX : T_ECARTCHANG ; M : RMVT ; Var XP,XD : Double ) ;
procedure Lett_TraiteMontant ( TEcr : TQUERY ; XX : T_ECARTCHANG ; M : RMVT ; Premier,Gaffe : boolean ; it : integer ) ;
function VentilLaRegul ( TEcr : TDataSet ; DEV : RDEVISE ; RL : RLETTR ; XX : T_ECARTCHANG ) : boolean ;
Procedure RensTvaEnc ( TEcr : TDataSet ) ;
function Signe ( X : Double ) : integer ;
function  CreerPartieDoubleLett ( M : RMVT ; RL : RLETTR ; XX : T_ECARTCHANG ; OL : TOBM ; ForceSaisieEuro : String = '') : TOBM ;
procedure ChargeLettrage ( XOrig : TL_Rappro ; LOrig : TList ; DEV : RDEVISE ) ;
procedure LettrageEncaDeca ( TEcheOrig : TList ; RR : RMVT ; DEV : RDEVISE ; TOBNumChq : TOB = nil ) ; // Modif SBO Fiche 12017 : pb maj E_NUMTRAITECHQ
Procedure CompleteNumTraiteChq( TOBNumChq : TOB ; XOrig : TL_Rappro ; vStCode : String ) ; // Ajout SBO Fiche 12017 : pb maj E_NUMTRAITECHQ
{$ENDIF SANSCOMPTA}

Procedure CalculCouv ( X : TL_Rappro ; DEV : RDEVISE ; Total : boolean ; Var CP,CD : Double ) ;
procedure FaireLettrage (LEcr : TList ; Total : boolean ; DMin,DMax : TDateTime ; Var CodeL : String ; DEV : RDEVISE ; EncaDeca : boolean ; CurTRACHQ : String = '' ;
                         ChampMarque : String = '' ; ValeurMarque : String ='' ; OptimiseUpdate : Boolean = FALSE) ;
Function RegulAFaire(Var Dev : RDevise ; OkDev,LettrageOppose : Boolean ; EcartRegulDevise : Double) : Boolean ;
Function LettrerUnPaquet ( LEcr : TList ; Vide,EncaDeca : boolean ;
                           ChampMarque : String = '' ; ValeurMarque : String ='' ;
                           OptimiseUpdate : Boolean = FALSE ; OkPourRegul : Boolean = FALSE) : String ;
procedure GetRappro ( Q : TQuery ; X : TL_Rappro ; DEV : RDEVISE ) ;
Function DansSaisie ( XOrig : TL_Rappro ; LSais : TList ) : boolean ;
Function DiffRappro ( X1,X2 : TL_Rappro ) : boolean ;



implementation

Procedure EquilibreDerniere ( TVentAna : TList ; OEcr : TOBM ; RY : R_YV ; DEV : RDevise ; Force : boolean; aDossier : string = '') ;
Var SommeDevise,SommePivot,SommePourc : Double ;
    i : integer ;
    OAna : TOBM ;
    Debit : Boolean ;
    DD,CD,DP,CP : Double ;
BEGIN
SommeDevise:=0 ; SommePivot:=0 ; SommePourc:=0 ;
for i:=0 to TVentAna.Count-1 do
    BEGIN
    OAna:=TOBM(TVentAna[i]) ;
    SommeDevise:=SommeDevise+OAna.GetMvt('Y_DEBITDEV')+OAna.GetMvt('Y_CREDITDEV') ;
    SommePivot:=SommePivot+OAna.GetMvt('Y_DEBIT')+OAna.GetMvt('Y_CREDIT') ;
    SommePourc:=SommePourc+OAna.GetMvt('Y_POURCENTAGE') ;
    END ;
if ((Force) or (Arrondi(SommePourc-100.0,ADecimP)=0)) then
   BEGIN
   {Equilibre sur la dernière}
   OAna:=TOBM(TVentAna[TVentAna.Count-1]) ;
   DD:=OAna.GetMvt('Y_DEBITDEV') ; CD:=OAna.GetMvt('Y_CREDITDEV') ;
   DP:=OAna.GetMvt('Y_DEBIT') ; CP:=OAna.GetMvt('Y_CREDIT') ;
   Debit:=(DD<>0) ;
   if Debit then
      BEGIN
      OAna.PutMvt('Y_DEBIT',DP+RY.TotPivot-SommePivot) ;
      OAna.PutMvt('Y_DEBITDEV',DD+RY.TotDevise-SommeDevise) ;
      OAna.PutMvt('Y_CREDIT',0) ; OAna.PutMvt('Y_CREDITDEV',0) ;
      END else
      BEGIN
      OAna.PutMvt('Y_CREDIT',CP+RY.TotPivot-SommePivot) ;
      OAna.PutMvt('Y_CREDITDEV',CD+RY.TotDevise-SommeDevise) ;
      OAna.PutMvt('Y_DEBIT',0) ; OAna.PutMvt('Y_DEBITDEV',0) ;
      END ;
   END else
   BEGIN
   {Equilibre sur une nouvelle ligne sur section attente}
   OAna:=TOBM.Create(EcrAna,RY.Axe,True) ; EcrToAna(OEcr,OAna) ;
   RY.Section:=GetInfosAxe(ValeurI(Copy(RY.Axe, 2, 1)), aDossier).Attente ; {JP 24/10/07 : Multi sociétés et partage}
   RY.Pourcentage:=Arrondi(100.0-SommePourc,ADecimP) ;
   RY.NumVentil:=TVentAna.Count+1 ;
   OAna.PutMvt('Y_SECTION',RY.Section) ; OAna.PutMvt('Y_POURCENTAGE',RY.Pourcentage) ;
   OAna.PutMvt('Y_NUMVENTIL',RY.NumVentil) ;
   FiniMontantsAna(OAna,RY,DEV,False) ;
   TVentAna.Add(OAna) ;
   Equilibrederniere(TVentAna,OEcr,RY,DEV,True, aDossier) ;
   END ;
END ;           

{JP 24/10/07 : FQ 21714 : gestion du multi sociétés : ajout du paramêtre aDossier}
Procedure VentileGenerale ( Ax : String3 ; OEcr : TOBM ; DEV : RDevise ; TVentAna : TList ; Preventil : boolean; aDossier : string = '' ) ;
Var OAna     : TOBM ;
    NumAxe   : integer ;
    RY       : R_YV ;
    Debit    : boolean ;
    Q        : TQuery ;
    CpteGene,StV : String ;
    NatP         : String3 ;
    TotPivot,TotDevise : Double ;
    InfoAxe : TInfoCpta;
BEGIN
NumAxe:=StrToInt(Copy(Ax,2,1)) ; CpteGene:=OEcr.GetMvt('E_GENERAL') ;
TotPivot:=OEcr.GetMvt('E_DEBIT')+OEcr.GetMvt('E_CREDIT') ; Debit:=(OEcr.GetMvt('E_DEBIT')<>0) ;
NatP:=OEcr.GetMvt('E_NATUREPIECE') ;
if NatP<>'ECC' then TotDevise:=OEcr.GetMvt('E_DEBITDEV')+OEcr.GetMvt('E_CREDITDEV')
               else TotDevise:=TotPivot ;
StV:='SELECT * FROM ' + GetTableDossier(aDossier, 'VENTIL') + ' WHERE V_NATURE="GE'+IntToStr(NumAxe)+'" AND V_COMPTE="'+CpteGene+'" '
    +'ORDER BY V_NATURE, V_COMPTE, V_NUMEROVENTIL' ;
Q:=OpenSQL(StV,True,-1, '', True) ;
if ((Not Q.EOF) and (Preventil)) then
   BEGIN
   While Not Q.EOF do
      BEGIN
      if GetParamSocDossierMS('SO_CROISAXE', False, aDossier) then
      begin
        RY.SousPlan1 := Q.FindField('V_SOUSPLAN1').AsString;
        RY.SousPlan2 := Q.FindField('V_SOUSPLAN2').AsString;
        RY.SousPlan3 := Q.FindField('V_SOUSPLAN3').AsString;
        RY.SousPlan4 := Q.FindField('V_SOUSPLAN4').AsString;
        RY.SousPlan5 := Q.FindField('V_SOUSPLAN5').AsString;
      end
      else
      begin
      end;
      RY.Axe:=Ax ; RY.Section:=Q.FindField('V_SECTION').AsString ;
      RY.Pourcentage:=Q.FindField('V_TAUXMONTANT').AsFloat ;
      RY.NumVentil:=Q.FindField('V_NUMEROVENTIL').AsInteger ;
      RY.TotDevise:=TotDevise ; RY.TotPivot:=TotPivot ; RY.Debit:=Debit ;
      if RY.Section<>'' then
         BEGIN
         OAna:=TOBM.Create(EcrAna,Ax,True) ;
         EcrToAna(OEcr,OAna) ; FiniMontantsAna(OAna,RY,DEV,False, aDossier) ;
         TVentAna.Add(OAna) ;
         END ;
      Q.Next ;
      END ;
   EquilibreDerniere(TVentAna,OEcr,RY,DEV,False, aDossier) ;
   END else
   BEGIN
   {JP 24/10/07 : FQ 21714 : gestion du multi sociétés et du partage du référentiel}
   InfoAxe := GetInfosAxe(NumAxe, aDossier);
   RY.Axe:=Ax ;
   RY.Section := InfoAxe.Attente ;
   RY.NumVentil:=1 ; RY.Pourcentage:=100.0 ;
   RY.TotDevise:=TotDevise ; RY.TotPivot:=TotPivot ; RY.Debit:=Debit ;
   if GetParamSocDossierMS('SO_CROISAXE', False, aDossier) then
   begin
     if GetParamSocDossierMS('SO_VENTILA1', False, aDossier) then begin
       InfoAxe := GetInfosAxe(1, aDossier);
       RY.SousPlan1 := InfoAxe.Attente;
     end
     else
       RY.SousPlan1 := '';

     if GetParamSocDossierMS('SO_VENTILA2', False, aDossier) then begin
       InfoAxe := GetInfosAxe(2, aDossier);
       RY.SousPlan2 := InfoAxe.Attente;
     end
     else
       RY.SousPlan2 := '';

     if GetParamSocDossierMS('SO_VENTILA3', False, aDossier) then begin
       InfoAxe := GetInfosAxe(3, aDossier);
       RY.SousPlan3 := InfoAxe.Attente;
     end
     else
       RY.SousPlan3 := '';

     if GetParamSocDossierMS('SO_VENTILA4', False, aDossier) then begin
       InfoAxe := GetInfosAxe(4, aDossier);
       RY.SousPlan4 := InfoAxe.Attente;
     end
     else
       RY.SousPlan4 := '';

     if GetParamSocDossierMS('SO_VENTILA5', False, aDossier) then begin
       InfoAxe := GetInfosAxe(5, aDossier);
       RY.SousPlan5 := InfoAxe.Attente;
     end
     else
       RY.SousPlan5 := '';
   end;
   if RY.Section<>'' then
      BEGIN
      OAna:=TOBM.Create(EcrAna,Ax,True) ;
      EcrToAna(OEcr,OAna) ; FiniMontantsAna(OAna,RY,DEV,True, aDossier) ;
      TVentAna.Add(OAna) ;
      END ;
   END ;
Ferme(Q) ;
END ;

procedure RemplirOFromM ( O : TOBM ;  M : RMVT ; Totale : boolean ) ;
BEGIN
O.PutMvt('E_JOURNAL',M.Jal)         ; O.PutMvt('E_EXERCICE',M.Exo) ;
O.PutMvt('E_DATECOMPTABLE',M.DateC) ; O.PutMvt('E_NATUREPIECE',M.Nature) ;
O.PutMvt('E_MODESAISIE',M.ModeSaisieJal) ;
O.PutMvt('E_PERIODE',GetPeriode(M.DateC)) ; O.PutMvt('E_SEMAINE',NumSemaine(M.DateC)) ;
O.PutMvt('E_NUMEROPIECE',M.Num)     ; O.PutMvt('E_QUALIFPIECE',M.Simul) ;
O.PutMvt('E_DEVISE',M.CodeD)        ; O.PutMvt('E_DATETAUXDEV',M.DateTaux) ; O.PutMvt('E_TAUXDEV',M.TauxD) ;
O.PutMvt('E_DATEECHEANCE',M.DateC)  ; O.PutMvt('E_ETABLISSEMENT',M.Etabl) ;
if M.Valide then O.PutMvt('E_VALIDE','X') else O.PutMvt('E_VALIDE','-') ;
if Totale then
   BEGIN
   O.PutMvt('E_NUMLIGNE',M.NumLigne)  ; O.PutMvt('E_NUMECHE',M.NumEche ) ;
   END ;
// CEGID V9
O.PutMvt('E_DATEORIGINE', iDate1900) ;
O.PutMvt('E_DATEPER',iDate1900) ;
O.PutMvt('E_ENTITY',0 );
O.PutMvt('E_REFGUID','') ;
END ;

{$IFNDEF SANSCOMPTA}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 18/06/2002
Modifié le ... :   /  /
Description .. : On ne prend que le 4 premiers caracteres du code lettre
Mots clefs ... :
*****************************************************************}
procedure Lett_RempliBase ( TEcr : TQUERY ; OL : TOBM ; Client : Boolean ; Regul : PG_Lettrage ; ForceSaisieEuro : String = '') ;
BEGIN
{$IFDEF EAGLCLIENT}
  TEcr.PutValue('E_REGIMETVA', OL.GetMvt('E_REGIMETVA'));
//  If ForceSaisieEuro<>'' Then TEcr.PutValue('E_SAISIEEURO', ForceSaisieEuro)
//                         Else TEcr.PutValue('E_SAISIEEURO', OL.GetMvt('E_SAISIEEURO'));

  if Client then BEGIN
    TEcr.PutValue('E_LETTRAGE', Copy(OL.GetMvt('E_LETTRAGE'),1,4));
    TEcr.PutValue('E_ETATLETTRAGE', OL.GetMvt('E_ETATLETTRAGE'));
    TEcr.PutValue('E_ETAT', OL.GetMvt('E_ETAT'));
    TEcr.PutValue('E_LETTRAGEDEV', OL.GetMvt('E_LETTRAGEDEV'));
//    TEcr.PutValue('E_LETTRAGEEURO', OL.GetMvt('E_LETTRAGEEURO'));
    TEcr.PutValue('E_MODEPAIE', OL.GetMvt('E_MODEPAIE'));
    END
  else BEGIN
    TEcr.PutValue('E_LETTRAGE', '');
    TEcr.PutValue('E_ETATLETTRAGE', 'RI');
    TEcr.PutValue('E_ETAT', '0000000000');
    TEcr.PutValue('E_LETTRAGEDEV', '-');
//    TEcr.PutValue('E_LETTRAGEEURO', '-');
    TEcr.PutValue('E_MODEPAIE', '');
  END ;
  if Regul=pglRegul   then TEcr.PutValue('E_QUALIFORIGINE', 'REG') else
  if Regul=pglConvert then TEcr.PutValue('E_QUALIFORIGINE', 'CON') else
  if Regul=pglInverse then TEcr.PutValue('E_QUALIFORIGINE', 'CON') else
  if Regul=pglEcart   then TEcr.PutValue('E_QUALIFORIGINE', 'ECC');
  // CEGID V9
  TEcr.PutValue('E_DATEORIGINE'), iDate1900) ;
  TEcr.PutValue('E_DATEPER'),Date1900) ;
  TEcr.PutValue('E_ENTITY'),0 );
  TEcr.PutValue('E_REFGUID'),'') ;
{$ELSE}
  TEcr.FindField('E_REGIMETVA').AsString:=OL.GetMvt('E_REGIMETVA') ;

if Client then
   BEGIN
   TEcr.FindField('E_LETTRAGE').AsString:=Copy(OL.GetMvt('E_LETTRAGE'),1,4) ;
   TEcr.FindField('E_ETATLETTRAGE').AsString:=OL.GetMvt('E_ETATLETTRAGE') ;
   TEcr.FindField('E_ETAT').AsString:=OL.GetMvt('E_ETAT') ;
   TEcr.FindField('E_LETTRAGEDEV').AsString:=OL.GetMvt('E_LETTRAGEDEV') ;
   TEcr.FindField('E_MODEPAIE').AsString:=OL.GetMvt('E_MODEPAIE');
   END else
   BEGIN
   TEcr.FindField('E_LETTRAGE').AsString:='' ;
   TEcr.FindField('E_ETATLETTRAGE').AsString:='RI' ;
   TEcr.FindField('E_ETAT').AsString:='0000000000' ;
   TEcr.FindField('E_LETTRAGEDEV').AsString:='-' ;
   TEcr.FindField('E_MODEPAIE').AsString:='' ;
   END ;
if Regul=pglRegul then TEcr.FindField('E_QUALIFORIGINE').AsString:='REG' else
 if Regul=pglConvert then TEcr.FindField('E_QUALIFORIGINE').AsString:='CON' else
  if Regul=pglInverse then TEcr.FindField('E_QUALIFORIGINE').AsString:='CON' else
   if Regul=pglEcart then TEcr.FindField('E_QUALIFORIGINE').AsString:='ECC' ;
  // CEGID V9
  TEcr.FindField('E_DATEORIGINE').AsDateTime  := iDate1900 ;
  TEcr.FindField('E_DATEPER').AsDateTime  := iDate1900 ;
  TEcr.FindField('E_ENTITY').AsInteger   := 0 ;
  TEcr.FindField('E_REFGUID').AsString   := '' ;
{$ENDIF}
END ;

procedure Lett_NewLigne ( TEcr : TQUERY ; M : RMVT ) ;
Var O : TOBM ;
BEGIN
InitNew(TEcr) ;
O:=TOBM.Create(EcrGen,'',True) ; RemplirOFromM(O,M,False) ;
O.EgalChamps(TEcr) ; O.Free ;
END ;

procedure CalculeMontants( XX : T_ECARTCHANG ; M : RMVT ; Var XP,XD : Double ) ;
BEGIN
XP:=0 ; XD:=0 ;
Case XX.Regul of
   pglRegul : BEGIN
              if M.CodeD<>V_PGI.DevisePivot then
                 BEGIN
                 XD:=Abs(XX.Montant1) ;
                 XP:=DeviseToEuro(XD,M.TauxD,XX.Quotite) ;
                 END else
                 BEGIN
                 if VH^.TenueEuro then
                    BEGIN
                      XP:=Abs(XX.Montant1) ; XD:=XP ;
                    END ;
                 END ;
              END ;
   pglEcart : BEGIN
              XD:=0 ;
              if VH^.TenueEuro then
                 BEGIN
                 XP:=XX.Montant1 ;
                 if XP<0 then XP:=-XP ;
                 END else
                 BEGIN
                  XP:=XX.Montant2 ;
                 END ;
              END ;
 pglConvert : BEGIN
              if VH^.TenueEuro then
                 BEGIN
                 XP:=XX.Montant1 ;
                 if (XP<0) or (XP=0) then BEGIN XP:=-XP ; END ;
                 END else
                 BEGIN
                 XP:=XX.Montant2 ;
                 END ;
              if M.CodeD=V_PGI.DevisePivot then XD:=XP else XD:=0 ;
              END ;
 pglInverse : BEGIN
              if VH^.TenueEuro then
                 BEGIN
                 XP:=XX.Montant1 ;
                 if (XP<0) or (XP=0) then BEGIN XP:=-XP ; END ;
                 END else
                 BEGIN
                 XP:=XX.Montant2 ;
                 END ;
              if M.CodeD=V_PGI.DevisePivot then XD:=XP else XD:=0 ;
              END ;
   END ; (* case *)
END ;

procedure Lett_TraiteMontant ( TEcr : TQUERY ; XX : T_ECARTCHANG ; M : RMVT ; Premier,Gaffe : boolean ; it : integer ) ;
Var XP,XD : Double ;
    Positif  : boolean ;
BEGIN
CalculeMontants(XX,M,XP,XD) ;
if XX.Regul=pglRegul then
   BEGIN
   Positif:=(XX.Montant1>0) ;
   if ((Positif) and (Premier)) or ((Not Positif) and (Not Premier)) then
      BEGIN
      {$IFDEF EAGLCLIENT}
      TEcr.PutValue('E_CREDIT', XP);     TEcr.PutValue('E_DEBIT', 0);
      TEcr.PutValue('E_CREDITDEV', XD);  TEcr.PutValue('E_DEBITDEV', 0);
      {$ELSE}
      TEcr.FindField('E_CREDIT').AsFloat:=XP     ; TEcr.FindField('E_DEBIT').AsFloat:=0 ;
      TEcr.FindField('E_CREDITDEV').AsFloat:=XD  ; TEcr.FindField('E_DEBITDEV').AsFloat:=0 ;
      {$ENDIF}
      END else
      BEGIN
      {$IFDEF EAGLCLIENT}
      TEcr.PutValue('E_CREDIT', 0);     TEcr.PutValue('E_DEBIT', XP);
      TEcr.PutValue('E_CREDITDEV', 0);  TEcr.PutValue('E_DEBITDEV', XD);
      {$ELSE}
      TEcr.FindField('E_CREDIT').AsFloat:=0     ; TEcr.FindField('E_DEBIT').AsFloat:=XP ;
      TEcr.FindField('E_CREDITDEV').AsFloat:=0  ; TEcr.FindField('E_DEBITDEV').AsFloat:=XD ;
      {$ENDIF}
      END ;
   if Premier then
      BEGIN
      {$IFDEF EAGLCLIENT}
      TEcr.PutValue('E_COUVERTURE', XP);
      TEcr.PutValue('E_COUVERTUREDEV', XD);
      {$ELSE}
      TEcr.FindField('E_COUVERTURE').AsFloat:=XP ;
      TEcr.FindField('E_COUVERTUREDEV').AsFloat:=XD ;
      {$ENDIF}
      END ;
   END else if XX.Regul=pglEcart then
   BEGIN
   if Gaffe and (it<>1) then XP:=0 ;
   Positif:=(XX.Montant1>0) ;
   if ((Positif) and (Premier)) or ((Not Positif) and (Not Premier)) then
      BEGIN
      {$IFDEF EAGLCLIENT}
      TEcr.PutValue('E_CREDIT', XP);     TEcr.PutValue('E_DEBIT', 0);
      TEcr.PutValue('E_CREDITDEV', XD);  TEcr.PutValue('E_DEBITDEV', 0);
      {$ELSE}
      TEcr.FindField('E_CREDIT').AsFloat:=XP     ; TEcr.FindField('E_DEBIT').AsFloat:=0 ;
      TEcr.FindField('E_CREDITDEV').AsFloat:=XD  ; TEcr.FindField('E_DEBITDEV').AsFloat:=0 ;
      {$ENDIF}
      END else
      BEGIN
      {$IFDEF EAGLCLIENT}
      TEcr.PutValue('E_CREDIT', 0);     TEcr.PutValue('E_DEBIT', XP);
      TEcr.PutValue('E_CREDITDEV', 0);  TEcr.PutValue('E_DEBITDEV', XD);
      {$ELSE}
      TEcr.FindField('E_CREDIT').AsFloat:=0     ; TEcr.FindField('E_DEBIT').AsFloat:=XP ;
      TEcr.FindField('E_CREDITDEV').AsFloat:=0  ; TEcr.FindField('E_DEBITDEV').AsFloat:=XD ;
      {$ENDIF}
      END ;
   if Premier then
      BEGIN
      {$IFDEF EAGLCLIENT}
      TEcr.PutValue('E_COUVERTURE', XP);
      TEcr.PutValue('E_COUVERTUREDEV', XD);
      {$ELSE}
      TEcr.FindField('E_COUVERTURE').AsFloat:=XP ;
      TEcr.FindField('E_COUVERTUREDEV').AsFloat:=XD ;
      {$ENDIF}
      END ;
   END else if XX.Regul=pglConvert then
   BEGIN
   if Gaffe and (it<>1) then XP:=0 ;
   Positif:=(XX.Montant1>0) or ((XX.Montant1=0) and (XX.Montant2>0)) ;
   if ((Positif) and (Premier)) or ((Not Positif) and (Not Premier)) then
      BEGIN
      {$IFDEF EAGLCLIENT}
      TEcr.PutValue('E_CREDIT', XP);     TEcr.PutValue('E_DEBIT', 0);
      TEcr.PutValue('E_CREDITDEV', XD);  TEcr.PutValue('E_DEBITDEV', 0);
      {$ELSE}
      TEcr.FindField('E_CREDIT').AsFloat:=XP     ; TEcr.FindField('E_DEBIT').AsFloat:=0 ;
      TEcr.FindField('E_CREDITDEV').AsFloat:=XD  ; TEcr.FindField('E_DEBITDEV').AsFloat:=0 ;
      {$ENDIF}
      END else
      BEGIN
      {$IFDEF EAGLCLIENT}
      TEcr.PutValue('E_CREDIT', 0);     TEcr.PutValue('E_DEBIT', XP);
      TEcr.PutValue('E_CREDITDEV', 0);  TEcr.PutValue('E_DEBITDEV', XD);
      {$ELSE}
      TEcr.FindField('E_CREDIT').AsFloat:=0     ; TEcr.FindField('E_DEBIT').AsFloat:=XP ;
      TEcr.FindField('E_CREDITDEV').AsFloat:=0  ; TEcr.FindField('E_DEBITDEV').AsFloat:=XD ;
      {$ENDIF}
      END ;
   if Premier then
      BEGIN
      {$IFDEF EAGLCLIENT}
      TEcr.PutValue('E_COUVERTURE', XP);
      TEcr.PutValue('E_COUVERTUREDEV', XD);
      {$ELSE}
      TEcr.FindField('E_COUVERTURE').AsFloat:=XP ;
      TEcr.FindField('E_COUVERTUREDEV').AsFloat:=XD ;
      {$ENDIF}
      END ;
   END else if XX.Regul=pglInverse then
   BEGIN
   if Gaffe and (it<>1) then XP:=0 ;
   Positif:=(XX.Montant1>0) or ((XX.Montant1=0) and (XX.Montant2>0)) ;
   if ((Positif) and (Premier)) or ((Not Positif) and (Not Premier)) then
      BEGIN
      {$IFDEF EAGLCLIENT}
      TEcr.PutValue('E_CREDIT', XP);     TEcr.PutValue('E_DEBIT', 0);
      TEcr.PutValue('E_CREDITDEV', XD);  TEcr.PutValue('E_DEBITDEV', 0);
      {$ELSE}
      TEcr.FindField('E_CREDIT').AsFloat:=XP     ; TEcr.FindField('E_DEBIT').AsFloat:=0 ;
      TEcr.FindField('E_CREDITDEV').AsFloat:=XD  ; TEcr.FindField('E_DEBITDEV').AsFloat:=0 ;
      {$ENDIF}
      END else
      BEGIN
      {$IFDEF EAGLCLIENT}
      TEcr.PutValue('E_CREDIT', 0);     TEcr.PutValue('E_DEBIT', XP);
      TEcr.PutValue('E_CREDITDEV', 0);  TEcr.PutValue('E_DEBITDEV', XD);
      {$ELSE}
      TEcr.FindField('E_CREDIT').AsFloat:=0     ; TEcr.FindField('E_DEBIT').AsFloat:=XP ;
      TEcr.FindField('E_CREDITDEV').AsFloat:=0  ; TEcr.FindField('E_DEBITDEV').AsFloat:=XD ;
      {$ENDIF}
      END ;
   if Premier then
      BEGIN
      {$IFDEF EAGLCLIENT}
      TEcr.PutValue('E_COUVERTURE', XP);
      TEcr.PutValue('E_COUVERTUREDEV', XD);
      {$ELSE}
      TEcr.FindField('E_COUVERTURE').AsFloat:=XP ;
      TEcr.FindField('E_COUVERTUREDEV').AsFloat:=XD ;
      {$ENDIF}
      END ;
   END ;
END ;

function VentilLaRegul ( TEcr : TDataSet ; DEV : RDEVISE ; RL : RLETTR ; XX : T_ECARTCHANG ) : boolean ;
Var NumAxe,k : integer ;
    Cpte,Ax : String ;
    Q    : TQuery ;
    QAna : TQuery ;
    TVentAna  : TList ;
    OEcr,OAna : TOBM ;
BEGIN
{$IFNDEF PGIIMMO}
Result:=False ;
{$IFDEF EAGLCLIENT}
Cpte:=TEcr.GetValue('E_GENERAL');
{$ELSE}
Cpte:=TEcr.FindField('E_GENERAL').AsString ;
{$ENDIF}
Q:=OpenSQL('Select G_VENTILABLE1, G_VENTILABLE2, G_VENTILABLE3, G_VENTILABLE4, G_VENTILABLE5, G_VENTILABLE '
          +'from GENERAUX Where G_GENERAL="'+Cpte+'"',True,-1, '', True) ;
if Q.EOF then BEGIN Ferme(Q) ; Exit ; END ;
if Q.FindField('G_VENTILABLE').AsString<>'X' then BEGIN Ferme(Q) ; Exit ; END ;
{$IFDEF EAGLCLIENT}
  QAna := TOB.Create('ANALYTIQ',nil,-1);
{$ELSE}
  QAna:=OpenSQL('SELECT * FROM ANALYTIQ WHERE Y_SECTION="Sfdf"',False) ;
{$ENDIF}
TVentAna:=TList.Create ; OEcr:=TOBM.Create(EcrGen,'',True) ;

{$IFDEF EAGLCLIENT}
{ CA - 19/07/2006 - FQ 18562 - Problème génération analytique écarts de changes en CWAS }
//OEcr.ChargeMvt(TEcr) ;
OEcr.Dupliquer(TEcr,False,True,True);
{$ELSE}
OEcr.ChargeMvt(TEcr) ;
{$ENDIF}
for NumAxe:=1 to MaxAxe do
    BEGIN
    if Q.FindField('G_VENTILABLE'+IntToStr(NumAxe)).AsString='X' then
       BEGIN
       Ax:='A'+IntToStr(NumAxe) ; VideListe(TVentAna) ;
       VentileGenerale(Ax,OEcr,DEV,TVentAna,True) ;
       for k:=0 to TVentAna.Count-1 do BEGIN
{$IFDEF EAGLCLIENT}
         OAna:=TOBM(TVentAna[k]) ;
         OAna.EgalChamps(QAna);
         if XX.Regul=pglEcart then ECCANA(QAna,RL.DeviseMvt) ;
         QAna.InsertDB(nil,False);
         MajSoldeSectionTOB(QAna,True);
{$ELSE}
         QAna.Insert ;
         OAna:=TOBM(TVentAna[k]) ;
         OAna.EgalChamps(QAna) ;
         if XX.Regul=pglEcart then ECCANA(QAna,RL.DeviseMvt) ;
         QAna.Post ;
         MajSoldeSection(QAna,True) ;
{$ENDIF}
           END ;
       END ;
    END ;
Ferme(Q) ; Ferme(QAna) ; VideListe(TVentAna) ; TVentAna.Free ; OEcr.Free ;
Result:=True ;
{$ENDIF}
END ;

Procedure RensTvaEnc ( TEcr : TDataSet ) ;
Var
  QQ : TQuery ;
  szGeneral : String;
BEGIN
  if Not VH^.OuiTvaEnc then Exit ;
{$IFDEF EAGLCLIENT}
  szGeneral := TEcr.GetValue('E_GENERAL');
{$ELSE}
  szGeneral := TEcr.FindField('E_GENERAL').AsString;
{$ENDIF}
  QQ:=OpenSQL('Select G_TVASURENCAISS, G_TVA from GENERAUX Where G_GENERAL="' + szGeneral + '" '
             +'AND (G_NATUREGENE="CHA" or G_NATUREGENE="PRO" or G_NATUREGENE="IMO")',True,-1, '', True) ;
  if Not QQ.EOF then BEGIN
  {$IFDEF EAGLCLIENT}
   TEcr.PutValue('E_TVAENCAISSEMENT', QQ.Fields[0].AsString);
   TEcr.PutValue('E_TVA', QQ.Fields[1].AsString);
  {$ELSE}
   TEcr.FindField('E_TVAENCAISSEMENT').AsString:=QQ.Fields[0].AsString ;
   TEcr.FindField('E_TVA').AsString:=QQ.Fields[1].AsString ;
  {$ENDIF}
   END ;
Ferme(QQ) ;
END ;

function Signe ( X : Double ) : integer ;
BEGIN
if X=Abs(X) then Signe:=1 else Signe:=-1 ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 11/06/2002
Modifié le ... : 23/04/2007
Description .. : -10/06/2002 - affectation des champs e_paquetrevision et
Suite ........ : e_io
Suite ........ : - 01/04/2004 - suppression de SetMPACCDB
Suite ........ : - LG - 08/12/2006 - FB 19052 - coorection du sesn des 
Suite ........ : ecriterus d'ecart de change en eAgl
Suite ........ : - LG - 23/04/2007 - FB 19567 - on lettrait la 2eme ligne sur 
Suite ........ : le compte de contrepartie
Mots clefs ... : 
*****************************************************************}
function  CreerPartieDoubleLett ( M : RMVT ; RL : RLETTR ; XX : T_ECARTCHANG ; OL : TOBM ; ForceSaisieEuro : String = '') : TOBM ;
Var
  {$IFDEF EAGLCLIENT}
    TEcrM, TEcr  : TOB ;
  {$ELSE}
    TEcr : TQuery ;
  {$ENDIF}
    it,NumE,NbTour : integer ;
    DEV  : RDEVISE ;
    OOO  : TOBM ;
    OKV1,OkV2,Gaffe  : boolean ;
BEGIN
Result:=Nil ; OkV1:=False ; OkV2:=False ; NbTour:=1 ; Gaffe:=False ;
{$IFNDEF EAGLCLIENT}
  TEcr:=OpenSQL('SELECT * FROM ECRITURE WHERE E_GENERAL="Sfdf"',False) ;
{$ENDIF}
if ((Signe(XX.Montant1)<>Signe(XX.Montant2)) and (XX.Montant1<>0) and (XX.Montant2<>0)) then BEGIN NbTour:=2 ; Gaffe:=True ; END ;
for it:=1 to NbTour do
    BEGIN
  {première ligne sur compte client}
{$IFDEF EAGLCLIENT}
  TEcrM:=TOB.Create('ECRITURE',nil,-1) ;
  TEcr :=TOB.Create('ECRITURE',TEcrM,-1) ;

//  RemplirOFromM(TOBM(TEcr),M,False) ;
  Lett_NewLigne(TEcr,M) ;
  Lett_RempliBase(TEcr,OL,True,XX.Regul,ForceSaisieEuro) ;
  Lett_TraiteMontant(TEcr,XX,M,True,Gaffe,it) ;

  TEcr.PutValue('E_GENERAL', RL.General);
  TEcr.PutValue('E_AUXILIAIRE', RL.Auxiliaire);
  TEcr.PutValue('E_CONTREPARTIEGEN', XX.Cpte);
  TEcr.PutValue('E_CONTREPARTIEAUX', '');
  TEcr.PutValue('E_REFINTERNE', XX.Ref);
  TEcr.PutValue('E_LIBELLE', XX.Lib);
  TEcr.PutValue('E_REFLETTRAGE', XX.RefLettrage);
  TEcr.PutValue('E_DATEPAQUETMIN', XX.DPMin);
  TEcr.PutValue('E_DATEPAQUETMAX', XX.DPMax);
  TEcr.PutValue('E_ECHE', OL.GetMvt('E_ECHE'));
  TEcr.PutValue('E_NUMLIGNE', 2*it-1);
  NumE:=OL.GetMvt('E_NUMECHE');
  if NumE>0 then TEcr.PutValue('E_NUMECHE', 1);
  TEcr.PutValue('E_TYPEMVT', 'DIV');
  TEcr.PutValue('E_IO', 'X'); // ecriture a synchronise
  TEcr.PutValue('E_PAQUETREVISION', 1); // verrous s5 sur un ecriture lettre
  TEcr.PutValue('E_CODEACCEPT', MPTOACC( TEcr.GetValue('E_MODEPAIE') ) ) ;
  // CEGID V9
  TEcr.PutValue('E_DATEORIGINE'), iDate1900 °;
  TEcr.PutValue('E_DATEPER'),iDate1900) ;
  TEcr.PutValue('E_ENTITY'),0 );
  TEcr.PutValue('E_REFGUID'),'') ;
{$ELSE}
  TEcr.Insert ;
    Lett_NewLigne(TEcr,M) ; Lett_RempliBase(TEcr,OL,True,XX.Regul,ForceSaisieEuro) ;
    Lett_TraiteMontant(TEcr,XX,M,True,Gaffe,it) ;
    TEcr.FindField('E_GENERAL').AsString:=RL.General ;
    TEcr.FindField('E_AUXILIAIRE').AsString:=RL.Auxiliaire ;
    TEcr.FindField('E_CONTREPARTIEGEN').AsString:=XX.Cpte ;
    TEcr.FindField('E_CONTREPARTIEAUX').AsString:='' ;
    TEcr.FindField('E_REFINTERNE').AsString:=XX.Ref ;
    TEcr.FindField('E_LIBELLE').AsString:=XX.Lib ;
    TEcr.FindField('E_REFLETTRAGE').AsString:=XX.RefLettrage ;
    TEcr.FindField('E_DATEPAQUETMIN').AsDateTime:=XX.DPMin ;
    TEcr.FindField('E_DATEPAQUETMAX').AsDateTime:=XX.DPMax ;
    TEcr.FindField('E_ECHE').AsString:=OL.GetMvt('E_ECHE') ;
    TEcr.FindField('E_NUMLIGNE').AsInteger:=2*it-1 ;
    NumE:=OL.GetMvt('E_NUMECHE') ;
    if NumE>0 then TEcr.FindField('E_NUMECHE').AsInteger:=1 ;
    TEcr.FindField('E_TYPEMVT').AsString:='DIV' ;
    TEcr.FindField('E_IO').AsString:='X' ; // ecriture a synchronise
    TEcr.FindField('E_PAQUETREVISION').asInteger:=1 ; // verrous s5 sur un ecriture lettre
    TEcr.FindField('E_CODEACCEPT').AsString := MPTOACC(TEcr.FindField('E_MODEPAIE').AsString) ;
    // CEGID V9
    TEcr.FindField('E_DATEORIGINE').AsDateTime  := iDate1900 ;
    TEcr.FindField('E_DATEPER').AsDateTime  := iDate1900 ;
    TEcr.FindField('E_ENTITY').AsInteger   := 0 ;
    TEcr.FindField('E_REFGUID').AsString   := '' ;
{$ENDIF}
    SetCotationDB(0,TEcr) ;
{$IFNDEF PGIIMMO}
    if XX.Regul=pglEcart then ECCECR(TEcr,RL.DeviseMvt) ;
{$ENDIF}
  {$IFNDEF EAGLSERVER}
    {$IFDEF EAGLCLIENT}
      MajSoldesEcritureTOB(TEcr,false) ;
      TEcr.InsertDB(nil) ;
    {$ELSE}
      MajSoldeCompte(TEcr) ;
      TEcr.Post ;
    {$ENDIF}
  {$ENDIF EAGLSERVER}

  OOO:=TOBM.Create(EcrGen,'',True) ;
{$IFDEF EAGLCLIENT}
  OOO.Dupliquer(TEcr,False,True,True);
{$ELSE}
  OOO.ChargeMvt(TEcr);
{$ENDIF}
  Result:=OOO ;
    {Ventilations analytiques}
    FillChar(DEV,Sizeof(DEV),#0) ;
    DEV.Code:=M.CodeD ; DEV.Decimale:=XX.Decimale ; DEV.DateTaux:=M.DateTaux ;
    DEV.Taux:=M.TauxD ; DEV.Quotite:=XX.Quotite ;
    if VentilLaRegul(TEcr,DEV,RL,XX) then OkV1:=True ;

  {deuxième ligne sur compte regul/écart de change}
{$IFDEF EAGLCLIENT}
  TEcrM.Free;
//   TEcr.Free; Inutile car libéré par le TEcrM.Free
  TEcrM:=TOB.Create('ECRITURE',nil,-1) ;
  TEcr :=TOB.Create('ECRITURE',TEcrM,-1) ;

//  RemplirOFromM(TOBM(TEcr),M,False) ;
  Lett_NewLigne(TEcr,M) ;
  Lett_RempliBase(TEcr,OL,False,XX.Regul,ForceSaisieEuro) ;  // LG 23/04/2007  true -> false pour le param demandant de lettrer la ligne
  Lett_TraiteMontant(TEcr,XX,M,false,Gaffe,it) ;
  TEcr.PutValue('E_GENERAL', XX.Cpte);
  TEcr.PutValue('E_AUXILIAIRE', '');
  TEcr.PutValue('E_CONTREPARTIEGEN', RL.General);
  TEcr.PutValue('E_CONTREPARTIEAUX', RL.Auxiliaire);
  TEcr.PutValue('E_REFINTERNE', XX.Ref);
  TEcr.PutValue('E_LIBELLE', XX.Lib);
  TEcr.PutValue('E_REFLETTRAGE', '');
  TEcr.PutValue('E_ECHE', '-');
  TEcr.PutValue('E_NUMECHE', 0);
  TEcr.PutValue('E_NUMLIGNE', 2*it);
  TEcr.PutValue('E_TYPEMVT', 'DIV');
  TEcr.PutValue('E_IO', 'X');
  TEcr.PutValue('E_CODEACCEPT', MPTOACC( TEcr.GetValue('E_MODEPAIE') ) ) ;
  // CEGID V9
  TEcr.PutValue('E_DATEORIGINE'), iDate1900 °;
  TEcr.PutValue('E_DATEPER'),iDate1900) ;
  TEcr.PutValue('E_ENTITY'),0 );
  TEcr.PutValue('E_REFGUID'),'') ;
{$ELSE}
  TEcr.Insert ;
    Lett_NewLigne(TEcr,M) ;
    Lett_RempliBase(TEcr,OL,False,XX.Regul,ForceSaisieEuro) ;
    Lett_TraiteMontant(TEcr,XX,M,False,Gaffe,it) ;
    TEcr.FindField('E_GENERAL').AsString:=XX.Cpte ;
    TEcr.FindField('E_AUXILIAIRE').AsString:='' ;
    TEcr.FindField('E_CONTREPARTIEGEN').AsString:=RL.General ;
    TEcr.FindField('E_CONTREPARTIEAUX').AsString:=RL.Auxiliaire ;
    TEcr.FindField('E_REFINTERNE').AsString:=XX.Ref ;
    TEcr.FindField('E_LIBELLE').AsString:=XX.Lib ;
    TEcr.FindField('E_REFLETTRAGE').AsString:='' ;
    TEcr.FindField('E_ECHE').AsString:='-' ;
    TEcr.FindField('E_NUMECHE').AsInteger:=0 ;
    TEcr.FindField('E_NUMLIGNE').AsInteger:=2*it ;
    TEcr.FindField('E_TYPEMVT').AsString:='DIV' ;
    TEcr.FindField('E_IO').AsString:='X' ;
    TEcr.FindField('E_CODEACCEPT').AsString := MPTOACC(TEcr.FindField('E_MODEPAIE').AsString) ;
    // CEGID V9
    TEcr.FindField('E_DATEORIGINE').AsDateTime  := iDate1900 ;
    TEcr.FindField('E_DATEPER').AsDateTime  := iDate1900 ;
    TEcr.FindField('E_ENTITY').AsInteger   := 0 ;
    TEcr.FindField('E_REFGUID').AsString   := '' ;
{$ENDIF}
{$IFNDEF PGIIMMO}
    if XX.Regul=pglEcart then ECCECR(TEcr,RL.DeviseMvt) ;
{$ENDIF}
    {#TVAENC}
    if ((VH^.OuiTvaEnc) and (XX.Regul=pglRegul)) then RensTvaEnc(TEcr) ;
    SetCotationDB(0,TEcr) ;
{$IFNDEF EAGLSERVER}
{$IFDEF EAGLCLIENT}
  MajSoldesEcritureTOB(TEcr,false) ;
  TEcr.InsertDB(nil);
{$ELSE}
  MajSoldeCompte(TEcr) ;
  TEcr.Post ;
{$ENDIF}
{$ENDIF EAGLSERVER}
    {Ventilations analytiques}
    FillChar(DEV,Sizeof(DEV),#0) ;
    DEV.Code:=M.CodeD ; DEV.Decimale:=XX.Decimale ; DEV.DateTaux:=M.DateTaux ;
    DEV.Taux:=M.TauxD ; DEV.Quotite:=XX.Quotite ;
    if VentilLaRegul(TEcr,DEV,RL,XX) then OkV2:=True ;
  {$IFDEF EAGLCLIENT}
  TEcrM.Free;
//  TEcr.Free; Inutile car libéré par le TEcrM.Free
  {$ENDIF}
END ;
{$IFNDEF EAGLCLIENT}
  Ferme(TEcr) ;
{$ENDIF}
ADevalider(M.Jal,M.DateC) ;
{$IFDEF EAGLCLIENT}
  // A FAIRE
{$ELSE}
{$IFNDEF EAGLSERVER}
  MarquerPublifi(True) ;
{$ENDIF EAGLSERVER}
{$ENDIF}
if OkV1 then ExecuteSQL('UPDATE ECRITURE SET E_ANA="X" Where '+WhereEcriture(tsGene,M,False)+' AND (E_NUMLIGNE=1 OR E_NUMLIGNE=3)') ;
if OkV2 then ExecuteSQL('UPDATE ECRITURE SET E_ANA="X" Where '+WhereEcriture(tsGene,M,False)+' AND (E_NUMLIGNE=2 OR E_NUMLIGNE=4)') ;
END ;
{$ENDIF}

{==================== Lettrage en importation =============================}

Procedure CalculCouv ( X : TL_Rappro ; DEV : RDEVISE ; Total : boolean ; Var CP,CD : Double ) ;
BEGIN
if Total then
   BEGIN
   CP:=X.Debit+X.Credit ; CD:=X.DebDev+X.CredDev ;
   END else
   BEGIN
    if DEV.Code<>V_PGI.DevisePivot then
      BEGIN
      CD:=X.CouvCur ;
      CP:=DeviseToEuro(CD,X.TauxDev,DEV.Quotite) ;
      END else
      BEGIN
      CP:=X.CouvCur ; CD:=CP ;
      END ;
   END ;
END ;

procedure FaireLettrage (LEcr : TList ; Total : boolean ; DMin,DMax : TDateTime ; Var CodeL : String ; DEV : RDEVISE ; EncaDeca : boolean ; CurTRACHQ : String = '' ;
                         ChampMarque : String = '' ; ValeurMarque : String ='' ; OptimiseUpdate : Boolean = FALSE) ;
Var i : integer ;
    X : TL_Rappro ;
    NowFutur : TDateTime ;
    CD,CP : Double ;
    SQL,EL,LDEV,STE,StLTE : String ;
    OkPourUpdate : Boolean ;
    lStTable : String ;
BEGIN
if Total then BEGIN EL:='TL' ; CodeL:=uppercase(CodeL) ; END else BEGIN EL:='PL' ; CodeL:=LOWERCASE(CodeL) ; END ;
StE:='---0AM0000' ; StLTE:='-' ;
if Not EncaDeca then CurTRACHQ:='' ;
if DEV.Code<>V_PGI.DevisePivot then
  BEGIN
  if Total then StE[3]:='X' ;
  LDEV:='X' ;
  END else LDEV:='-' ;
NowFutur:=NowH ;
for i:=0 to LEcr.Count-1 do
    BEGIN
    X:=TL_Rappro(LEcr[i]) ;
    CalculCouv(X,DEV,Total,CP,CD) ;
    if EstMultiSoc
      then lStTable := GetTableDossier( X.Dossier, 'ECRITURE' )
      else lStTable := 'ECRITURE' ;
    SQL:='UPDATE ' + lStTable + ' SET E_LETTRAGE="'+CodeL+'", E_ETATLETTRAGE="'+EL+'", E_DATEPAQUETMIN="'+USDATETIME(DMin)+'", '
        +'E_DATEPAQUETMAX="'+USDATETIME(DMax)+'", E_COUVERTURE='+StrfPoint(CP)+', E_COUVERTUREDEV='+StrfPoint(CD)+', '
        +'E_LETTRAGEDEV="'+LDEV+'", E_ETAT="'+StE+'", E_DATEMODIF="'+USTime(NowFutur)+'" ';

    // BPY le 20/08/2004 =< Fiche n° 11941 : remise a zero du champs E_SAISIMP apres un reglement partiel !
    SQL := SQL + ', E_SAISIMP=0';
    // Fin BPY

    // GG COM
    SQL:=SQL+', E_PAQUETREVISION=1 ' ;
    OkPourUpdate:=TRUE ;
    If OptimiseUpdate And (Arrondi(CP,V_PGI.OkDECV)=0) Then OkPourUpdate:=FALSE ;
    if EncaDeca then
       BEGIN
       SQL:=SQL+', E_SUIVDEC=""' ;
       if CurTRACHQ<>'' then SQL:=SQL+', E_NUMTRAITECHQ="'+CurTRACHQ+'" ' ;
       END ;
    If (ChampMarque<>'') And (ValeurMarque<>'') Then
      BEGIN
      SQL:=SQL+', '+CHAMPMARQUE+'="'+VALEURMARQUE+'" ' ;
      END ;
    {#TVAENC}
    if VH^.OuiTvaEnc then
       BEGIN
       if EstCollFact(X.General) then
          BEGIN
          {$IFNDEF EAGLSERVER}
          if RazTvaEnc(X) then SQL:=SQL+', E_ECHEENC1=0, E_ECHEENC2=0, E_ECHEENC3=0, E_ECHEENC4=0, E_ECHEDEBIT=0 ' ;
          {$ENDIF EAGLSERVER}
          END ;
       END ;

    {JP 26/04/04 : pour l'échéancier de la Tréso}
    SQL := SQL + ', E_TRESOSYNCHRO = "LET" ';

    SQL:=SQL+'WHERE E_GENERAL="'+X.General+'" AND E_AUXILIAIRE="'+X.Auxiliaire+'" AND E_EXERCICE="'+X.Exo+'" '
            +'AND E_JOURNAL="'+X.Jal+'" AND E_DATECOMPTABLE="'+USDATETIME(X.DateC)+'" AND E_NUMEROPIECE='+IntToStr(X.Numero)+' '
            +'AND E_NUMLIGNE='+IntToStr(X.NumLigne)+' AND E_NUMECHE='+IntToStr(X.NumEche) ;
    if EncaDeca then SQL:=SQL+' AND E_ETATLETTRAGE<>"TL" ' ;
    If OkPourUpdate Then
      BEGIN
      if ExecuteSQL(SQL)<>1 then BEGIN V_PGI.IoError:=oeUnknown ; Exit ; END ;
      END ;
    END ;
END ;

Function RegulAFaire(Var Dev : RDevise ; OkDev,LettrageOppose : Boolean ; EcartRegulDevise : Double) : Boolean ;
Var DevF : RDevise ;
    EqRegFranc,EqRegEuro,LimiteF,LimiteE,SeuilD,SeuilC : double ;
BEGIN
Result:=FALSE ;
SeuilD:=DEV.MaxDebit ; SeuilC:=DEV.MaxCredit ;
if ((VH^.TenueEuro) and (LettrageOppose)) then
   BEGIN
   DEVF.Code:=V_PGI.DeviseFongible ; GetInfosDevise(DEVF) ;
   END else
   BEGIN
   DEVF:=DEV ;
   END ;
if ((DEV.Code=V_PGI.DevisePivot) and (LettrageOppose)) then
   BEGIN
   if VH^.TenueEuro then BEGIN SeuilD:=DEVF.MaxDebit ; SeuilC:=DEVF.MaxCredit ; END
                    else BEGIN SeuilD:=PivotToEuro(DEV.MaxDebit) ; SeuilC:=PivotToEuro(DEV.MaxCredit) ; END ;
   END ;
if EcartRegulDevise>0 then BEGIN if EcartRegulDevise>SeuilD then Exit; END ;
if EcartRegulDevise<0 then BEGIN if Abs(EcartRegulDevise)>SeuilC then Exit; END ;
if Not VH^.TenueEuro then
   BEGIN
   EqRegFranc:=DeviseToPivot(EcartRegulDevise,DEV.Taux,DEV.Quotite) ;
   EqRegEuro:=DeviseToEuro(EcartRegulDevise,DEV.Taux,DEV.Quotite) ;
   END else
   BEGIN
   EqRegEuro:=DeviseToPivot(EcartRegulDevise,DEV.Taux,DEV.Quotite) ;
   EqRegFranc:=DeviseToFranc(EcartRegulDevise,DEV.Taux,DEV.Quotite) ;
   END ;
if VH^.TenueEuro then
   BEGIN
   LimiteF:=Resolution(V_PGI.OkDecE) ; LimiteE:=Resolution(V_PGI.OkDecV) ;
   END else
   BEGIN
   LimiteF:=Resolution(V_PGI.OkDecV) ; LimiteE:=Resolution(V_PGI.OkDecE) ;
   END ;
// if DEV.Code<>V_PGI.DevisePivot then BEGIN LimiteF:=2*LimiteF ; LimiteE:=2*LimiteE ; END ;
if EcartRegulDevise<>0 then
   BEGIN
   if VH^.TenueEuro then
      BEGIN
      if Abs(EqRegEuro)<LimiteE then BEGIN (*HLettre.Execute(25,'','') ;*) Exit ; END ;  // Montant pivot nul ???
      END else
      BEGIN
      if Abs(EqRegFranc)<LimiteF then BEGIN (*HLettre.Execute(25,'','') ;*) Exit ; END ; // Montant pivot nul ???
      END ;
   END ;
//LettrageTotCur:=True ;
Result:=True ;
END ;

Function LettrerUnPaquet ( LEcr : TList ; Vide,EncaDeca : boolean ;
                           ChampMarque : String = '' ; ValeurMarque : String ='' ;
                           OptimiseUpdate : Boolean = FALSE ; OkPourRegul : Boolean = FALSE) : String ;
Var i : integer ;
    TL : TL_Rappro ;
    LD,LC : TList ;
    DMin,DMax : TDateTime ;
    TotD,TotC,TotDDev,TotCDev,XD,XC : double ;
    CodeL,CurTraCHQ : String ;
    Total:boolean ;
    OkDev,OkTraChq : boolean ;
    DEV : RDEVISE ;
    AvecRegul : Boolean ;
    Solde : Double ;
BEGIN
Result:='' ;
if LEcr.Count=0 then Exit ;
DMin:=IDate2099 ; DMax:=IDate1900 ;
TotD:=0 ; TotC:=0 ; TotDDev:=0 ; TotCDev:=0 ;
CurTRACHQ:='' ; OkTraCHQ:=True ;
OkDev:=True ; //if EncaDeca then LettEuro:=True else LettEuro:=False ;
for i:=0 to LEcr.Count-1 do
    BEGIN
    TL:=TL_Rappro(LEcr[i]) ;
    if i=0 then
       BEGIN
       DEV.Code:=TL.CodeD ; GetInfosDevise(DEV) ; Dev.Taux:=TL.TauxDev ;
       if DEV.Code=V_PGI.DevisePivot then OkDev:=False ;
       END else
       BEGIN
       if DEV.Code<>TL.CodeD then OkDev:=False ;
       END ;
    if TL.DateC>DMax then DMax:=TL.DateC ;
    if TL.DateC<DMin then DMin:=TL.DateC ;
    TotD:=TotD+TL.Debit ; TotC:=TotC+TL.Credit ;
    TotDDev:=TotDDev+TL.DebDev ; TotCDev:=TotCDev+TL.CredDev ;
    if TL.CodeL<>'' then CodeL:=TL.CodeL ;
  //  if ((Not TL.SaisieEuro) or (DEV.Code<>V_PGI.DevisePivot)) then LettEuro:=False ;
    if CurTRACHQ='' then CurTraCHQ:=TL.NumTraCHQ else
       BEGIN
       if ((CurTRACHQ<>TL.NumTraCHQ) and (TL.NumTraCHQ<>'')) then OkTRACHQ:=False ;
       END ;
    END ;
if Not OKTRACHQ then CurTRACHQ:='' ;
if Not OkDev then
   BEGIN
   DEV.Code:=V_PGI.DevisePivot ; GetInfosDevise(DEV) ;
   END ;
if EncaDeca then
   BEGIN
   if Not OkDev then BEGIN TotD:=TotDDev ; TotC:=TotCDev ; END ;
   for i:=0 to LEcr.Count-1 do
       BEGIN
       TL:=TL_Rappro(LEcr[i]) ;
       XD:=TL.DebDev ; XC:=TL.CredDev ;
       TL.DebDev:=TL.Debit ; TL.CredDev:=TL.Credit ;
       TL.Debit:=XD ; TL.Credit:=XC ;
       if OkDev then
          BEGIN
          TL.DebitCur:=TL.DebDev ; TL.CreditCur:=TL.CredDev ;
          END ;
       END ;
   END ;
//Lettrage déja effectué -> même code
{$IFNDEF EAGLSERVER}
if CodeL='' then CodeL:=GetSetCodeLettre(TL_Rappro(LEcr[0]).General,TL_Rappro(LEcr[0]).Auxiliaire) ;
{$ENDIF EAGLSERVER}
LD:=TList.Create ; LC:=TList.Create ;
if ((OkDev) and (EncaDeca)) then
   BEGIN
   Solde:=Arrondi(TotD-TotC,DEV.Decimale) ;
   Total:=(Arrondi(TotD-TotC,DEV.Decimale)=0) ;
   END else
   BEGIN
     Solde:=Arrondi(TotD-TotC,DEV.Decimale) ;
     Total:=(Arrondi(TotD-TotC,DEV.Decimale)=0) ;
   END ;
If (Not Total) And OkPourRegul Then
  BEGIN
  Result:='' ;
  AvecRegul:=RegulAFaire(Dev,OkDev,false,Solde) ;
  If AvecRegul Then Result:='X;' ;
  END ;
for i:=0 to LEcr.Count-1 do
  BEGIN
  TL:=TL_Rappro(LEcr[i]) ;
  if Total then
     BEGIN
      if TL.Debit<>0 then TL.Couv:=TL.Debit else TL.Couv:=TL.Credit ;
     END else
     BEGIN
     if TL.Debit>0 then LD.Add(LEcr[i]) else LC.Add(LEcr[i]) ;
     END ;
  END ;
{$IFNDEF EAGLSERVER}
if Not Total then RefEpuise(LD,LC) ;
{$ENDIF !EAGLSERVER}
FaireLettrage(LEcr,Total,DMin,DMax,CodeL,DEV,EncaDeca,CurTRACHQ,ChampMarque,ValeurMarque,OptimiseUpdate) ;
if Vide then VideListe(LD) ; LD.Clear ; LD.Free ;
if Vide then VideListe(LC) ; LC.Clear ; LC.Free ;
If Total Then CodeL:=CodeL+'T' Else CodeL:=CodeL+'P' ;
If Result='' Then Result:=CodeL Else
  If Result='X;' Then Result:=Result+CodeL+';';
END ;

procedure GetRappro ( Q : TQuery ; X : TL_Rappro ; DEV : RDEVISE ) ;
BEGIN
X.General:=Q.FindField('E_GENERAL').AsString      ; X.Auxiliaire:=Q.FindField('E_AUXILIAIRE').AsString ;
X.Exo:=Q.FindField('E_EXERCICE').AsString         ; X.DateC:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
X.DateE:=Q.FindField('E_DATEECHEANCE').AsDateTime ; X.DateR:=Q.FindField('E_DATEECHEANCE').AsDateTime ; {JLD}
X.RefI:=Q.FindField('E_REFINTERNE').AsString      ; X.RefL:=Q.FindField('E_REFLIBRE').AsString ;
X.RefE:=Q.FindField('E_REFEXTERNE').AsString      ; X.Lib:=Q.FindField('E_LIBELLE').AsString ;
X.Jal:=Q.FindField('E_JOURNAL').AsString          ; X.Numero:=Q.FindField('E_NUMEROPIECE').AsInteger ;
X.NumLigne:=Q.FindField('E_NUMLIGNE').AsInteger   ; X.NumEche:=Q.FindField('E_NUMECHE').AsInteger ;
X.CodeD:=Q.FindField('E_DEVISE').AsString         ; X.TauxDev:=Q.FindField('E_TAUXDEV').AsFloat ;
X.CodeL:=Q.FindField('E_LETTRAGE').AsString       ; X.Decim:=DEV.Decimale ;
{Attention, les 2 lignes après ne sont pas une erreur}
X.Debit:=Q.FindField('E_DEBITDEV').AsFloat        ; X.Credit:=Q.FindField('E_CREDITDEV').AsFloat ;
X.DebDev:=Q.FindField('E_DEBIT').AsFloat          ; X.CredDev:=Q.FindField('E_CREDIT').AsFloat ;
X.DebitCur:=Q.FindField('E_DEBIT').AsFloat        ; X.CreditCur:=Q.FindField('E_CREDIT').AsFloat ;
X.Nature:=Q.FindField('E_NATUREPIECE').AsString ;
X.NumTraCHQ:=Q.FindField('E_NUMTRAITECHQ').AsString ;
X.EditeEtatTva:=(Q.FindField('E_EDITEETATTVA').AsString='X') ;
X.Facture:=((X.Nature='FC') or (X.Nature='FF') or (X.Nature='AC') or (X.Nature='AF')) ;
X.Client:=((X.Nature='FC') or (X.Nature='AC') or (X.Nature='RC') or (X.Nature='OC')) ;
X.Solution:=0 ;
END ;

Function DansSaisie ( XOrig : TL_Rappro ; LSais : TList ) : boolean ;
Var Find : boolean ;
    i    : integer ;
    XSais : TL_Rappro ;
BEGIN
Find:=False ;
for i:=0 to LSais.Count-1 do
    BEGIN
    XSais:=TL_Rappro(LSais[i]) ;
    if ((XOrig.General=XSais.General) and (XOrig.Auxiliaire=XSais.Auxiliaire)) then BEGIN Find:=True ; Break ; END ;
    END ;
Result:=Find ;
END ;

Function DiffRappro ( X1,X2 : TL_Rappro ) : boolean ;
BEGIN
Result:=True ;
if X1.General<>X2.General then Exit ; if X1.Auxiliaire<>X2.Auxiliaire then Exit ;
if X1.DateC<>X2.DateC then Exit     ; if X1.DateE<>X2.DateE then Exit ;
if X1.Jal<>X2.Jal then Exit         ; if X1.Nature<>X2.Nature then Exit ;
if X1.Numero<>X2.Numero then Exit   ; if X1.NumLigne<>X2.NumLigne then Exit ;
if X1.NumEche<>X2.NumEche then Exit ;
Result:=False ;
END ;

{$IFNDEF EAGLCLIENT}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure MajSoldeCompte ( TEcr : TDataSet ) ;
Var TE    : TTypeExo ;
    DateC : TDateTime ;
    Num   : Longint ;
    Lig   : integer ;
    CG,CA,Jal : String ;
    XD,XC : Double ;
    FRM   : TFRM ;
    ll    : LongInt ;
BEGIN
{capture des infos en vue des MAJ}
{$IFNDEF PGIIMMO}
  DateC := TEcr.FindField('E_DATECOMPTABLE').AsDateTime ;
  Num   := TEcr.FindField('E_NUMEROPIECE').AsInteger ;
  Lig   := TEcr.FindField('E_NUMLIGNE').AsInteger ;
  CG    := TEcr.FindField('E_GENERAL').AsString ;
  CA    := TEcr.FindField('E_AUXILIAIRE').AsString ;
  XD    := TEcr.FindField('E_DEBIT').AsFloat ;
  XC    := TEcr.FindField('E_CREDIT').AsFloat ;
  Jal   := TEcr.FindField('E_JOURNAL').AsString ;

  TE := CGetTypeExo( DateC ) ;

  // Général MAJ
  FRM.Cpt   := CG ;
  FRM.NumD  := Num ;
  FRM.DateD := DateC ;
  FRM.LigD  := Lig ;
  AttribParamsNew ( FRM, XD, XC, TE ) ;
  ll := ExecReqMAJ ( fbGene, False, False, FRM ) ;
  If ll<>1 then V_PGI.IoError:=oeSaisie ;

  // Auxi MAJ
  if CA<>'' then
    BEGIN
    FRM.Cpt   := CA ;
    FRM.NumD  := Num ;
    FRM.DateD := DateC ;
    FRM.LigD  := Lig ;
    AttribParamsNew ( FRM, XD, XC, TE ) ;
    ll := ExecReqMAJ ( fbAux, False, False, FRM ) ;
    If ll<>1 then V_PGI.IoError:=oeSaisie ;
    END ;

   // Journal MAJ
   FRM.Cpt   := Jal ;
   FRM.NumD  := Num ;
   FRM.DateD := DateC ;
   FRM.LigD  := Lig ;
   AttribParamsNew ( FRM, XD, XC, TE ) ;
   ll := ExecReqMAJ ( fbJal, False, False, FRM ) ;
   If ll<>1 then V_PGI.IoError:=oeSaisie ;
{$ENDIF}
END ;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF SANSCOMPTA}
procedure MajSoldeSection ( TAna : TDataSet ; Plus : boolean ) ;
Var XD,XC : double ;
    DateC : TDateTime ;
    GeneTypeExo : TTypeExo ;
    FRM   : TFRM ;
    ll    : LongInt ;
BEGIN
{$IFNDEF PGIIMMO}
  XD    := TAna.FindField('Y_DEBIT').AsFloat ;
  XC    := TAna.FindField('Y_CREDIT').AsFloat ;
  DateC := TAna.FindField('Y_DATECOMPTABLE').AsDateTime ;
  if Not Plus then BEGIN XD:=-XD ; XC:=-XC ; END ;
  GeneTypeExo := CGetTypeExo( DateC ) ;
  FRM.Cpt   := TAna.FindField('Y_SECTION').AsString ;
  FRM.Axe   := TAna.FindField('Y_AXE').AsString ;
  FRM.NumD  := TAna.FindField('Y_NUMEROPIECE').AsInteger ;
  FRM.DateD := DateC ;
  FRM.LigD  := TAna.FindField('Y_NUMLIGNE').AsInteger ;
  AttribParamsNew ( FRM, XD, XC, GeneTypeExo ) ;
  ll := ExecReqMAJ ( fbSect, False, False, FRM ) ;
  If ll<>1 then V_PGI.IoError:=oeSaisie ;
{$ENDIF}
END ;
{$ENDIF} // SANSCOMPTA
{$ENDIF} // EAGLCLIENT

{$IFNDEF SANSCOMPTA}
procedure ChargeLettrage ( XOrig : TL_Rappro ; LOrig : TList ; DEV : RDEVISE ) ;
Var X1,X2 : Tl_Rappro ;
    Q     : TQuery ;
    SQL   : String ;
    Find  : boolean ;
    i     : integer ;
BEGIN
{Recharger le paquet lettré sauf ceux déjà présents dans LOrig ou déjà chargés}
SQL:='Select * from Ecriture Where E_AUXILIAIRE="'+XOrig.Auxiliaire+'" AND E_GENERAL="'+XOrig.General+'" AND E_ETATLETTRAGE="PL" AND E_LETTRAGE="'+XOrig.CodeL+'"' ;
Q:=OpenSQL(SQL,True,-1, '', True) ;
While Not Q.EOF do
   BEGIN
   X2:=TL_Rappro.Create ;
   GetRappro(Q,X2,DEV) ; Find:=False ;
   for i:=0 to LORig.Count-1 do
       BEGIN
       X1:=TL_Rappro(LOrig[i]) ;
       if Not DiffRappro(X1,X2) then BEGIN Find:=True ; Break ; END ;
       END ;
   if Not Find then LOrig.Add(X2) else X2.Free ;
   Q.Next ;
   END ;
Ferme(Q) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/08/2004
Modifié le ... : 02/11/2004
Description .. : Modification de l'algorythme pour faire la gestion du lettrage
Suite ........ : sur reference avant le lettrage sans reference !
Suite ........ : 
Suite ........ : 02/11/2004 : FQ 14827 SBO prise en compte de
Suite ........ : l'évolution de l'état du lettrage pour lettrage en plusieurs
Suite ........ : étape.
Mots clefs ... : 
*****************************************************************}
procedure LettrageEncaDeca ( TEcheOrig : TList ; RR : RMVT ; DEV : RDEVISE ; TOBNumChq : TOB ) ;
Var
    M : RMVT ;
    i,j,k : integer ;
    T : HTStrings ;
    Q : TQuery ;
    LOrig,LSais,LEcr : TList ;
    XOrig,XSais,XOrig2 : TL_Rappro ;
    FindC,OkMontants : boolean;
    CodeL            : String ;
    EtatL            : String ;
    RetourLett       : String ;
BEGIN
    LOrig := TList.Create;
    LSais := TList.Create;
    LEcr := TList.Create;

{ Ecriture générée }

    Q := OpenSQl('Select * from Ecriture Where '+WhereEcriture(tsGene,RR,False)+' AND E_ECHE="X" AND E_NUMECHE>0',true,-1, '', True);
    while (not Q.EOF) do
    begin
        XSais := TL_Rappro.Create;
        GetRappro(Q,XSais,DEV);
        LSais.Add(XSais);
        Q.Next;
    end;
    Ferme(Q) ;

{ Echéances origine }

    for i := 0 to TEcheOrig.Count-1 do
    begin
        T := HTStrings(TEcheOrig[i]) ;
        for j := 0 to T.Count-1 do
        begin
            M := DecodeLC(T[j]);
            Q := OpenSQl('Select * from Ecriture Where '+WhereEcriture(tsGene,M,true),true,-1, '', True);
            if (not Q.EOF) then
            begin
                XOrig := TL_Rappro.Create;
                GetRappro(Q,XOrig,DEV);
                // Ajout SBO Fiche 12017 : pb maj E_NUMTRAITECHQ
                CompleteNumTraiteChq(TobNumChq,XOrig,T[j]);
                LOrig.Add(XOrig);
            end;
            Ferme(Q);
        end;
    end;

{ Lecture des paquets pour partiellement lettrés }

    for i := LOrig.Count-1 downto 0 do
    begin
        XOrig := TL_Rappro(LOrig[i]);
        if ((XOrig.CodeL<>'') and (DansSaisie(XOrig,LSais))) then ChargeLettrage(XOrig,LOrig,DEV);
    end;

{ Lettrage des ecriture avec utilisation de la reference ! }

  { Lecture des origines passage 1 pour 1 => Lettrage total }

    for i:=0 to LOrig.Count-1 do // pour chaque ligne d'origine
    begin
        XOrig := TL_Rappro(LOrig[i]);

        if (XOrig.Solution = 1) then Continue;

        for j:=0 to LSais.Count-1 do // pour chaque ligne generé
        begin
            XSais := TL_Rappro(LSais[j]);
            // ligne deja utilisé
            if (XSais.Solution = 1) then continue;

            // meme cpt general et auxiliaire ??? et Meme Reference ??
            if ((XSais.General = XOrig.General) and (XSais.Auxiliaire = XOrig.Auxiliaire) and (XSais.RefI = XOrig.RefI)) then
            begin
                // meme montant ???
                if (DEV.Code <> V_PGI.DevisePivot)
                  then OkMontants := ((XSais.Debit = XOrig.Credit) and (XSais.Credit = XOrig.Debit))
                  else OkMontants := ((XSais.DebDev = XOrig.CredDev) and (XSais.CredDev = XOrig.DebDev));

                if (OkMontants) then
                begin
                    // mettre les deux ligne dans un paquet
                    LEcr.Clear;
                    LEcr.Add(XOrig);
                    LEcr.Add(XSais);
                    // lettrage 1 pour 1
                    retourLett := LettrerUnPaquet(LEcr,false,true);
                    if RetourLett <> '' then
                      begin
                      // ne pas les reprendre
                      XOrig.Solution := 1;
                      XSais.Solution := 1;
                      // indicateur lettrage total // SBO 28/10/2004 : 14827
                      XOrig.EtatLettrage := 'TL';
                      XSais.EtatLettrage := 'TL';
                      // pour bloqué le double lettrage
                      break;
                      end ;
                end;
            end;
        end;
    end;

  { Lecture des origines passage n pour p => Lettrage partiel ou total pour plusieur lignes ! }

    for i:=0 to LOrig.Count-1 do
    begin
        XOrig := TL_Rappro(LOrig[i]);
        FindC := false;
        LEcr.Clear;

        if (XOrig.Solution = 1) then Continue;

        // lecture a la recherche des ligne ayant meme reference !!
        for j:=0 to LSais.Count-1 do
        begin
            XSais := TL_Rappro(LSais[j]);
            if (XSais.Solution = 1) then Continue;

            // meme cpt general et auxiliaire ??? et Meme Reference ??
            if ((XSais.General = XOrig.General) and (XSais.Auxiliaire = XOrig.Auxiliaire) and (XSais.RefI = XOrig.RefI)) then
            begin
                // ajout au paquet de lettrage !
                FindC := True;
                LEcr.Add(XSais);
                XSais.Solution := 1;
            end;
        end;

        // Si trouvé des echés saisies, rechercher les origines sur les mêmes comptes
        if (FindC) then
        begin
            // ajout de l'origine au paquet de lettrage
            LEcr.Add(XOrig);
            XOrig.Solution := 1;

            // recherche d'autre origine qui on les memes comptes avec meme reference
            for k:=i+1 to LOrig.Count-1 do
            begin
                XOrig2 := TL_Rappro(LOrig[k]);
                if (XOrig2.Solution = 1) then Continue;

                // meme cpt general et auxiliaire ??? et Meme Reference ??
                if ((XOrig2.General = XOrig.General) and (XOrig2.Auxiliaire = XOrig.Auxiliaire) and (XOrig2.RefI = XOrig.RefI)) then
                begin
                    // ajout des origines supplementaire au paquet de lettrage !
                    LEcr.Add(XOrig2);
                    XOrig2.Solution := 1;
                end;
            end;

            // si on a qqchose dans le paquet de lettrage on lettre !
            if (LEcr.Count > 0) then
              begin
              RetourLett := LettrerUnPaquet(LEcr,False,True); // retourne le code lett + ( 'P' ou 'L' ) pour l'état

              // Maj de l'état de lettrage des ecritures traitées  // SBO 28/10/2004 : 14827
              if RetourLett <> '' then
                begin
                CodeL      := Copy(RetourLett,1,4) ;
                EtatL      := Copy(RetourLett,5,1) + 'L' ;
                for k:=0 to LEcr.count-1 do
                  begin
                  TL_Rappro( LEcr[k] ).EtatLettrage := EtatL ;
                  TL_Rappro( LEcr[k] ).CodeL        := CodeL ;
                  end;
                end ;

              end ;
        end;
    end;

{ Lettrage des ecriture sans utilisation de la reference ! }

  { Lecture des origines passage 1 pour 1 => Lettrage total }

    for i:=0 to LOrig.Count-1 do // pour chaque ligne d'origine
    begin
        XOrig := TL_Rappro(LOrig[i]);

        if (XOrig.Solution = 1) then Continue;

        for j:=0 to LSais.Count-1 do // pour chaque ligne generé
        begin
            XSais := TL_Rappro(LSais[j]);

            // ligne deja utilisé
            if (XSais.Solution = 1) then continue;

            // meme cpt general et auxiliaire ???
            if ((XSais.General = XOrig.General) and (XSais.Auxiliaire = XOrig.Auxiliaire)) then
            begin
                // meme montant ???
                if (DEV.Code <> V_PGI.DevisePivot)
                  then OkMontants := ((XSais.Debit = XOrig.Credit) and (XSais.Credit = XOrig.Debit))
                  else OkMontants := ((XSais.DebDev = XOrig.CredDev) and (XSais.CredDev = XOrig.DebDev));

                if (OkMontants) then
                begin
                    // mettre les deux ligne dans un paquet
                    LEcr.Clear;
                    LEcr.Add(XOrig);
                    LEcr.Add(XSais);
                    // lettrage 1 pour 1
                    retourLett := LettrerUnPaquet(LEcr,false,true);
                    if RetourLett <> '' then
                      begin
                      // ne pas les reprendre
                      XOrig.Solution := 1;
                      XSais.Solution := 1;
                      // indicateur lettrage total // SBO 28/10/2004 : 14827
                      XOrig.EtatLettrage := 'TL';
                      XSais.EtatLettrage := 'TL';
                      // pour bloqué le double lettrage
                      break;
                      end ;
                end;
            end;
        end;
    end;

  { Lecture des origines passage n pour p => Lettrage partiel ou total pour plusieur lignes ! }

    for i:=0 to LOrig.Count-1 do
    begin
        XOrig := TL_Rappro(LOrig[i]);
        FindC := false;
        LEcr.Clear;

        // Attention, on évite les TL uniquement ! SBO 28/10/2004 : 14827
        if (XOrig.EtatLettrage = 'TL') then Continue;

        // Lecture des echés de saisie (normalement doivent exister)
        for j:=0 to LSais.Count-1 do
        begin
            XSais := TL_Rappro(LSais[j]);

            // Attention, on évite les TL uniquement ! SBO 28/10/2004 : 14827
            if (XSais.EtatLettrage = 'TL') then Continue;

            // meme cpt general et auxiliaire ???
            if ((XSais.General = XOrig.General) and (XSais.Auxiliaire = XOrig.Auxiliaire)) then
            begin
                // ajout au paquet de lettrage !
                FindC := True;
                LEcr.Add(XSais);
                XSais.Solution := 1;
            end;
        end;

        // Si trouvé des echés saisies, rechercher les origines sur les mêmes comptes
        if (FindC) then
        begin
            // ajout de l'origine au paquet de lettrage
            LEcr.Add(XOrig);
            XOrig.Solution := 1;

            // recherche d'autre origine qui on les memes comptes avec meme reference
            for k:=i+1 to LOrig.Count-1 do
            begin
                XOrig2 := TL_Rappro(LOrig[k]);
                // Attention, on évite les TL uniquement ! SBO 28/10/2004 : 14827
                if (XOrig2.EtatLettrage = 'TL') then Continue;

                // meme cpt general et auxiliaire ???
                if ((XOrig2.General = XOrig.General) and (XOrig2.Auxiliaire = XOrig.Auxiliaire)) then
                begin
                    // ajout des origines supplementaire au paquet de lettrage !
                    LEcr.Add(XOrig2);
                    XOrig2.Solution := 1;
                end;
            end;

            // si on a qqchose dans le paquet de lettrage on lettre !
            if (LEcr.Count > 0) then
              begin
              RetourLett := LettrerUnPaquet(LEcr,False,True); // retourne le code lett + ( 'P' ou 'L' ) pour l'état

              // Maj de l'état de lettrage des ecritures traitées  // SBO 28/10/2004 : 14827
              if RetourLett <> '' then
                begin
                CodeL      := Copy(RetourLett,1,4) ;
                EtatL      := Copy(RetourLett,5,1) + 'L' ;
                for k:=0 to LEcr.count-1 do
                  begin
                  TL_Rappro( LEcr[k] ).EtatLettrage := EtatL ;
                  TL_Rappro( LEcr[k] ).CodeL        := CodeL ;
                  end;
                end ;
              end ;
        end;
    end;

{ Lecture des origines sur elles-même }

    // BPY le 31/10/2003
    // c'est ici que ca bug si on as plusieur ligne d'origine sur un meme compte mais non selectionné
    for i:=0 to LOrig.Count-1 do
    begin
        XOrig := TL_Rappro(LOrig[i]);
        FindC := False;
        LEcr.Clear;

        if (XOrig.Solution = 1) then Continue;

        {Lecture des autres origines}
        for j:=i+1 to LOrig.Count-1 do
        begin
            XOrig2 := TL_Rappro(LOrig[j]);
            if (XOrig2.Solution = 1) then Continue ;
            if ((XOrig2.General=XOrig.General) and (XOrig2.Auxiliaire=XOrig.Auxiliaire)) then
            begin
                LEcr.Add(XOrig2);
                XOrig2.Solution := 1;
                FindC := True;
            end;
        end;
        if (FindC) then
        begin
            LEcr.Add(XOrig);
            XOrig.Solution := 1;
            LettrerUnPaquet(LEcr,False,True);
        end;
    end;

{ Frees }

    VideListe(LOrig);
    LOrig.Free;
    VideListe(LSais);
    LSais.Free ;
    LEcr.Clear;
    LEcr.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 25/08/2003
Modifié le ... :   /  /
Description .. : Recherche le numéro de traite/chq associé à l'échéance
Suite ........ : d'origine et l'affect au TL_Rappro
Mots clefs ... : FICHE 12017
*****************************************************************}
Procedure CompleteNumTraiteChq( TOBNumChq : TOB ; XOrig : TL_Rappro ; vStCode : String ) ;
var TOBD    : TOB ;
begin
  if TOBNumChq=nil then Exit ;
  TOBD := TOBNumChq.FindFirst(['CODELC'],[vStCode], False) ;
  if TOBD<> nil then
    {JP 23/08/05 : GetValue, sous Oracle renvoie parfois #0 => GetString} // FQ 16486
    XOrig.NumTraChq := TOBD.GetString('NUMTRAITECHQ') ;
end ;

{$ENDIF}

end.

