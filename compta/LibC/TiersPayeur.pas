{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 06/08/2004
Modifié le ... :   /  /
Description .. : - CA - 06/08/2004 - FQ 12248 : mise à jour du mode de
Suite ........ : saisie en fonction du journal TP
Mots clefs ... :
*****************************************************************}
unit TiersPayeur;

interface

uses Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, Utob,
     SaisUtil, SaisComm, HEnt1,
{$IFDEF EAGLCLIENT}
{$ELSE}
     Db,
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
     Hctrls,
     LettUtil,
     ParamSoc,
     Ent1, UtilSais,
     uLibAnalytique,
     uLibExercice;  //ctxExercice

procedure GenerePiecesPayeur(M : RMVT;  LePayeur : string='';ListEcr : TList=nil; TobResult : TOB= nil) ;
{$IFDEF COMSX}
procedure GenerePiecesPayeurCom ;
{$ENDIF}
Procedure InverseSoldesTP ( PieceTP : String ) ;
Procedure ExtournePieceTP ( TOBECR: TOB ; DateAnnul : TdateTime) ;
Function  ExisteLettrageSurTP ( Auxi,PieceTP : String ) : boolean ;
Function  WherePieceTP ( PieceTP : String ; Gene : Boolean ) : String ;
Procedure ZoomPieceTP ( PieceTP : String ) ;
Function  VerifPayeurOk ( Auxpay : String ) : Boolean ;

implementation


uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  {$IFNDEF GCGC}
  {$IFNDEF SANSCOMPTA}
  {$IFNDEF IMP}
  {$IFNDEF EAGLSERVER}
  Saisie,
  {$ENDIF}
  {$ENDIF}
  {$ENDIF}
  {$ENDIF}
  UtilPGI;

Function MajRefTP ( TOBTP  : TOB  ): String ;
Var DateComptable : String ;
    Numpiece,NumLigne,Numeche : integer ;
    Journal : String ;
BEGIN
NumPiece:=TOBTP.GetValue('E_NUMEROPIECE') ;
NumLigne:=TOBTP.GetValue('E_NUMLIGNE') ;
NumEche:=TOBTP.GetValue('E_NUMECHE') ;
DateComptable:=FormatDateTime('yyyymmdd',TOBTP.GetValue('E_DATECOMPTABLE')) ;
Journal:=TOBTP.GetValue('E_JOURNAL') ;
Result:=DateComptable+';'+IntToStr(Numpiece)+';'+IntToStr(NumLigne)+';'+IntToStr(NumEche)+';'+Journal+';' ;
END ;

Procedure LettrageTP ( TOBTP : TOB ; CodeLettre : String ) ;
Var Couverture : Double ;
BEGIN
Couverture:=TOBTP.GetValue('E_DEBIT')+TOBTP.GetValue('E_CREDIT') ;
TOBTP.PutValue('E_COUVERTURE',Couverture) ;
Couverture:=TOBTP.GetValue('E_DEBITDEV')+TOBTP.GetValue('E_CREDITDEV') ;
TOBTP.PutValue('E_COUVERTUREDEV',Couverture) ;
TOBTP.PutValue('E_DATEPAQUETMIN',TOBTP.GetValue('E_DATECOMPTABLE')) ;
TOBTP.PutValue('E_DATEPAQUETMAX',TOBTP.GetValue('E_DATECOMPTABLE')) ;
TOBTP.PutValue('E_LETTRAGE',CodeLettre) ;
TOBTP.PutValue('E_ETATLETTRAGE','TL');
if ((TOBTP.GetValue('E_DEVISE')<>V_PGI.DevisePivot) and (TOBTP.GetValue('E_DEVISE')<>V_PGI.DeviseFongible))
   then TOBTP.PutValue('E_LETTRAGEDEV','X') else TOBTP.PutValue('E_LETTRAGEDEV','-') ;
END ;

Procedure ChangeSens ( TOBTP : TOB ; NomChampDeb,NomChampCre : String ) ;
Var X : Double ;
BEGIN
X:=TOBTP.GetValue(NomChampDeb) ;
TOBTP.PutValue(NomChampDeb,TOBTP.GetValue(NomChampCre)) ;
TOBTP.PutValue(NomChampCre,X) ;
END ;

Procedure MajAnaTP ( TOBTP : TOB ; NumP : Integer ; Contrepartie : Boolean ; Jal : String ) ;
Var i,j : Integer ;
    TOBAxeTP,TOBAnaTP : TOB ;
BEGIN
for i:=0 to TOBTP.Detail.Count-1 do
    BEGIN
    TOBAxeTP:=TOBTP.Detail[i] ;
    for j:=0 to TOBAxeTP.Detail.Count-1 do
        BEGIN
        TOBAnaTP:=TOBAxeTP.Detail[j] ;
        TOBAnaTP.PutValue('Y_JOURNAL',Jal) ;
        TOBAnaTP.PutValue('Y_VALIDE','X') ;
        TOBAnaTP.PutValue('Y_CONTREPARTIEGEN',TOBTP.GetValue('E_CONTREPARTIEGEN')) ;
        TOBAnaTP.PutValue('Y_CONTREPARTIEAUX',TOBTP.GetValue('E_CONTREPARTIEAUX')) ;
        TOBAnaTP.PutValue('Y_NUMEROPIECE',NumP) ;
        if ContrePartie then
          BEGIN
          TOBAnaTP.PutValue('Y_NUMLIGNE',2) ;
          END else
          BEGIN
          TOBAnaTP.PutValue('Y_NUMLIGNE',1) ;
          ChangeSens(TOBAnaTP,'Y_DEBIT','Y_CREDIT') ;
          ChangeSens(TOBAnaTP,'Y_DEBITDEV','Y_CREDITDEV') ;
          END ;
        END ;
    END ;
END ;

Procedure InitTOBTP ( TOBTP : TOB ) ;
BEGIN
TOBTP.PutValue('E_LETTRAGE','')  ; TOBTP.PutValue('E_ETATLETTRAGE','AL') ;
TOBTP.PutValue('E_COUVERTURE',0) ; TOBTP.PutValue('E_COUVERTUREDEV',0) ; 
TOBTP.PutValue('E_DATEPAQUETMIN',TOBTP.GetValue('E_DATECOMPTABLE')) ;
TOBTP.PutValue('E_DATEPAQUETMAX',TOBTP.GetValue('E_DATECOMPTABLE')) ;
TOBTP.PutValue('E_CONFIDENTIEL','0') ; TOBTP.PutValue('E_QUALIFORIGINE','TP') ;
TOBTP.PutValue('E_ETAT','0000000000') ; TOBTP.PutValue('E_TIERSPAYEUR','') ;
END ;

Procedure AffecteFromTiers ( TOBTP : TOB ) ;
Var AuxPay : String ;
    Q      : TQuery ;
BEGIN
AuxPay:=TOBTP.GetValue('E_AUXILIAIRE') ;
TOBTP.PutValue('E_RIB',GetRIBPrincipal(AuxPay)) ;
Q:=OpenSQL('Select T_CONSO from TIERS Where T_AUXILIAIRE="'+AuxPay+'"',True,-1,'',true) ;
if Not Q.EOF then TOBTP.PutValue('E_CONSO',Q.Fields[0].AsString) ;
Ferme(Q) ;
END ;

Procedure MajEcrTP ( TOBTP : TOB ; NumP : Integer ; Contrepartie : Boolean ; AuxPay : String ) ;
Var i : integer ;
    lQuery : TQuery ;
BEGIN
InitTOBTP(TOBTP) ;
TOBTP.PutValue('E_NUMEROPIECE',NumP) ; TOBTP.PutValue('E_NUMECHE',1) ; TOBTP.PutValue('E_VALIDE','X') ;
TOBTP.PutValue('E_CONTREPARTIEGEN',TOBTP.GetValue('E_GENERAL')) ;

if Contrepartie then
   BEGIN
    //YMO 18/12/2006 FQ19284 Reprise du compte collectif du tiers payeur sur les nouvelles écritures
    {FQ19284  29.05.07  YMO Uniquement sur la ligne tiers payeur}
    lQuery := OpenSQL('SELECT T_COLLECTIF FROM TIERS WHERE T_AUXILIAIRE="'+AuxPay+'"', True,-1,'',true);
    if not lQuery.Eof then TOBTP.PutValue('E_GENERAL',lQuery.FindField('T_COLLECTIF').AsString) ;
    Ferme(lQuery);
   TOBTP.PutValue('E_NUMLIGNE',2) ;
   TOBTP.PutValue('E_CONTREPARTIEAUX',TOBTP.GetValue('E_AUXILIAIRE')) ;
   TOBTP.PutValue('E_AUXILIAIRE',AuxPay) ;
   END else
   BEGIN
   TOBTP.PutValue('E_NUMLIGNE',1) ;
   TOBTP.PutValue('E_CONTREPARTIEAUX',AuxPay) ;
   ChangeSens(TOBTP,'E_DEBIT','E_CREDIT') ;
   ChangeSens(TOBTP,'E_DEBITDEV','E_CREDITDEV') ;
   TOBTP.PutValue('E_ETAT','---0AM0000') ;
   for i:=1 to 4 do TOBTP.PutValue('E_ECHEENC'+IntToStr(i),-TOBTP.GetValue('E_ECHEENC'+IntToStr(i))) ;
   TOBTP.PutValue('E_ECHEDEBIT',-TOBTP.GetValue('E_ECHEDEBIT')) ;
   END ;
AffecteFromTiers(TOBTP) ;
if ((TOBTP.GetValue('E_DEBIT')>0) or (TOBTP.GetValue('E_CREDIT')<0)) then TOBTP.PutValue('E_ENCAISSEMENT','ENC') else TOBTP.PutValue('E_ENCAISSEMENT','DEC') ;
END ;


Procedure ChainageTP ( TOBContrePass,TOBContrepartie,TOBOrig : TOB ; AuxPay : String ) ;
Var Q : TQuery ;
    Jal,Aux, CodeL,PieceTP : String ;
    NumP : Integer ;
    Nat : String ;
    ModeSaisie : string;
BEGIN
PieceTP:=MajRefTP(TOBOrig) ;
{FQ20298 29.05.07 YMO On prend le 1er jour de l'exercice en cours}
{FQ20298 31.05.07 YMO Dates changées SI exercice fermé}
{$IFDEF NOVH}
If ctxExercice.EstExoClos(TOBOrig.GetValue('E_EXERCICE')) then
begin
  TOBContrepass.PutValue('E_DATECOMPTABLE',DateToStr(GetEncours.Deb)) ; TOBContrepartie.PutValue('E_DATECOMPTABLE',DateToStr(GetEncours.Deb)) ;
  TOBContrepass.PutValue('E_EXERCICE', GetEncours.Code) ; TOBContrepartie.PutValue('E_EXERCICE', GetEncours.Code) ;
end;
{$ELSE}
If ctxExercice.EstExoClos(TOBOrig.GetValue('E_EXERCICE')) then
begin
  TOBContrepass.PutValue('E_DATECOMPTABLE',DateToStr(VH^.Encours.Deb)) ; TOBContrepartie.PutValue('E_DATECOMPTABLE',DateToStr(VH^.Encours.Deb)) ;
  TOBContrepass.PutValue('E_EXERCICE', VH^.Encours.Code) ; TOBContrepartie.PutValue('E_EXERCICE', VH^.Encours.Code) ;
end;
{$ENDIF}
{FQ20300  18.06.07  YMO Datecreation = today}
TOBContrepass.PutValue('E_DATECREATION', DateToStr(Date)) ; TOBContrepartie.PutValue('E_DATECREATION', DateToStr(Date)) ;

TOBContrepass.PutValue('E_PIECETP',PieceTP) ; TOBContrepartie.PutValue('E_PIECETP',PieceTP) ;
// Pièce de Solde
Nat:=TOBOrig.GetValue('E_NATUREPIECE') ;
{$IFNDEF EAGLSERVER}
if ((Nat='FC') or (Nat='AC')) then Jal:=VH^.JalVTP else Jal:=VH^.JalATP ;
{$ELSE}
if ((Nat='FC') or (Nat='AC')) then
Jal := GetParamSocSecur('SO_JALVTP','') else
Jal := GetParamSocSecur('SO_JALATP','');
{$ENDIF}
NumP:=GetNewNumJal(Jal,True,TOBOrig.GetValue('E_DATECOMPTABLE')) ;
TOBContrepass.PutValue('E_JOURNAL',Jal) ; TOBContrepartie.PutValue('E_JOURNAL',Jal) ;
{ CA - 06/08/2004 - FQ 12248 : mise à jour du mode de saisie en fonction du journal TP }
ModeSaisie := GetColonneSQL('JOURNAL','J_MODESAISIE','J_JOURNAL="' + Jal + '"');
TOBContrepass.PutValue('E_MODESAISIE',ModeSaisie) ; TOBContrepartie.PutValue('E_MODESAISIE',ModeSaisie) ;
if ModeSaisie <> '-' then
begin
  TOBContrepass.PutValue('E_NUMGROUPEECR',1) ;
  TOBContrepartie.PutValue('E_NUMGROUPEECR',1) ;
end;
// Lettrage tiers
Aux:=TOBOrig.GetValue('E_AUXILIAIRE') ;
  Q:=OpenSQL('Select T_DERNLETTRAGE from TIERS Where T_AUXILIAIRE="'+Aux+'"',true,-1,'',true) ;
  if Not Q.EOF then
  BEGIN
    CodeL:=Q.FindField('T_DERNLETTRAGE').AsString ;
    CodeL:=CodeSuivant(CodeL) ;
    Ferme (Q);
    ExecuteSQL('UPDATE TIERS SET T_DERNLETTRAGE="'+CodeL+'" WHERE T_AUXILIAIRE = "'+Aux+'"') ;
  END else Ferme(Q) ;

// Ecritures
MajEcrTP(TOBContrepass,NumP,False,AuxPay) ;
MajEcrTP(TOBContrePartie,NumP,True,AuxPay) ;
// Analytiques
If TOBOrig.GetValue('E_ANA')='X' then
  BEGIN
  MajAnaTP(TOBContrepass,NumP,False,Jal) ;
  MajAnaTP(TOBContrepartie,NumP,True,Jal) ;
  END ;
// Facture  origine
PieceTP:=MajRefTP(TOBContrepass) ;
TOBOrig.PutValue('E_PIECETP',PieceTP) ;
TOBOrig.PutValue('E_TIERSPAYEUR',AuxPay) ;
LettrageTP(TOBOrig,CodeL) ;
LettrageTP(TOBContrePass,CodeL) ;
END ;
//{$ENDIF}

procedure GenereDetailPayeur ( TOBEcr : TOB ; AuxPay : String ; ListEcr : TList; TobResult : TOB=nil) ;
Var
  TPass,TContre : TOB ;
  OkInsert : Boolean;
  TobL : TOB;
BEGIN
  // Création des lignes
  TPass:=TOB.Create('ECRITURE',Nil,-1) ;
  TPass.Dupliquer(TOBEcr,True,True) ; TPass.putvalue('E_REFGESCOM','');
  TContre:=TOB.Create('ECRITURE',Nil,-1) ;
  TContre.Dupliquer(TOBEcr,True,True) ; TContre.putvalue('E_REFGESCOM','');
  //{$IFNDEF EAGLSERVER}
  // Traitement de maj des infos
  ChainageTP(TPass,TContre,TOBEcr,AuxPay) ;
  //{$ENDIF}
  // Ecriture base
  OkInsert := TPass.InsertDB(Nil) ;
  if OkInsert then
    OkInsert := TContre.InsertDB(Nil) ;
  TOBEcr.UpDateDB(False) ;
  MajDesSoldesTOB(TPass,True) ;
  MajDesSoldesTOB(TContre,True) ;
  if (OkInsert) and (Assigned(TobResult)) then
  begin
    TobL := TOB.Create('ECRITURE', TobResult, -1);
    TobL.Dupliquer(TPass, True, True);
    TobL := TOB.Create('ECRITURE', TobResult, -1);
    TobL.Dupliquer(TContre, True, True);
  end;
  If ListEcr<>nil then
  begin
    ListEcr.Add(TPass) ;
    ListEcr.Add(TContre) ;
  end
  else
  begin
  // GP MEMCHECK
   TPass.Free ;
   TContre.Free ;
  end;
END ;

procedure GenerePiecesPayeur(M : RMVT;  LePayeur : string='';ListEcr : TList=nil; TobResult : TOB= nil) ;
Var Q,QAna : TQuery ;
    Ax     : String ;
    SQL,AuxPay,SAna,NumLigne : String ;

    TOBEcr,TOBAna : TOB ;
    WithPayeur : boolean;
BEGIN

  // Test préalables
  {$IFNDEF EAGLSERVER}
  if Not VH^.OuiTP then Exit ;
  if Not EstJalFact(M.Jal) then Exit ;
  {$ELSE}
  if not GetParamSocSecur('SO_OUITP','') then exit;

  if not ExisteSQL('SELECT J_JOURNAL, J_NATUREJAL FROM JOURNAL WHERE J_JOURNAL = "'+ M.Jal+'"'
                 + ' AND (J_NATUREJAL = "VTE" OR J_NATUREJAL = "ACH")') then exit;
  {$ENDIF}

  {$IFDEF COMPTA}
  //YMO 10/01/2006 si pas de gestion auto on sort (sauf avec liste de restitution : cas du lancement manuel)
  If (ListEcr=nil) and (not GetParamSocSecur('SO_AUTOTP',False)) then exit;
  {$ENDIF}

  {Pour les cas d'une pièce avec plusieurs tiers, lors du lancement manuel, on différencie les lignes}
  If not GetParamSocSecur('SO_AUTOTP',False) then
    NumLigne:=' AND E_NUMLIGNE='+IntToStr(M.NumLigne)
  else
  {Pour de la génération lors de la saisie, pas de distinction, on génère tous les TP}
    NumLigne:='';

  if ((M.Nature <> 'FC') and (M.Nature <> 'FF') and (M.Nature <> 'AC') and (M.Nature <> 'AF')) then Exit ;

  {$IFNDEF EAGLSERVER}
  if (VH^.JalVTP = '') and ((M.Nature = 'FC') or (M.Nature = 'AC')) then exit;
  if (VH^.JalATP = '') and ((M.Nature = 'FF') or (M.Nature = 'AF')) then exit;
  {$ENDIF}
  // Traitement
  SQL:= 'SELECT ECRITURE.*, T_PAYEUR, T_ISPAYEUR, T_AVOIRRBT from ECRITURE'
      + ' LEFT JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE'
      + ' WHERE ' + WhereEcriture(tsGene,M,False) + ' AND E_AUXILIAIRE <> "" AND T_PAYEUR <> "" AND T_ISPAYEUR = "-" AND E_ETATLETTRAGE = "AL"'
      + ' AND ((E_NATUREPIECE <> "AC" AND E_NATUREPIECE <> "AF") OR (T_AVOIRRBT = "-"))'
      + ' AND E_ETATLETTRAGE = "AL"' + NumLigne
      + ' ORDER BY E_NUMLIGNE, E_NUMECHE';
  Q := OpenSQL(SQl, True,-1,'',true) ;
  // Si on ne trouve pas, on refait la requête sans le tiers et on affecte la valeur
  // AuxPay en prenant celui passé en paramètre
  if (Q.EOF) and (LePayeur <> '') then
  begin
    WithPayeur := false;
    Ferme(Q);
    SQL:= 'SELECT ECRITURE.*, T_PAYEUR, T_ISPAYEUR, T_AVOIRRBT FROM ECRITURE'
        + ' LEFT JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE'
        + ' WHERE ' + WhereEcriture(tsGene,M,False) + ' AND E_AUXILIAIRE <> "" AND E_ETATLETTRAGE = "AL"'
        + ' AND ((E_NATUREPIECE <> "AC" AND E_NATUREPIECE <> "AF") OR (T_AVOIRRBT = "-"))'
        + ' AND E_ETATLETTRAGE = "AL"' + NumLigne
        + ' ORDER BY E_NUMLIGNE, E_NUMECHE' ;
    Q := OpenSQL(SQL, True,-1,'',true);
  end else
    WithPayeur := True;
  While Not Q.EOF do
  begin
    TOBEcr := TOB.Create('ECRITURE',Nil,-1) ;
    TOBEcr.SelectDB('',Q) ;
    if WithPayeur then
      AuxPay := Q.FindField('T_PAYEUR').AsString
      else
      AuxPay := LePayeur;
    // Création des axes analytiques
    AlloueAxe(TobEcr) ; // Modif SBO 27/09/2004 : FQ 14632
    if TOBEcr.GetBoolean('E_ANA') then
    begin
      SAna :='SELECT * FROM ANALYTIQ WHERE '+WhereEcriture(tsAnal,M,False)
           + ' AND Y_NUMLIGNE='+IntToStr(TOBEcr.GetValue('E_NUMLIGNE'))
           + ' ORDER BY Y_AXE, Y_NUMLIGNE, Y_NUMVENTIL' ;
      QAna := OpenSQL(SAna,True,-1,'',true) ;
      if not QAna.EOF then
      begin
        While not QAna.EOF do
        begin
          TOBAna := TOB.Create('ANALYTIQ',Nil,-1) ;
          TOBAna.SelectDB('',QAna) ;
          Ax := TOBAna.GetValue('Y_AXE') ;
          TOBAna.ChangeParent(TOBEcr.Detail[Ord(Ax[2])-49],-1) ;
          QAna.Next ;
        end;
      end;
      Ferme(QAna) ;
    end;
    If VerifPayeurOk(Auxpay) then {FQ19277  19.06.07  YMO Verification compte ouvert}
        GenereDetailPayeur(TOBEcr, AuxPay, ListEcr, TobResult) ;
    TOBEcr.Free ;
    Q.Next ;
  end;
  Ferme(Q) ;
end;

{$IFDEF COMSX}
procedure GenerePiecesPayeurCom ;
Var Q,QAna : TQuery ;
    Ax     : String ;
    SQL,AuxPay,SAna : String ;
    TOBEcr,TOBAna : TOB ;
BEGIN
// Traitement
SQL:='Select Ecriture.*, T_PAYEUR, T_ISPAYEUR, T_AVOIRRBT from ECRITURE '
    +'Left Join Tiers on E_AUXILIAIRE=T_AUXILIAIRE '
    +'Where E_FLAGECR="TP" AND E_AUXILIAIRE<>"" AND T_PAYEUR<>"" AND T_ISPAYEUR="-" AND E_ETATLETTRAGE="AL" '
    +'AND ((E_NATUREPIECE<>"AC" AND E_NATUREPIECE<>"AF") OR (T_AVOIRRBT="-")) '
    +'AND E_ETATLETTRAGE="AL" '
    +'Order By E_NUMLIGNE, E_NUMECHE ' ;
Q:=OpenSQL(SQl,True,-1,'',true) ;
While Not Q.EOF do
   BEGIN
   TOBEcr:=TOB.Create('ECRITURE',Nil,-1) ;
   TOBEcr.SelectDB('',Q) ;
   AuxPay:=Q.FindField('T_PAYEUR').AsString ;
   // Création des axes analytiques
   AlloueAxe( TobEcr ) ; // Modif SBO 27/09/2004 : FQ 14632
   if TOBEcr.GetValue('E_ANA')='X' then
      BEGIN
      SAna:='Select * from ANALYTIQ Where Y_JOURNAL="' + TOBEcr.GetValue ('E_JOURNAL')+'" AND Y_EXERCICE="'+TOBEcr.GetValue ('E_EXERCICE')+'"'
              +' AND Y_DATECOMPTABLE="'+UsDateTime(TOBEcr.GetValue ('E_DATECOMPTABLE'))+'" AND Y_NUMEROPIECE='+TOBEcr.GetValue ('Y_NUMEROPIECE')
              +' AND Y_QUALIFPIECE="'+TOBEcr.GetValue ('E_QUALIFPIECE')+'" AND Y_NATUREPIECE="'+TOBEcr.GetValue ('E_NATUREPIECE')+'"'
              +' AND Y_NUMLIGNE='+IntToStr(TOBEcr.GetValue('E_NUMLIGNE'))
              +' ORDER BY Y_AXE, Y_NUMLIGNE, Y_NUMVENTIL' ;
      QAna:=OpenSQL(SAna,True,-1,'',true) ;
      if Not QAna.EOF then
         BEGIN
         While Not QAna.EOF do
            BEGIN
            TOBAna:=TOB.Create('ANALYTIQ',Nil,-1) ;
            TOBAna.SelectDB('',QAna) ;
            Ax:=TOBAna.GetValue('Y_AXE') ;
            TOBAna.ChangeParent(TOBEcr.Detail[Ord(Ax[2])-49],-1) ;
            QAna.Next ;
            END ;
         END ;
      Ferme(QAna) ;
      END ;
   If VerifPayeurOk(Auxpay) then {FQ19277  19.06.07  YMO Verification compte ouvert}
       GenereDetailPayeur(TOBEcr,AuxPay, nil) ;
   TOBEcr.Free ;
   Q.Next ;
   END ;
Ferme(Q) ;
END ;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 19/06/2007
Modifié le ... :   /  /    
Description .. : FQ19277  YMO Verification compte TP ouvert
Mots clefs ... : FQ19277 COMPTE FERMÉ OUVERT
*****************************************************************}
Function VerifPayeurOk ( Auxpay : String ) : Boolean ;
begin
Result:=False ;
If ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE="'+AuxPay+'" AND T_FERME="-"') then
  Result:=True;
end;

Procedure ZoomPieceTP ( PieceTP : String ) ;
{$IFNDEF EAGLSERVER}
{$IFNDEF GCGC}
{$IFNDEF IMP}
Var RRR : RMVT ;
    QQ  : TQuery ;
    SQL : String ;
{$ENDIF !GCGC}
{$ENDIF IMP}
{$ENDIF EAGLSERVER}
BEGIN
{$IFNDEF EAGLSERVER}
{$IFNDEF SANSCOMPTA}
{$IFNDEF IMP}
{$IFNDEF GCGC}
FillChar(RRR,Sizeof(RRR),#0) ;
SQL:='SELECT ' + SQLForIdent( fbGene ) + ' FROM ECRITURE WHERE '+WherePieceTP(PieceTP,True) ;
QQ:=OpenSQL(SQL,True) ;
if Not QQ.EOF then RRR:=MvtToIdent(QQ,fbGene,True) ;
Ferme(QQ) ;
if RRR.Jal<>'' then
  {$IFNDEF EAGLCLIENT}
    LanceSaisie(Nil,taConsult,RRR) ;
  {$ELSE}
    LanceSaisie(Nil,taConsult,RRR) ;
  {$ENDIF}
{$ENDIF}
{$ENDIF}
{$ENDIF}
{$ENDIF}
END ;

Function WherePieceTP ( PieceTP : String ; Gene : Boolean ) : String ;
// DBR Pour supprimer les conseils et avertissements
//Var NumP,NumL : integer ;
Var NumP : integer ;
    Jal,Exo,St,StC : String ;
    DateC   : TDateTime ;
BEGIN
St:=PieceTP ;
StC:=ReadTokenSt(St) ; DateC:=EncodeDate(ValeurI(Copy(StC,1,4)),ValeurI(Copy(StC,5,2)),ValeurI(Copy(StC,7,2))) ;
StC:=ReadTokenSt(St) ; NumP:=ValeurI(StC) ;
StC:=ReadTokenSt(St) ;
StC:=ReadTokenSt(St) ;
StC:=ReadTokenSt(St) ; Jal:=StC ;
Exo:=QuelExoDT(DateC) ;
if Gene then Result:=' E_JOURNAL="'+Jal+'" AND E_EXERCICE="'+Exo+'" AND E_DATECOMPTABLE="'+UsDateTime(DateC)+'" AND E_NUMEROPIECE='+IntToStr(NumP)+' AND E_QUALIFPIECE="N" '
        else Result:=' Y_JOURNAL="'+Jal+'" AND Y_EXERCICE="'+Exo+'" AND Y_DATECOMPTABLE="'+UsDateTime(DateC)+'" AND Y_NUMEROPIECE='+IntToStr(NumP)+' AND Y_QUALIFPIECE="N" '
END ;

Function ExisteLettrageSurTP ( Auxi,PieceTP : String ) : boolean ;
Var SQL : String ;
    Q   : TQuery ;
BEGIN
Result:=False ;
SQL:='Select E_ETATLETTRAGE FROM ECRITURE WHERE '+WherePieceTP(PieceTP,True)+' AND E_AUXILIAIRE="'+Auxi+'"' ;
Q:=OpenSQL(SQL,True,-1,'',true) ;
if Not Q.EOF then Result:=(Q.Fields[0].AsString<>'AL') ;
Ferme(Q) ;
END ;

procedure ChargeEcrTiersPayeur (TOBECRTP : TOB; PieceTP : string);
var Q : TQuery;
		TOBAna,TOBA,TOBE,TOBAxe : TOB;
    i,k,NumA,NumL : integer;
    Ax : String ;
begin
  Q:=OpenSQL('SELECT * FROM ECRITURE WHERE '+WherePieceTP(PieceTP,True),True,-1, '', True) ;
  if Not Q.EOF then TOBECRTP.LoadDetailDB('ECRITURE','','',Q,False,True) ;
  Ferme(Q) ;
  {Analytiques}
  TOBAna:=TOB.Create('',Nil,-1) ;
  Q:=OpenSQL('SELECT * FROM ANALYTIQ WHERE '+WherePieceTP(PieceTP,false),True,-1, '', True) ;
  if Not Q.EOF then TOBAna.LoadDetailDB('ANALYTIQ','','',Q,False,True) ;
  Ferme(Q) ;
  {Changement de parent}
  TOBAna.Detail.Sort('Y_NUMLIGNE;Y_NUMVENTIL') ;
  for i:=TOBAna.Detail.Count-1 downto 0 do
  BEGIN
    TOBA:=TOBAna.Detail[i] ;
    NumL:=TOBA.GetValue('Y_NUMLIGNE') ;
    Ax:=TOBA.GetValue('Y_AXE') ; NumA:=Ord(Ax[2])-48 ;
    TOBE:=TOBECRTP.FindFirst(['E_NUMLIGNE'],[NumL],False) ;
    if TOBE.Detail.Count<=0 then BEGIN for k:=1 to 5 do TOB.Create('A'+IntToStr(k),TOBE,-1) ; END ;
    TOBAxe:=TOBE.Detail[NumA-1] ; TOBA.ChangeParent(TOBAxe,-1) ;
  END ;
end;

Procedure ExtournePieceTP ( TOBECR: TOB ; DateAnnul : TdateTime) ;
var TOBE,TOBECRTP,TOBEcrEx : TOB;
		PieceTP : string;
BEGIN
  TOBECRTP := TOB.Create ('LES ECR',nil,-1);
  TOBEcrEx := TOB.Create ('LES ECR',nil,-1);
  TRY
    if Not VH^.OuiTP then Exit ;
    if TOBECR.Detail.Count<=0 then Exit ;
    TOBE:=TOBECR.Detail[0] ;
    if TOBE.GetValue('E_QUALIFPIECE')<>'N' then Exit ;
    PieceTP:= TOBE.getValue('E_PIECETP');
    if PieceTP = '' then exit;
    ChargeEcrTiersPayeur (TOBEcrTP,PieceTp);
    {Extourne ecriture}
    TOBEcrEx := ExtourneEcriture(TOBECRTP,false,StrToDate(DateToStr(DateAnnul)),TobEcr.detail[0].GetString('E_QUALIFPIECE'),
    								 						 GetParamSocSecur('SO_MONTANTNEGATIF', false), '', '', '', False);
    if not TOBEcrEx.InsertDBByNivel(False) then
    begin
      V_PGI.IoError := oeLettrage;
    end else
    begin
      MajSoldesEcritureTOB(TobEcrEx, true);
      if V_PGI.IOError <> oeOk then
      begin
        V_PGI.IoError := oeLettrage;
      end;
    end;
  FINALLY
  	TOBECRTP.free;
    TOBEcrEx.free;
  END;
END ;

Procedure InverseSoldesTP ( PieceTP : String ) ;
Var LaTOB : TOB ;
    Q     : TQuery ;
BEGIN
if PieceTP='' then Exit ;
LaTOB:=TOB.Create('ECRITURE',Nil,-1) ;
Q:=OpenSQL('Select * from Ecriture Where '+WherePieceTP(PieceTP,True),True,-1,'',true) ;
While Not Q.EOF do
   BEGIN
   LaTOB.SelectDB('',Q) ; MajSoldeCompteTOB(LaTOB,False) ;
   Q.Next ;
   END ;
Ferme(Q) ;
LaTOB:=TOB.Create('ANALYTIQ',Nil,-1) ;
Q:=OpenSQL('Select * from Analytiq Where '+WherePieceTP(PieceTP,False),True,-1,'',true) ;
While Not Q.EOF do
   BEGIN
   LaTOB.SelectDB('',Q) ; MajSoldeSectionTOB(LaTOB,False) ;
   Q.Next ;
   END ;
Ferme(Q) ;
END ;

end.

