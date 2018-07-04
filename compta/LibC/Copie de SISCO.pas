{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 04/06/2003
Modifié le ... : 19/09/2003
Description .. : FQ 12163 - CA - 04/06/2003 - Récupération écritures sur
Suite ........ : collectif paie.
Suite ........ : FQ 10001 - CA - 19/09/2003 - Paramsoc généraux non mis 
Suite ........ : à jour dans le cas d'un récup PCL car fait par TRFS5 lors de 
Suite ........ : la création du dossier
Mots clefs ... : 
*****************************************************************}
unit SISCO;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Hctrls, StdCtrls, Buttons, ExtCtrls, PrintDBG, Ent1, HSysMenu,
  hmsgbox,HStatus,HEnt1,DBTables, ColMemo, ComCtrls, HTB97, VisuEnr, HCompte,RappType,ImpFicU,TImpFic,
  Menus,ImpUtil,CpteUtil,CritEdt,Math, ParamSoc, UtilTrans, UTOB, RecupUtil ;


Function TransfertSISCO(StFichier : String ; RemplaceLeFichier : Boolean ; Var NbFic : Integer ; NewRep : String = '' ;
                        InfoImp : PtTInfoImport = Nil ; ForceDecoupe : Boolean = FALSE) : Integer ;
Function InitRecupDossier(StFichier : String ; Var InfoImp : TInfoImport) : Integer ;
Procedure RecupSectionSISCO(Var St,St1 : String ; Var Infos : TInfoGeneSISCO) ;
Procedure RecupCptSISCO(Var St,St1,StSup :String ; Var Infos : TInfoGeneSISCO ; InfoImp : PtTInfoImport = Nil; Ident : string = 'CAU') ;
Procedure chargeCptSISCO(Var CptSISCO : TTabCptCollSISCO ; ProfilPCL : String = '') ;
function GetNatureDiv(Cpt : string; Sisco : TTabCptCollSISCO) : string ;
Procedure RecupSociete(i : Integer ; St : String ; Var Info : TInfoGeneSISCO ; InfoImp : PtTInfoImport = Nil) ;
Procedure ChargeCharRemplace(Var InfoGene : TInfoGeneSISCO) ;
Function CptRemplace(Cpt : String ; CR : TCharRemplace) : String ;
Procedure ChargeDiversDefaut(Var InfoGene : TInfoGeneSISCO) ;
Procedure ExportBALSISCO(StFichier : String ; Var CritEdt : TCritEdt ; Valeur : String ; EtatRevision : TCheckBoxState) ;
Function TrouveCptSISCO(Cpt : String ; Var CptSISCO : TTabCptCollSISCO ; Var Nat : String) : String ;
Procedure QuelCompte(Cpt : String ; Var CptSISCO : TTabCptCollSISCO ; Var Gen,Aux : String) ;

// ajout me
Function RecupJalSISCO(Var St,St1 : String ; Var Infos : TInfoGeneSISCO ; InfoImp : PtTInfoImport = Nil) : Boolean ;
Function EcritPiece(Var NewFichier : TextFile ; ListePiece : TStringList ; Var IdentPiece : TIdentPiece ; Var Infos : TInfoGeneSISCO ;
                    Var PieceEcrite : Boolean ;
                    InfoImp : PtTInfoImport = Nil) : Integer ;
Function FaitPiece(St : String ; ListePiece : TStringList ; Var IdentPiece : TIdentPiece ;
                    Var OldStSISCO : TStSISCO ; Var Infos : TInfoGeneSISCO ;
                    InfoImp : PtTInfoImport = Nil ; RuptOrigineSaisie : Boolean = FALSE) : Integer ;
Function PieceEquilibree(Var IdentPiece : TIdentPiece) : Boolean ;
Function DecimSISCO(Var IdentPiece : TIdentPiece ; Oppose : Boolean) : Integer ;
Function RuptureOrigineSaisie(Var St : String ; Var IdentPiece : TIdentPiece ; Var ChronoEclateIncremente : Boolean) : Boolean ;
Function RuptureJour(Var St : String ; OldStSISCO : TStSISCO) : Boolean ;
// ajout me return string
Function CloseEtRenameNewFile(RemplaceLeFichier : Boolean ; StFichier,StNewFichier,NewRep  : String ; CompteurFichier : Integer) : string ;
Function FourchetteSISCOExiste(What : Integer ; Var CptSISCO : TTabCptCollSISCO) : Boolean ;
Procedure AlimCptResultat(Var CritEdt : TCritEdt) ;
Procedure FaitSISCODebut(Var FF : TextFile) ;                    // ajout me
Procedure FaitSISCO00(Var FF : TextFile ; Var CritEdt : TCritEdt; Typarchive : string='B') ;
Procedure FaitSISCOCptGen(Var FF : TextFile ; Var CritEdt : TCritEdt; bExit : boolean = True) ;
Procedure FaitSISCOPer(Var FF : TextFile ; DD : TDateTime ; Exo : TExoDate) ;
Procedure FaitSISCOFolio(Var FF : TextFile ; NF : Integer) ;
FUNCTION STRFMONTANTSISCO ( Tot : Extended ; Long,Dec : Integer ; symbole : string3 ; Separateur : Boolean) : string ;
Procedure FaitSISCOFin(Var FF : TextFile) ;
Function  GetNomDossier(StFichier : String) : String ;

Type tRuptEc = (OnQued,OnFolio,OnJal) ;

Type TRecupSISCO   = Record
                     NbLCpt,NbLMvt : Integer ;
                     NomFicOrigine,NomFicMvt,NomFicCpt : String ;
                     NomFicCptCGE,NomFicMvtCGE : String ;
                     RuptEc : TRuptEc ;
                     LM : TStringList ;
                     Ind : Integer ;
                     LastIndEcr : Integer ;
                     LigSISCO : Integer ;
                     InfoSISCO : TInfoGeneSISCO ;
                     End ;

implementation

Const DecalListe = 8 ;
Const OkCROS : Boolean = FALSE ;
Const ForceTenue : Boolean = TRUE ;

//Procedure CloseEtRenameNewFile(RemplaceLeFichier : Boolean ; StFichier,StNewFichier,NewRep  : String ; CompteurFichier : Integer) ;
Function CloseEtRenameNewFile(RemplaceLeFichier : Boolean ; StFichier,StNewFichier,NewRep  : String ; CompteurFichier : Integer) : string ;  // ajout me
Var NewFichier : TextFile ;
    StNewFichier2 : String ;
    StPlus : String ;
BEGIN
StPlus:='' ;
// ajout me
StNewFichier2 := '';
If CompteurFichier>-1 Then StPlus:=formatFloat('000',CompteurFichier) ;
If RemplaceLeFichier Then
  BEGIN
  StNewFichier2:=NewNomFic(StFichier,'OLD') ;
  {$i-}
  AssignFile(NewFichier,StNewFichier2) ; Erase(NewFichier) ;
  {$i+}
  Hrenamefile(StFichier,StNewFichier2) ;
  Hrenamefile(StNewFichier,StFichier) ;
  END Else
  BEGIN
  If NewRep='' Then
    BEGIN
    If RecupSISCO Then StNewFichier2:=NewNomFic(StFichier,'CGE'+StPlus)
                  Else StNewFichier2:=NewNomFic(StFichier,'CGN') ;
    END Else
    BEGIN
    If RecupSISCO Then StNewFichier2:=NewNomFicEtDir(StFichier,'CGE'+StPlus,NewRep)
                  Else StNewFichier2:=NewNomFicEtDir(StFichier,'CGN'+StPlus,NewRep) ;
    END ;
  {$i-}
  AssignFile(NewFichier,StNewFichier2) ; Erase(NewFichier) ;
  {$i+}
  Hrenamefile(StNewFichier,StNewFichier2) ;
  END ;
FichierOnDisk(StNewFichier2,FALSE) ;
// ajout me
Result := StNewFichier2;
END ;

Procedure CloseOpenNewFile(Var NewFichier : TextFile ; Var StNewFichier : String ;
                           RemplaceLeFichier : Boolean ; StFichier,NewRep : String ;
                           Var CompteurFichier : Integer)  ;
BEGIN
CloseFile(NewFichier) ;
CloseEtRenameNewFile(RemplaceLeFichier,StFichier,StNewFichier,NewRep,CompteurFichier) ;
If CompteurFichier>-1 Then Inc(CompteurFichier) ;
StNewFichier:=FileTemp('.PNM') ;
AssignFile(NewFichier,StNewFichier) ; Rewrite(NewFichier) ;
Writeln(NewFichier,StFichier) ;
END ;

Function DecimSISCO(Var IdentPiece : TIdentPiece ; Oppose : Boolean) : Integer ;
BEGIN
Result:=V_PGI.OkDecV ;
{$IFDEF NOCONNECT}
Exit ;
{$ENDIF}
If IdentPiece.CodeMontant<>'' Then
  BEGIN
  If (VH^.TenueEuro) And (IdentPiece.CodeMontant[1]='F') Then Result:=V_PGI.OkDecE Else
  If (Not VH^.TenueEuro) And (IdentPiece.CodeMontant[1]='E') Then Result:=V_PGI.OkDecE ;
  END ;
If Oppose Then
  BEGIN
  If Result=V_PGI.OkDecV Then Result:=V_PGI.OkDecE Else Result:=V_PGI.OkDecV ;
  END ;
END ;

Function CptRemplace(Cpt : String ; CR : TCharRemplace) : String ;
Var ll : Integer ;
BEGIN
If VH^.RecupPCL Then
  BEGIN
  While Pos('%',Cpt)>0 Do BEGIN ll:=Pos('%',Cpt) ; Cpt[ll]:='0' ; END ;
  While Pos('_',Cpt)>0 Do BEGIN ll:=Pos('_',Cpt) ; Cpt[ll]:='0' ; END ;
  While Pos('''',Cpt)>0 Do BEGIN ll:=Pos('''',Cpt) ; Cpt[ll]:='0' ; END ;
  While Pos('"',Cpt)>0 Do BEGIN ll:=Pos('"',Cpt) ; Cpt[ll]:='0' ; END ;
  END Else
  BEGIN
  If Trim(CR.St1)<>'' Then While Pos('%',Cpt)>0 Do BEGIN ll:=Pos('%',Cpt) ; Cpt[ll]:=CR.St1[1] ; END ;
  If Trim(CR.St2)<>'' Then While Pos('_',Cpt)>0 Do BEGIN ll:=Pos('_',Cpt) ; Cpt[ll]:=CR.St2[1] ; END ;
  END ;
Result:=Cpt ;
END ;

Procedure RecupCharRemplace(St : String ; Var CH1,CH2 : String) ;
//CHG CHT CHS CHJ
Var Q : TQuery ;
BEGIN
CH1:='' ; CH2:='' ;
{$IFDEF NOCONNECT}
Exit ;
{$ENDIF}
Q:=OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="SIS" AND CC_CODE="'+St+'"',TRUE) ;
If Not Q.Eof Then
  BEGIN
  CH1:=Q.FindField('CC_LIBELLE').AsString ; CH2:=Q.FindField('CC_ABREGE').AsString ;
  END ;
Ferme(Q) ;
END ;

Function RecupDiversDefaut(St : String) : String ;
// RGT : Régime Tva, MRT : Mode règlement
Var Q : TQuery ;
BEGIN
{$IFDEF NOCONNECT}
Exit ;
{$ENDIF}
Result:='' ;
Q:=OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="SIS" AND CC_CODE="'+St+'"',TRUE) ;
If Not Q.Eof Then Result:=Q.FindField('CC_LIBELLE').AsString ;
Ferme(Q) ;
END ;

Procedure ChargeDiversDefaut(Var InfoGene : TInfoGeneSISCO) ;
BEGIN
InfoGene.RGT:=RecupDiversDefaut('RGT') ;
InfoGene.MRT:=RecupDiversDefaut('MRT') ;
END ;

Procedure ChargeCharRemplace(Var InfoGene : TInfoGeneSISCO) ;
BEGIN
{$IFDEF NOCONNECT}
Exit ;
{$ENDIF}
RecupCharRemplace('CHG',InfoGene.CharRG.St1,InfoGene.CharRG.St2) ;
RecupCharRemplace('CHT',InfoGene.CharRG.St1,InfoGene.CharRT.St2) ;
RecupCharRemplace('CHS',InfoGene.CharRG.St1,InfoGene.CharRS.St2) ;
RecupCharRemplace('CHJ',InfoGene.CharRG.St1,InfoGene.CharRJ.St2) ;
END ;

Function TrouveCptSISCO(Cpt : String ; Var CptSISCO : TTabCptCollSISCO ; Var Nat : String) : String ;
Var i : Integer ;
    ll,ll1,ll2,ll3 : Integer ;
BEGIN
Result:='' ; Nat:='' ;
{$IFDEF NOCONNECT}
Exit ;
{$ENDIF}
For i:=0 To 50 Do
  BEGIN
  If CptSISCO[i].Cpt='' Then Exit ;
  ll1:=Length(CptSISCO[i].Aux1) ; ll2:=Length(CptSISCO[i].Aux2) ;
//  ll3:=Length(Cpt) ;
// CA - 02/08/2002 - Corrige problème compte 9A avec coll. de 9A à 9ZZZZ : 9A n'était pas vu comme
// appartenant à la tranche car des blancs après A augmentaient Length(Cpt)
  ll3:=Length(Trim(Cpt)) ;
  ll:=ll3 ; If ll1<ll Then ll:=ll1 ; If ll2<ll Then ll:=ll2 ;
  If (Copy(Cpt,1,ll)>=Copy(CptSISCO[i].Aux1,1,ll)) And (Copy(Cpt,1,ll)<=Copy(CptSISCO[i].Aux2,1,ll)) Then
    BEGIN
    Result:=CptSISCO[i].Cpt ; Nat:=CptSISCO[i].Nat ; Exit ;
    END ;
  END ;
END ;

Function TrouveCollSISCO(Cpt : String ; Var CptSISCO : TTabCptCollSISCO) : Boolean ;
Var i : Integer ;
    ll : Integer ;
BEGIN
Result:=FALSE ;
For i:=0 To 50 Do
  BEGIN
  If CptSISCO[i].Cpt='' Then Exit ;
  ll:=Length(Cpt) ;
  If (Copy(CptSISCO[i].Cpt,1,ll)=Cpt) Then BEGIN Result:=TRUE ; Exit ; END ;
  END ;
END ;

Function FourchetteSISCOExiste(What : Integer ; Var CptSISCO : TTabCptCollSISCO) : Boolean ;
Var St : String ;
    i : Integer ;
BEGIN
Result:=FALSE ;
For i:=0 To 50 Do
  BEGIN
  If CptSISCO[i].Cpt='' Then Exit ;
  St:=Copy(CptSISCO[i].Cpt,1,2) ;
  If St<>'' Then
    BEGIN
    Result:=((St='40') And (What=0)) Or ((St='41') And (What=1)) ;
    If Result Then Break ;
    END ;
  END ;
END ;

Function EstLettrableSISCO(Cpt,Racine : String) : Boolean ;
Var St : String ;
    ll : Integer ;
BEGIN
If Trim(Racine)='' Then BEGIN Result:=FALSE ; Exit ; END ;
Result:=TRUE ; If Racine[Length(Racine)]<>';' Then Racine:=Racine+';' ;
While Racine<>'' Do
  BEGIN
  St:=ReadTokenSt(Racine) ;
  If St<>'' Then
    BEGIN
    ll:=Length(St) ;
    If Copy(Cpt,1,ll)=St Then BEGIN Result:=TRUE ; Exit ; END ;
    END ;
  END ;
Result:=FALSE ;
END ;

Function OkCollCEGID(Cpt : String ; Var coll416 : Boolean) : Boolean ;
Var St,St1 : String ;
BEGIN
Result:=FALSE ; coll416:=FALSE ;
St:=Copy(Cpt,1,5) ;
If (St='41100') Or (St='41101') Or (St='41111') Or
   (St='40100') Or (St='40110') Then Result:=TRUE ;
St:=Copy(Cpt,1,7) ;
If (St='4160011') Or (St='4160012') Or (St='4160013') Or
   (St='4160014') Or (St='4160015') Or (St='4160019') Then Result:=TRUE ;
If (St='4160021') Or (St='4160022') Or (St='4160023') Or
   (St='4160024') Or (St='4160029') Then BEGIN Result:=TRUE ; coll416:=TRUE ; END ;

END ;

Function PasOkCollCEGID(Cpt : String) : Boolean ;
Var St,St1 : String ;
BEGIN
Result:=TRUE ;
St:=Copy(Cpt,1,8) ;
If (St='41600191') Or (St='41600291') Or (St='41600170') Then Result:=FALSE ;
END ;

Procedure TraiteCptCegid(Var Gen,Aux : String ; OkColl,coll416 : Boolean) ;
Var Q : TQuery ;
    St,Cpt2,Aux2,Lib2 : String ;
BEGIN
Gen:=BourreEtLess(Gen,fbGene) ;
Q:=OpenSQL('SELECT * FROM CORRESP WHERE CR_TYPE="GOR" AND CR_CORRESP="'+Gen+'" ',TRUE) ;
If Not Q.Eof Then
  BEGIN
  Cpt2:=Q.FindField('CR_LIBRETEXTE1').AsString ;
  Aux2:=Q.FindField('CR_LIBRETEXTE2').AsString ;
  Lib2:=Q.FindField('CR_LIBELLE').AsString ;
  If Not OkColl THen Aux:=Aux2 ;
  Gen:=Cpt2 ;
  END ;
Ferme(Q) ;
If (Not OkColl) And (Aux2<>'') And (Not coll416) Then
  BEGIN
  Aux2:=BourreEtLess(Aux2,fbAux) ; If Lib2='' Then Lib2:=Aux2 ;
  Q:=OpenSQL('SELECT * FROM TIERS WHERE T_AUXILIAIRE="'+Aux2+'"',FALSE) ;
  If Q.Eof Then
    BEGIN
    Q.Insert ;
    InitNew(Q) ;
    Q.FindField('T_AUXILIAIRE').AsString:=Aux2 ;
    Q.FindField('T_COLLECTIF').AsString:=BourreEtLess('4550',fbGene) ;
    Q.FindField('T_NATUREAUXI').AsString:='DIV' ;
    Q.FindField('T_LIBELLE').AsString:=Lib2 ;
    Q.FindField('T_ABREGE').AsString:=Copy(Lib2,1,17) ;
    Q.FindField('T_REGIMETVA').AsString:='001' ;
    Q.FindField('T_MODEREGLE').AsString:='999' ;
    Q.Post ;
    END ;
  Ferme(Q) ;
  END ;

END ;

procedure VerifCptGenCEGID(Gen : String) ;
Var Q : TQuery ;
    Gen2 : String ;
BEGIN
If Trim(Gen)='' Then Exit ;
Gen2:=BourreEtLess(Gen,fbGene) ;
Q:=OpenSQL('SELECT * FROM GENERAUX WHERE G_GENERAL="'+Gen2+'"',FALSE) ;
If Q.Eof Then
  BEGIN
  Q.Insert ;
  InitNew(Q) ;
  Q.FindField('G_GENERAL').AsString:=Gen2 ;
  Q.FindField('G_LIBELLE').AsString:='COMPTE CREE PAR IMPORT' ;
  Q.FindField('G_ABREGE').AsString:='IMPORT' ;
  If Gen2[1]='6' Then Q.FindField('G_NATUREGENE').AsString:='CHA' Else
    If Gen2[1]='7' Then Q.FindField('G_NATUREGENE').AsString:='PRO' Else
      Q.FindField('G_NATUREGENE').AsString:='DIV' ;
  Q.FindField('G_CREERPAR').AsString:='IMP' ;
  Q.FindField('G_SENS').AsString:='M' ;
  Q.FindField('G_COLLECTIF').AsString:='-' ;
  Q.Post ;
  END ;
Ferme(Q) ;
END ;

// ajout me 12-02-02
Function OkAuxiliaire (CptSISCO : TTabCptCollSISCO; cpt : string) : Boolean;
var
i        : integer;
EstAux   : Boolean;
Cptsav   : string;
begin
EstAux := FALSE;
For i:=0 To 50 Do
  BEGIN
       if CptSISCO[i].Cpt='' Then break ;
       // ajout me 02-04-02
       CptSISCO[i].Aux1 := BourreOuTronque(CptSISCO[i].Aux1,fbAux);
       CptSISCO[i].Aux2 := BourreOuTronque(CptSISCO[i].Aux2,fbAux);
       Cptsav := BourreOuTronque(Cpt,fbAux);
       if (Cptsav >= CptSISCO[i].Aux1)  and (CptSISCO[i].Aux2 >= Cptsav) then
       begin
          EstAux := TRUE; break;
       end;
  END ;
  Result := EstAux;
end;

Procedure QuelCompte(Cpt : String ; Var CptSISCO : TTabCptCollSISCO ; Var Gen,Aux : String) ;
Var EstAux : Boolean ;
    Nat : String ;
    PasTouche : Boolean ;
    OkColl,coll416 : Boolean ;
    i              : integer;
BEGIN
Gen:='' ; Aux:='' ;
Cpt:=AnsiUpperCase(Cpt) ;
If VH^.RecupPCL Then EstAux:=Cpt[1] In ['0','9','a'..'z','A'..'Z']
                Else EstAux:=Cpt[1] In ['0','9','C','F'] ;
if VH^.RecupComSx and (not EstAux) then // ajout me 12-02-02
   EstAux := OkAuxiliaire (CptSISCO, Cpt);

{$IFDEF CEGID}
Gen:=TrouveCptSISCO(Cpt,CptSISCO,Nat) ;
PasTouche:=FALSE ; OkColl:=FALSE ; coll416:=FALSE ;
If Gen<>'' Then
  BEGIN
  OkColl:=OkCollCegid(Gen,coll416) ;
  If OkColl Then OkColl:=PasOkCollCegid(Gen) ;
  If OkColl Then
    BEGIN
    Aux:=AnsiUpperCase(Cpt) ;
    If coll416 Then delete(Aux,2,2) ;
    END Else
    BEGIN
    Gen:=AnsiUpperCase(Cpt) ;
    END ;
  END Else Gen:=AnsiUpperCase(Cpt) ;
TraiteCptCegid(Gen,Aux,OkColl,coll416) ;
If (Aux='') And (Gen<>'') Then VerifCptGenCEGID(Gen) ;
{$ELSE}
If Not EstAux Then
  BEGIN
//  If VH^.RecupPCL And (Not VH^.RecupSISCOPGI)Then
    If (VH^.RecupPCL or VH^.RecupSISCOPGI ) Then
    BEGIN
    If (Copy(Cpt,1,2)='42') Or (Copy(Cpt,1,2)='43') Or (Copy(Cpt,1,2)='46') Then
      BEGIN
      Gen:=TrouveCptSISCO(Cpt,CptSISCO,Nat) ;
      { CA - 04/06/2003 - FQ 12163 - Astuce (!!!) pour le cas où la paie à déjà créé un collectif salarié,
                          on remonte alors les écritures du collectif sur le salarié d'attente }
      if (Gen='') and (BourreEtLess(Cpt,fbGene) = VH^.DefautSal) then
        Aux := VH^.TiersDefSAl;
      If Gen<>'' Then Aux:=AnsiUpperCase(Cpt) Else Gen:=AnsiUpperCase(Cpt) ;
      END Else Gen:=Cpt ;
    END Else Gen:=Cpt ;
  END Else
  BEGIN
{$IFDEF NOCONNECT}
Exit ;
{$ENDIF}
  Gen:=TrouveCptSISCO(Cpt,CptSISCO,Nat) ;
  If Gen<>'' Then Aux:=AnsiUpperCase(Cpt) Else Gen:=AnsiUpperCase(Cpt) ;
  END ;
{$ENDIF}
If VH^.RecupLTL Then
  BEGIN
  If Trim(Gen)<>'' Then Gen:=BourreOuTronque(Gen,fbGene) ;
  If Trim(Aux)<>'' Then Aux:=BourreOuTronque(Aux,fbAux) ;
  END ;
END ;

Function RetoucheListePiece(ListePiece : TStringList ; Var IdentPiece : TIdentPiece ; InfoImp : PtTInfoImport = Nil) : Integer ;
Var i : Integer ;
    St1,StMontant,Sens,LibelCGE : String ;
    TD,TC,M,Solde : Double ;
    OkOk,OkOk1 : Boolean ;
    OkMajTP,ODDetecte : Boolean ;
    NatP,NatPGlobal : String ;
    YY,MM,JJ : Word ;
BEGIN
Result:=0 ;
{ Tope les lignes sur compte général sans section indiquée }
{$IFDEF NOCONNECT}
Exit ;
{$ENDIF}
TD:=0 ; TC:=0 ; OkMajTP:=FALSE ; NatPGlobal:='' ;
ODDetecte:=FALSE ;
For i:=0 To ListePiece.Count-1 Do
  BEGIN
  OkOk:=TRUE ;
  (*
  If i<=ListePiece.Count-2 Then
    BEGIN
    If (Copy(ListePiece[i],31+DecalListe,1)=' ') And (Copy(ListePiece[i+1],31+DecalListe,1)='H') Then
      BEGIN
      St:=ListePiece[i] ; St[31+DecalListe]:='#' ; ListePiece[i]:=St ; OkOk:=FALSE ;
      END ;
    END ;
  *)
  NatP:=Trim(Copy(ListePiece[i],12+DecalListe,2)) ;
  If NatP='OD' Then ODDetecte:=TRUE ;
  If (NatP='RC') Or (NatP='RF') Or (NatPGlobal<>'') Then
    BEGIN
    If NatPGlobal='' Then NatPGlobal:=NatP ;
    If (NatP<>NatPGlobal) Then
      BEGIN
      OkMajTP:=TRUE ;
      If ((NatP='RC') And (NatPGlobal='RF')) Or ((NatP='RF') And (NatPGlobal='RC')) Then NatPGlobal:='OD' ;
      END Else If ODDetecte Then OkMajTP:=TRUE ;
    END ;
  If Copy(ListePiece[i],31+DecalListe,1)='A' Then OkOk:=FALSE ;
  If OkOk Then
    BEGIN
    M:=StrToFloat(StPoint(Copy(ListePiece[i],131+DecalListe,20))) ;
    If Copy(ListePiece[i],130+DecalListe,1)='D' Then TD:=Arrondi(TD+M,DecimSISCO(IdentPiece,FALSE)) Else TC:=Arrondi(TC+M,DecimSISCO(IdentPiece,FALSE)) ;
    END ;
  END ;
{ Recalcul de l'équilibrage }
Solde:=Arrondi(TD-TC,DecimSISCO(IdentPiece,FALSE)) ;
If Solde<>0 Then
  BEGIN
  DecodeDate(IdentPiece.dateP,YY,MM,JJ) ; Result:=YY ;
  St1:=ListePiece[ListePiece.Count-1] ;
  St1:=Insere(St1,FormatFloat('00000000',IdentPiece.Chrono+1),1,DecalListe) ;
  If VH^.RecupSISCOPGI Then St1:=Insere(St1,Format_String(VH^.Cpta[fbGene].AxGenAttente,17),14+DecalListe,17)
                       Else St1:=Insere(St1,Format_String(VH^.Cpta[fbGene].Attente,17),14+DecalListe,17) ;
  St1:=Insere(St1,Format_String(' ',17),32+DecalListe,17) ;
  If VH^.RecupPCL Then LibelCGE:=Format_String('Ecart de conversion',35) Else
    If VH^.RecupSISCOPGI Then LibelCGE:=Format_String('Equilibrage bordereau',35) Else
       LibelCGE:=Format_String('Equilibrage au jour',35) ;
  St1:=Insere(St1,LibelCGE,84+DecalListe,35) ;
  St1[31+DecalListe]:=' ' ;
  For i:=32+DecalListe To 32+DecalListe+17-1 Do St1[i]:=' ' ;
  If VH^.RecupSISCOPGI Then
    BEGIN
    Sens:='D' ;
    Solde:=-Solde ;
    END Else
    BEGIN
    If TD>TC Then Sens:='C' Else Sens:='D' ;
    Solde:=Abs(Solde) ;
    END ;
  St1[130+DecalListe]:=Sens[1] ;
  StMontant:=AlignDroite(StrfMontant(Solde,20,DecimSISCO(IdentPiece,FALSE),'',False),20) ;
  St1:=Insere(St1,StMontant,131+DecalListe,20) ;
  If VH^.RecupPCL Then
    BEGIN
    St1:=Insere(St1,Format_String(' ',20),176+DecalListe,20) ;
    END ;
  St1:=Insere(St1,Format_String(' ',2),219+DecalListe,2) ;
  St1:=Insere(St1,Format_String(' ',2),221+DecalListe,2) ;
  (*
  St1:=FormatFloat('00000000',IdentPiece.Chrono+1)+Jal+Date+TP+General+TC+AuxSect+Reference+
     Libelle+MP+Echeance+S+Montant+TE+NumP+Dev+TauxDev+CodeMontant+Montant2+Montant3+Etab+Axe+NumE ;
  *)
  {$IFNDEF PCL}
  OkOk1:=TRUE ;
  If InfoImp<>NIl Then OkOk1:=EstJalARecuperer(IdentPiece.JalP,InfoImp^.JalFaux) ;
  If OkOk1 Then ListePiece.Add(St1) ;
  {$ENDIF}
  END ;
If VH^.RecupSISCOPGI Then OkMajTP:=FALSE ;
If OkMajTP And ((NatPGlobal='RC') Or (NatPGlobal='RF') Or (NatPGlobal='OD'))Then
  BEGIN
  For i:=0 To ListePiece.Count-1 Do
    BEGIN
    St1:=ListePiece[i] ;
    St1:=Insere(St1,NatPGlobal,12+DecalListe,2) ;
    LIstePiece[i]:=St1 ;
    END ;
  END ;
(*
For i:=0 To ListePiece.Count-1 Do
  BEGIN
  StDate:=Copy(ListePiece[i],8+DecalListe,4)+Copy(ListePiece[i],6++DecalListe,2)+Copy(ListePiece[i],4+DecalListe,2)+
  Insere(ListePiece[i],StDate,1,DecalListe) ;
  END ;
ListePiece.Sorted:=TRUE ; ListePiece.Duplicates:=dupAccept ;
*)
END ;

Function EcritPiece(Var NewFichier : TextFile ; ListePiece : TStringList ; Var IdentPiece : TIdentPiece ; Var Infos : TInfoGeneSISCO ;
                    Var PieceEcrite : Boolean ;
                    InfoImp : PtTInfoImport = Nil) : Integer ;
Var LeNum : Integer ;
    i : Integer ;
    St :String ;
    LaDate : TDateTime ;
    LeJal,LeQual,LeFolio : String ;
    LeChrono : Integer ;
    LeNoFolio : Integer ;
    LeNumChronoEclate : Integer ;
BEGIN
PieceEcrite:=ListePiece.Count>0 ;
Result:=RetoucheListePiece(ListePiece,IdentPiece,InfoImp) ;
for i:=0 to ListePiece.Count-1 do
  BEGIN
  St:=ListePiece[i] ; St:=Copy(St,1+DecalListe,Length(St)-DecalListe) ;
  If St[31]<>'#' Then WriteLn(NewFichier,St) ;
  END ;
VideStringList(ListePiece) ;
ListePiece.Sorted:=FALSE ;
LeNum:=IdentPiece.NumP ; Inc(LeNum) ; LaDate:=IdentPiece.DateP ;
LeJal:=IdentPiece.JalP ; LeQual:=IdentPiece.QualP ; LeFolio:=IdentPiece.FolioSISCO ;
LeNoFolio:=IdentPiece.NoFolioSISCO ;
LeNumChronoEclate:=IdentPiece.NumChronoEclate ;
LeChrono:=IdentPiece.Chrono ;
Fillchar(IdentPiece,SizeOf(IdentPiece),#0) ;
IdentPiece.NumP:=LeNum ; IdentPiece.DateP:=LaDate ;
IdentPiece.JalP:=LeJal ; IdentPiece.QualP:=LeQual ;
IdentPiece.Chrono:=LeChrono ; IdentPiece.FolioSISCO:=LeFolio ;
IdentPiece.NoFolioSISCO:=LeNoFolio ;
IdentPiece.NumChronoEclate:=LeNumChronoEclate ;
END ;

Function RecupMois(St,StSup : String) : TDateTime ;
Var St1 : String ;
    DD1,DD2 : TDateTime ;
    Err : Integer ;
    Exo : TExoDate ;
BEGIN
St1:='M'+Copy(St,2,2)+StSup ;
WhatDate(St1,DD1,DD2,Err,Exo) ;
Result:=DD1 ;
END ;

Function RecupMontant(St : String ; Deb,Long,Decim : Integer) : Double ;
Var St1,St2 : String ;
BEGIN
Result:=0 ;
St1:=Trim(Copy(St,Deb,Long)) ;
If St1<>'' Then
  BEGIN
  If Length(St1)<3 Then St1:='00'+St1 ;
  St2:=Copy(St1,1,Length(St1)-Decim)+DecimalSeparator+Copy(St1,Length(St1)-Decim+1,Decim) ;
  Result:=Arrondi(Valeur(St2),Decim) ;
  END ;
END ;

function DecodeIdentEcr(IdentEcr:string) : string ;
var
  Num, NumPiece : integer ;
begin
    Result := '';
    NumPiece:=0 ;
    if IsNumeric(IdentEcr) then
    begin
      Num:=StrToInt(Identecr) ;
      NumPiece:=num div 1000 ;
      Result := Format_String(IntToStr(NumPiece), 17)
    end;
end ;

Procedure FAitStSISCO(Var St : String ; Var StSISCO,OldStSISCO : TStSISCO ;
                      Var IdentPiece : TIdentPiece ; Var Infos : TInfoGeneSISCO ;
                      RuptOrigineSaisie : Boolean ; CodeMontantSISCO : String ;
                      InfoImp : PtTInfoImport = Nil) ;
Var Y,M,D : Word ;
    StY,StM,StJ,LeSens,Lettre : String ;
    MontantAnal : Double ;
    Jour,NCE : Integer ;
    CptLuJ : TCptLu ;
    ll : Integer ;
    MonoExo : Boolean ;
    Exo : TExoDate ;
    ModifStSISCO : Boolean ;
    OldStDate : String ;
    NatG : String ;
    SQte : String ;
    i : Integer ;
    Qte : Double ;
    OkQte : Boolean ;
    CptOrigineCegid,CodeStat : String ;
    SPointe,SNumChq : String ;
    St1,St2 : String ;
    DecalAna : Integer ;
BEGIN
If (InfoImp<>Nil) And (Not InfoImp^.RecupODA) And ((St[1] in ['G','w'])=TRUE) Then Exit ;
DecodeDate(IdentPiece.DateP,Y,M,D) ; ModifStSISCO:=FALSE ;
StY:=FormatFloat('0000',Y) ; StM:=FormatFloat('00',M) ;
StSISCO.RefLibre:='' ; StSISCO.Qte:='' ; StSISCO.Affaire:='' ; CptOrigineCegid:='' ;
Case St[1] Of
  'E','G' : BEGIN
        If St[1]='G' Then DecalAna:=40 Else DecalAna:=0 ;
        StSISCO.Jal        :=AnsiUpperCase(Format_String(IdentPiece.JalP,3)) ;
        StSISCO.Jal:=CptRemplace(StSISCO.Jal,Infos.CharRJ) ;
        Jour:=StrToInt(Copy(St,2,2)) ;
        StSISCO.Date       :=FormatFloat('00',Jour)+StM+StY ;
        StSISCO.TP         :='OD' ;
        Fillchar(CptLuJ,SizeOf(CptLuJ),#0) ; CptLuJ.Cpt:=Trim(StSISCO.Jal) ;
//        If ChercheCptLu(Infos.LJalLu,CptLuJ) Then
{$IFNDEF NOCONNECT}

        If AlimLTabCptLu(3,Infos.QFiche[3],Infos.LJalLu,NIL,CptLuJ) Then
          BEGIN
          If CptLuJ.Nature='ACH' Then StSISCO.TP:='FF' ;
          If CptLuJ.Nature='VTE' Then StSISCO.TP:='FC' ;
          If CptLuJ.Nature='ANO' Then
            BEGIN
            If QuelExoDate(IdentPiece.DateP,IdentPiece.DateP,MonoExo,Exo) And (Exo.Code=Infos.ExoSISCO.Code) Then
              BEGIN
              DecodeDate(Exo.Deb,Y,M,D) ;
              StY:=FormatFloat('0000',Y) ; StM:=FormatFloat('00',M) ; StJ:=FormatFloat('00',D) ;
              OldStDate:=StSISCO.Date ;
              StSISCO.Date:=StJ+StM+StY ;
              ModifStSISCO:=TRUE ;
//              IdentPiece.DateP:=RecupMois(St,Infos.StPlusExo) ;
              END ;
            END ;
          END ;
{$ENDIF}
        CptOrigineCegid:=Copy(St,4,10) ;
        QuelCompte(Copy(St,4,10),Infos.FourchetteSISCO,StSISCO.General,StSISCO.AuxSect) ;
        if VH^.RecupComSx then // ajout me 20-03-02
        begin
           if Infos.Troncaux then
           begin
                if CptOrigineCegid[1] In ['0','9','C','F'] then StSISCO.AuxSect := BourreOuTronque(Copy(StSISCO.AuxSect,2,Length(StSISCO.AuxSect)-1),fbaux);
           end
           else StSISCO.AuxSect := BourreOuTronque(StSISCO.AuxSect,fbaux);
           StSISCO.General := BourreOuTronque(StSISCO.General,fbGene);
        end;
        IdentPiece.LigneEnCours.Gen:=StSISCO.General ;
        StSISCO.General    :=Format_String(StSISCO.General,17) ;
        StSISCO.General    :=CptRemplace(StSISCO.General,Infos.CharRG) ;
{$IFNDEF NOCONNECT}
        If CptLuJ.Nature='BQE' Then
          BEGIN
          NatG:=NatureCptImport(Trim(StSISCO.General),Infos.FourchetteImport) ;
          If NatG='---' Then NatG:=GetNatureDiv(Trim(StSISCO.General), Infos.FourchetteSISCO) ;
          If NatG='COC' Then StSISCO.TP:='RC' Else
            If NatG='COF' Then StSISCO.TP:='RF' Else If VH^.RecupSISCOPGI Then
              BEGIN
              St1:=CptLuJ.CollTiers ;
              St2:=BourreOuTronque(StSISCO.General,fbGene) ;
              If St1=St2 Then
                BEGIN
                If IdentPiece.LigneEnCours.DP=0 Then StSISCO.TP:='RF' Else StSISCO.TP:='RC' ;
                END Else
                BEGIN
                StSISCO.TP:='OD' ;
                END ;
              END ;
          END ;
{$ENDIF}
        StSISCO.AuxSect    :=Format_String(StSISCO.AuxSect,17) ;
        StSISCO.AuxSect    :=CptRemplace(StSISCO.AuxSect,Infos.CharRT) ;
        StSISCO.TC         := ' ' ;
        If Trim(StSISCO.AuxSect)<>'' Then StSISCO.TC:='X' ;
        StSISCO.Reference  :=Format_String(Copy(St,34,10),35) ;
        StSISCO.Libelle    :=Format_String(Copy(St,14,20),35) ;
        StSISCO.MP         :=AnsiUpperCase(Format_String(Copy(St,82,2),3)) ;
        StSISCO.Echeance   :=Trim(Copy(St,76,6)) ;
        If StSISCO.Echeance='' Then StSISCO.Echeance:=Format_String(' ',8) Else
          BEGIN
          If StrToInt(Copy(StSISCO.Echeance,5,2))<90 Then StSISCO.Echeance:=Copy(StSISCO.Echeance,1,4)+'20'+Copy(StSISCO.Echeance,5,2)
                                                     Else StSISCO.Echeance:=Copy(StSISCO.Echeance,1,4)+'19'+Copy(StSISCO.Echeance,5,2) ;
          END ;
        StSISCO.DateRefExterne   :=Trim(Copy(St,92,6)) ;
        If StSISCO.DateRefExterne='' Then StSISCO.DateRefExterne:=Format_String(' ',8) Else
          BEGIN
          If StrToInt(Copy(StSISCO.DateRefExterne,5,2))<90 Then StSISCO.DateRefExterne:=Copy(StSISCO.DateRefExterne,1,4)+'20'+Copy(StSISCO.DateRefExterne,5,2)
                                                           Else StSISCO.DateRefExterne:=Copy(StSISCO.DateRefExterne,1,4)+'19'+Copy(StSISCO.DateRefExterne,5,2) ;
          END ;
        StSISCO.S          :='D' ;
        If IdentPiece.LigneEnCours.DP=0 Then StSISCO.S:='C' ;
        StSISCO.Montant    :=AlignDroite(StrfMontant(IdentPiece.LigneEnCours.DP+IdentPiece.LigneEnCours.CP,20,DecimSISCO(IdentPiece,FALSE),'',False),20) ;
        StSISCO.TE         :='N' ;
        If VH^.RecupPCL Then
          BEGIN
          If RuptOrigineSaisie And OkCROS And (CodeMontantSISCO<>IdentPiece.CodeMontant) And (CodeMontantSISCO<>'') Then
            BEGIN
            NCE:=1000*IdentPiece.NumChronoEclate+IdentPiece.NoFolioSISCO ;
            StSISCO.NumP       :=FormatFloat('00000000',NCE) ;
            END Else StSISCO.NumP       :=FormatFloat('00000000',IdentPiece.NoFolioSISCO) ;
          END Else StSISCO.NumP       :=FormatFloat('00000000',IdentPiece.NumP) ;
        StSISCO.Dev        :=Format_String(' ',3) ;
        StSISCO.TauxDev    :=Format_String(' ',10) ;
        StSISCO.CodeMontant:=IdentPiece.CodeMontant ;
        If VH^.RecupPCL And (Arrondi(IdentPiece.LigneEnCours.DE+IdentPiece.LigneEnCours.CE,2)<>0)
           Then StSISCO.Montant2   :=AlignDroite(StrfMontant(IdentPiece.LigneEnCours.DE+IdentPiece.LigneEnCours.CE,20,DecimSISCO(IdentPiece,FALSE),'',False),20) 
           Else StSISCO.Montant2   :=Format_String(' ',20) ;
        StSISCO.Montant3   :=Format_String(' ',20) ;
        If (VH^.RecupCegid) And (InfoImp<>NIL) Then
          BEGIN
          StSISCO.Etab:=Format_String(InfoImp^.BaseOrigine,3) ;      
          END Else StSISCO.Etab       :=Format_String(' ',3) ;
        StSISCO.Axe        :=Format_String(' ',2) ;
        StSISCO.NumE       :=Format_String(' ',2) ;
        StSISCO.RefLibre   :=Format_String(' ',35) ;
        StSISCO.Affaire   :=Format_String(' ',17) ;
        StSISCO.Qte        :=Format_String(' ',20) ;
        StSisco.RefP:='' ;
        If (VH^.RecupSISCOPGI or VH^.RecupPCL) and (DecalAna=0) Then
          BEGIN
          SPointe:=Trim(Copy(St,98,1)) ; SNumChq:=Trim(Copy(St,99,10)) ;
          If SPointe='' Then SPointe:='-' Else SPointe:='X' ;
          StSisco.RefP:=SPointe+';'+SNumChq+';' ;
          END ;
        // ajout me 16-12-2002
        if VH^.RecupComSx then
        begin
            Lettre :=Trim(Copy(St,44,1)) ;
            If (Lettre<>'') And (Lettre<>'*') Then
            begin
                 if Trim(Copy(St,138,1)) <> '' then
                    StSISCO.lettre := 'AA'+Trim(Copy(St,138,1))+Lettre
                 else
                    StSISCO.lettre := 'AAA'+Lettre;
                 StSISCO.lettre := Uppercase(StSISCO.lettre);
            end;
            if (Trim(Copy(St,138,1)) <> '') AND (Trim(Copy(St,153,8)) <> '00000000')
            and (Trim(Copy(St,153,8)) <> '')  then
               StSISCO.datepaquetmin       :=Copy(St,153,8)
            else
               StSISCO.datepaquetmin := '01011900';
        end;
                           // ajout me 10/07/2003 pour les codes stat
        If RecupSISCO or (VH^.RecupComSx and (InfoImp <> nil)) Then
        BEGIN
             ll:=Pos('&[+',St) ;
             if ll<=0 Then StSISCO.RefExterne:=IdentPiece.FolioSISCO+';' Else // Recup N° FOLIO SISCO
             BEGIN
                  StSISCO.RefExterne:=IdentPiece.FolioSISCO+' Ligne '+Copy(St,ll+3,2)+';' ;
             END ;
             If (IdentPiece.DateP>=VH^.EnCours.Deb) Then
             BEGIN
                  Lettre:=Trim(Copy(St,44,1)) ;
                  If (Lettre<>'') And (Lettre<>'*') And (DecalAna=0) Then
                    BEGIN
                    If VH^.RecupLTL Or VH^.RecupSISCOPGI Then StSISCO.RefReleve:=Format_String(Lettre+CptOrigineCEGID,10) ;
                    StSISCO.RefExterne:=Format_String(StSISCO.RefExterne+'AAAA;',35) ;
                    END Else StSISCO.RefExterne:=Format_String(StSISCO.RefExterne,35) ;
             END ;
             StSISCO.RefLibre:=Format_String(Copy(St,72,4),35) ;
             If ((InfoImp<>NIL) And (VH^.RecupSISCOPGI or VH^.RecupPCL) And (DecalAna=0))
             // ajout me 10/07/2003 pour les codes stat
           {JP 01/09/03 : or au lieu de and sinon, lorsque l'on n'est en ComSx, il n'est pas possible de récupérer les codes stats}
               or (VH^.RecupComSx and (InfoImp <> nil)) Then
             BEGIN
                  CodeStat:=Trim(Copy(St,72,4)) ;
                  If CodeStat<>'' Then InfoImp^.LCS.Add(CodeStat) ;
             END ;
             If VH^.RecupCEGID Or VH^.RecupLTL Then StSISCO.Affaire:=Format_String(CptOrigineCEGID,17)
             // Dans le cas de la récup. PCL, on récupère l'identifiant écriture dans le code affaire
             // cet indentifiant permet de récupérer le numéro de pièce de S1 dans le cas
             // d'un synchro S1-Sisco.
             else if VH^.RecupPCL or VH^.RecupComSx then
             begin
                if Copy (St,153,8)<>'00000000' then
                begin
//                  StSISCO.Affaire := Format_String(Copy (St, 153, 8) ,17);
                    // ajout me 27-01-2005 identification de la pièce s1
                    St1 := Trim(Copy(St, 153, 8));
                    if St1 <> '' then StSISCO.Affaire := DecodeIdentEcr(St1);
                end;
             end;
             If DecalAna=0 Then SQte:=Copy(St,84,8) Else SQte:='' ;
             If Trim(SQte)<>'' Then
             BEGIN
                  OkQte:=TRUE ;
                  For i:= 1 To Length(SQte) Do If (SQte[i] in ['0'..'9','.',',','-']) = FALSE Then OkQte:=FALSE ;
                  If OkQte Then
                     BEGIN
                     Qte:=RecupMontant(St,84,8,2) ;
                     If Arrondi(Qte,2)<>0 Then StSisco.Qte:=AlignDroite(StrfMontant(Qte,20,4,'',False),20) ;
                     END ;
             END ;
           END
           ELSE   // ajout me 27-01-2005 identification de la pièce s1
           if VH^.RecupComSx then
           begin
                if Copy (St,153,8)<>'00000000' then
                begin
                    St1 := Trim(Copy(St, 153, 8));
                    if St1 <> '' then StSISCO.Affaire := DecodeIdentEcr(St1);
                end;
           end;

        OldStSISCO:=StSISCO ;
        If ModifStSISCO Then OldStSISCO.Date:=OldStDate ;

        END ;
  'v','w' : BEGIN
{$IFDEF NOCONNECT}
        Exit ;
{$ENDIF}
        CptOrigineCegid:=Copy(St,3,15) ;
        StSISCO.Jal        :=OldStSISCO.Jal ;
        StSISCO.Date       :=OldStSISCO.Date ;
        StSISCO.TP         :=OldStSISCO.TP ;
        StSISCO.General    :=OldStSISCO.General ;
        StSISCO.AuxSect    :=Format_String(Copy(St,3,15),17) ;
        StSISCO.AuxSect:=CptRemplace(StSISCO.AuxSect,Infos.CharRS) ;
//        StSISCO.TC         :='H' ;
        StSISCO.TC         :='A' ;
        StSISCO.Reference  :=OldStSISCO.Reference ;
        StSISCO.Libelle    :=Format_String(Copy(St,39,20),35) ;
        StSISCO.MP         :=Format_String(' ',3) ;
        StSISCO.Echeance   :=Format_String(' ',8);
        StSISCO.S          :=OldStSISCO.S ;
        LeSens:=Copy(St,38,1) ;
//        MontantAnal:=RecupMontant(St,18,20,DecimSISCO(IdentPiece,FALSE)) ;
        MontantAnal:=IdentPiece.LigneEnCours.DP+IdentPiece.LigneEnCours.CP ;
        If LeSens<>OldStSISCO.S Then MontantAnal:=-1*MontantAnal ;
        StSISCO.Montant    :=AlignDroite(StrfMontant(MontantAnal,20,DecimSISCO(IdentPiece,FALSE),'',False),20) ;
        StSISCO.TE         :=OldStSISCO.TE ;
        StSISCO.NumP       :=OldStSISCO.NumP ;
        StSISCO.Dev        :=OldStSISCO.Dev ;
        StSISCO.TauxDev    :=OldStSISCO.TauxDev ;
        StSISCO.CodeMontant:=OldStSISCO.CodeMontant ;
        StSISCO.Montant2   :=Format_String(' ',20) ;
        StSISCO.Montant3   :=Format_String(' ',20) ;
        StSISCO.Etab       :=OldStSISCO.Etab ;
        StSISCO.Axe        :='A1' ;
        StSISCO.NumE       :=Format_String(' ',2) ;
        StSISCO.RefExterne :=OldStSISCO.RefExterne ;
        StSISCO.RefLibre   :=Format_String(' ',35) ;
        StSISCO.Affaire    :=Format_String(' ',17) ;
        StSISCO.Qte        :=Format_String(' ',20) ;
        stSISCO.DateRefExterne:=OldStSISCO.DateRefExterne ;
        if VH^.RecupCegid Or VH^.RecupLTL Then StSisco.Affaire:=Format_String(CptOrigineCegid,17) ;
        END ;
  END ;
END ;

{$IFDEF RECUPPCL}
Function EstUneLigneEcart(Cpt : String ; InfoImp : PtTInfoImport = Nil) : Boolean ;
Var CptD,CptC : String ;
BEGIN
Result:=FALSE ;
Cpt:=BourreOuTronque(Cpt,fbGene) ;
If VH^.RecupCegid Then
  BEGIN
  CptD:=InfoImp^.CptEccSISCO  ; CptC:=InfoImp^.CptEccSISCO  ;
  END Else
  BEGIN
  CptD:=GetParamSoc('SO_ECCEURODEBIT') ;
  CptC:=GetParamSoc('SO_ECCEUROCREDIT') ;
  END ;
CptD:=BourreOuTronque(CptD,fbGene) ;
CptC:=BourreOuTronque(CptC,fbGene) ;
If (Cpt=CptD) Or (Cpt=CptC) Then Result:=TRUE ;
END ;
{$ENDIF}

Procedure MAJListeSISCO(Var StSISCO : TStSISCO ; ListePiece : TStringList ;
                        Var IdentPiece : TIdentPiece ; InfoImp : PtTInfoImport = Nil) ;
Var St1 : String ;
    OkLigne,OkOk : Boolean ;
    i : Integer ;
    AffaireCEGID : String ;
    LaRefReleve : String ;
    StSup : String ;
BEGIN
{$IFDEF NOCONNECT}
If Length(StSisco.General)<3 Then Exit ;
If (Copy(StSisco.General,1,3)<>'512') And (Copy(StSisco.General,1,3)<>'514') Then Exit ;
{$ENDIF}
With StSISCO Do
  BEGIN
  If VH^.RecupCEGID And (InfoImp<>NIL) And (InfoImp^.BaseOrigine='CLI') Then
    BEGIN
    i:=StrToInt(NumP) ; i:=1000+i ; NumP:=FormatFloat('00000000',i) ;
    END ;
  St1:=FormatFloat('00000000',IdentPiece.Chrono)+Jal+Date+TP+General+TC+AuxSect+Reference+
       Libelle+MP+Echeance+S+Montant+TE+NumP+Dev+TauxDev+CodeMontant+Montant2+Montant3+Etab+Axe+NumE ;
  If RecupSISCO Then
    BEGIN
    AffaireCegid:=Format_String(' ',17) ;
    LaRefReleve:=Format_String(' ',10) ;
    If VH^.RecupCegid Or VH^.RecupLTL Or VH^.RecupSISCOPGI Then AffaireCegid:=Affaire ;
    If VH^.RecupLTL Or VH^.RecupSISCOPGI Then LaRefReleve:=StSISCO.RefReleve ;
    St1:=St1+RefExterne ;
    StSup:='' ;
    If VH^.RecupLTL Or VH^.RecupSISCOPGI Then
      StSup:=Format_String(' ',1)+Format_String(' ',3)+Format_String(' ',3)+
             Format_String(' ',3)+Format_String(' ',17)+Format_String(' ',17)+
             Format_String(RefP,17)+Format_String(' ',8)+Format_String(' ',8)+
             Format_String(' ',8)+Format_String(' ',35)+LaRefReleve ;
    If VH^.RecupCegid Then
       BEGIN
       St1:=St1+DateRefExterne+Format_String(' ',8)+Format_String(' ',3)+AffaireCegid
               +Format_String(' ',8)+Format_String(' ',3)
               +Qte
               +Format_String(' ',20)+Format_String(' ',3)+Format_String(' ',3)
               +RefLibre+StSup ;
       END else
       BEGIN
        // Mise à jour du code affaire (identifiant écriture Sisco II
        if VH^.RecupPCL then AffaireCegid := Affaire;
        St1:=St1+Format_String(' ',8)+Format_String(' ',8)+Format_String(' ',3)+AffaireCegid
               +Format_String(' ',8)+Format_String(' ',3)
               +Qte
               +Format_String(' ',20)+Format_String(' ',3)+Format_String(' ',3)
               +RefLibre+StSup ;
       END ;
    END
    ELSE
    BEGIN
         if VH^.RecupComSx then
         begin
              // ajout me 27-01-2005 pour sauvegarder la pièce S1 dans affaire
              if (StSISCO.lettre <> '')  then
              St1:=St1+Format_String(' ',35)+StSISCO.DateRefExterne+
              Format_String(' ',8)+Format_String(' ',3)+Affaire +
              Format_String(' ',704)
              //Format_String(' ',732)
              +StSISCO.datepaquetmin+ '01011900'+
              StSISCO.lettre + ' --TL'+StSISCO.RefLibre
              else
              if (StSISCO.RefLibre <> '')  then
              St1:=St1+Format_String(' ',35)+StSISCO.DateRefExterne+
              Format_String(' ',8)+Format_String(' ',3)+Affaire +
              Format_String(' ',704)
              //Format_String(' ',732)
              +StSISCO.datepaquetmin+ '01011900'+
              Format_String(' ',10) + StSISCO.RefLibre;
         end;
    END;

  END ;
OkLigne:=TRUE ;
{$IFDEF RECUPPCL}
//OkLigne:=Not EstUneLigneEcart(StSisco.General,InfoImp)  ;
{$ENDIF}
If VH^.RecupSISCOPGI Then okLigne:=TRUE ;
OkOk:=TRUE ;
If InfoImp<>NIl Then OkOk:=EstJalARecuperer(StSisco.Jal,InfoImp^.JalFaux) ;
If OkLigne Then If OkOk Then ListePiece.Add(St1) ;
(*
E_MODESAISIE "BOR", "-" ou "LIB"
E_EQUILIBRE X ou - Ligne d'équilibrage générée
*)
END ;

Procedure TraiteMontantSISCOEcr(Var St : String ; Var IdentPiece : TIdentPiece ; Var Infos : TInfoGeneSISCO ; CodeMontantSISCO : String ;
                                RuptOrigineSaisie : Boolean = FALSE ; DecalAna : Integer = 0) ;
Var ModeOppose : Boolean ;
    D,C,DE,CE,M,DP,CP : Double ;
BEGIN
(*
With IdentPiece Do
  BEGIN
  ModeOppose:=(Infos.EnEuro And (CodeMontantSISCO='F')) Or ((Not Infos.EnEuro) And (CodeMontantSISCO='E')) ;
  D:=0 ; C:=0 ; DE:=0 ; CE:=0 ;
  If ModeOppose Then
    BEGIN
    M:=RecupMontant(St,140,13,DecimSISCO(IdentPiece,TRUE)) ;
    If M>0 Then D:=M Else If M<0 Then C:=Abs(M) Else
      BEGIN
      D:=RecupMontant(St,46,13,DecimSISCO(IdentPiece,FALSE)) ;
      C:=RecupMontant(St,59,13,DecimSISCO(IdentPiece,FALSE)) ;
      If IdentPiece.CodeMontant='F--' Then IdentPiece.CodeMontant:='E--' Else IdentPiece.CodeMontant:='F--' ;
      END ;
    END Else
    BEGIN
    D:=RecupMontant(St,46,13,DecimSISCO(IdentPiece,FALSE)) ;
    C:=RecupMontant(St,59,13,DecimSISCO(IdentPiece,FALSE)) ;
    END ;
  LigneEnCours.DP:=D ; LigneEnCours.DE:=DE ;
  LigneEnCours.CP:=C ; LigneEnCours.CE:=CE ;
  LigneEnCours.QualQ1:=CodeMontantSISCO ;
  TotDP:=TotDP+LigneEnCours.DP ; TotDE:=TotDE+LigneEnCours.DE ;
  TotCP:=TotCP+LigneEnCours.CP ; TotCE:=TotCE+LigneEnCours.CE ;
  END ;
*)
With IdentPiece Do
  BEGIN
  ModeOppose:=(Infos.EnEuro And (CodeMontantSISCO='F')) Or ((Not Infos.EnEuro) And (CodeMontantSISCO='E')) ;
  D:=0 ; C:=0 ; DE:=0 ; CE:=0 ;
  If ModeOppose Then
    BEGIN
    M:=RecupMontant(St,140,13,DecimSISCO(IdentPiece,TRUE)) ;
    DP:=RecupMontant(St,46,13,DecimSISCO(IdentPiece,FALSE)) ;
    CP:=RecupMontant(St,59,13,DecimSISCO(IdentPiece,FALSE)) ;
    If M>0 Then D:=M Else If M<0 Then C:=Abs(M) Else
      BEGIN
      D:=RecupMontant(St,46,13,DecimSISCO(IdentPiece,FALSE)) ;
      C:=RecupMontant(St,59,13,DecimSISCO(IdentPiece,FALSE)) ;
      DP:=D ; CP:=C ;
      If VH^.RecupPCL Then
        BEGIN
        DE:=D ; CE:=C ;
        If Infos.EnEuro Then
          BEGIN
          D:=D*6.55957 ; D:=Arrondi(D,2) ;
          C:=C*6.55957 ; C:=Arrondi(C,2) ;
          End Else
          BEGIN
          D:=D/6.55957 ; D:=Arrondi(D,2) ;
          C:=C/6.55957 ; C:=Arrondi(C,2) ;
          END ;
        END Else
        BEGIN
        If IdentPiece.CodeMontant='F--' Then IdentPiece.CodeMontant:='E--' Else IdentPiece.CodeMontant:='F--' ;
        END ;
      END ;
    END Else
    BEGIN
    D:=RecupMontant(St,46,13,DecimSISCO(IdentPiece,FALSE)) ;
    C:=RecupMontant(St,59,13,DecimSISCO(IdentPiece,FALSE)) ;
    DP:=D ; CP:=C ;
    END ;
  LigneEnCours.DP:=D ; LigneEnCours.DE:=DE ;
  LigneEnCours.CP:=C ; LigneEnCours.CE:=CE ;
  LigneEnCours.QualQ1:=CodeMontantSISCO ;
  If VH^.RecupPCL Then
    BEGIN
    If Arrondi(DE+CE,2)<>0 Then
      BEGIN
      TotDP:=TotDP+LigneEnCours.DE ; TotDE:=TotDE+LigneEnCours.DE ;
      TotCP:=TotCP+LigneEnCours.CE ; TotCE:=TotCE+LigneEnCours.CE ;
      END Else
      BEGIN
      TotDP:=TotDP+DP ; TotDE:=TotDE+LigneEnCours.DE ;
      TotCP:=TotCP+CP ; TotCE:=TotCE+LigneEnCours.CE ;
      END ;
    END Else
    BEGIN
    TotDP:=TotDP+LigneEnCours.DP ; TotDE:=TotDE+LigneEnCours.DE ;
    TotCP:=TotCP+LigneEnCours.CP ; TotCE:=TotCE+LigneEnCours.CE ;
    END ;
  END ;
END ;

Procedure TraiteMontantSISCOAna(Var St : String ; Var IdentPiece : TIdentPiece ; Var Infos : TInfoGeneSISCO ; CodeMontantSISCO : String ;
                                RuptOrigineSaisie : Boolean = FALSE) ;
Var ModeOppose : Boolean ;
    D,C,DE,CE,M : Double ;
BEGIN
With IdentPiece Do
  BEGIN
  ModeOppose:=(Infos.EnEuro And (CodeMontantSISCO='F')) Or ((Not Infos.EnEuro) And (CodeMontantSISCO='E')) ;
  D:=0 ; C:=0 ; DE:=0 ; CE:=0 ;
  If ModeOppose Then
    BEGIN
    M:=RecupMontant(St,86,13,DecimSISCO(IdentPiece,TRUE)) ;
    If M<>0 Then D:=M Else
      BEGIN
      D:=RecupMontant(St,18,20,DecimSISCO(IdentPiece,FALSE)) ;
      END ;
    END Else
    BEGIN
    D:=RecupMontant(St,18,20,DecimSISCO(IdentPiece,FALSE)) ;
    END ;
  LigneEnCours.DP:=D ; LigneEnCours.DE:=DE ;
  LigneEnCours.CP:=C ; LigneEnCours.CE:=CE ;
  END ;
END ;

Function EstJalAN(Jal : String ; RacineJal : String) : Boolean ;
Var St : String ;
BEGIN
Result:=TRUE ;
While RacineJal<>'' Do
  BEGIN
  St:=AnsiUpperCase(ReadTokenSt(RacineJal)) ;
  Jal:=AnsiUpperCase(Jal) ;
  If Trim(St)=Trim(Jal) Then Exit ;
  END ;
Result:=FALSE ;
END ;

Function FaitPiece(St : String ; ListePiece : TStringList ; Var IdentPiece : TIdentPiece ;
                    Var OldStSISCO : TStSISCO ; Var Infos : TInfoGeneSISCO ;
                    InfoImp : PtTInfoImport = Nil ; RuptOrigineSaisie : Boolean = FALSE) : Integer ;
Var StSISCO : TStSISCO ;
    OkOk : Boolean ;
    CodeMontantSISCO,OldCodeMontantSISCO : String ;
    Q : TQuery ;
    StSect,St1,St2 : String ;
    i : Integer ;
    NewJal : Boolean ;
    DecalAna : Integer ;

BEGIN
Result:=0 ;
If (InfoImp<>Nil) And (Not InfoImp^.RecupODA) And ((St[1] in ['G','w'])=TRUE) Then Exit ;
Case St[1] Of
  'M' : BEGIN
        IdentPiece.DateP:=RecupMois(St,Infos.StPlusExo) ;
        If IdentPiece.DateP>Infos.ExoSISCO.Fin Then Result:=1 ;
        If IdentPiece.DateP<Infos.ExoSISCO.Deb Then Result:=1 ;
        END ;
  'J' : BEGIN
        IdentPiece.JalP:=AnsiUpperCase(Trim(Copy(St,2,2))) ;
        If VH^.RecupCEGID And (InfoImp<>NIL) And (InfoImp.BaseOrigine='CLI') Then IdentPiece.JalP:='Z'+IdentPiece.JalP ;
        If VH^.RecupSISCOPGI  Then
          BEGIN
          NewJal:=TRUE ;
          If (InfoImp<>NIL) And (InfoImp^.RacineJalAN<>'') And (EstJalAN(IdentPiece.JalP,InfoImp^.RacineJalAN)) Then NewJal:=FALSE ;
          If NewJal Then IdentPiece.JalP:=IdentPiece.JalP+CharJal ;
          END ;
        IdentPiece.NatP:='OD' ; IdentPiece.QualP:='N' ;
        END ;
  'E','G' : BEGIN
        If St[1]='G' Then DecalAna:=40 Else DecalAna:=0 ;
        Inc(IdentPiece.Chrono) ;
        CodeMontantSISCO:=Trim(Copy(St,139-DecalAna,1)) ;
        If (VH^.RecupPCL) And RuptOrigineSaisie And (IdentPiece.codeMontant<>'') And (Not OkCROS) Then
          BEGIN
          If CodeMontantSISCO='F' Then CodeMontantSISCO:='E' Else CodeMontantSISCO:='F' ;
          END ;
        If IdentPiece.CodeMontant='' Then
          BEGIN
          If (CodeMontantSISCO<>'F') And (CodeMontantSISCO<>'E') Then
            BEGIN
            If Infos.EnEuro Then IdentPiece.CodeMontant:='E--' Else IdentPiece.CodeMontant:='F--' ;
            END Else
            BEGIN
            OldCodeMontantSISCO:=IdentPiece.CodeMontant ;
            IdentPiece.CodeMontant:=CodeMontantSISCO+'--' ;
            END ;
          END ;
        If ForceTenue Then
          BEGIN
          If Infos.EnEuro Then
            BEGIN
            IdentPiece.CodeMontant:='E--' ;
            CodeMontantSISCO:='E' ;
            END Else
            BEGIN
            IdentPiece.CodeMontant:='F--' ;
            CodeMontantSISCO:='F' ;
            END ;
          RuptOrigineSaisie:=FALSE ;
          END ;
        TraiteMontantSISCOEcr(St,IdentPiece,Infos,CodeMontantSISCO,RuptOrigineSaisie,DecalAna) ;
        FaitStSISCO(St,StSISCO,OldStSISCO,IdentPiece,Infos,RuptOrigineSaisie,CodeMontantSISCO,InfoImp) ;
        MAJListeSISCO(StSISCO,ListePiece,IdentPiece,InfoImp) ;
        END ;
  'v','w' : BEGIN
{$IFDEF NOCONNECT}
Exit ;
{$ENDIF}
        OkOk:=TRUE ;
        If (Copy(IdentPiece.ligneEnCours.Gen,1,1)='6') And Infos.ShuntAnaCharge Then OkOk:=FALSE ;
        If (Copy(IdentPiece.ligneEnCours.Gen,1,1)='7') And Infos.ShuntAnaProduit Then OkOk:=FALSE ;
        If OkOk Then
          BEGIN
          Inc(IdentPiece.Chrono) ; CodeMontantSISCO:=Trim(Copy(St,85,1)) ;
          If (VH^.RecupPCL) And (Not OkCROS) And RuptOrigineSaisie And (IdentPiece.codeMontant<>'') Then
            BEGIN
            If CodeMontantSISCO='F' Then CodeMontantSISCO:='E' Else CodeMontantSISCO:='F' ;
            END ;
        If ForceTenue Then
          BEGIN
          If Infos.EnEuro Then
            BEGIN
            CodeMontantSISCO:='E' ;
            END Else
            BEGIN
            CodeMontantSISCO:='F' ;
            END ;
          RuptOrigineSaisie:=FALSE ;
          END ;
          TraiteMontantSISCOAna(St,IdentPiece,Infos,CodeMontantSISCO,RuptOrigineSaisie) ;
          FaitStSISCO(St,StSISCO,OldStSISCO,IdentPiece,Infos,RuptOrigineSaisie,CodeMontantSISCO) ;
{$IFDEF CEGID}
          St2:=Trim(BourreEtLEss(StSISCO.AuxSect,fbGene)) ;
          Q:=OpenSQL('SELECT * FROM CORRESP WHERE CR_TYPE="ANA" AND CR_CORRESP="'+St2+'" ',TRUE) ;
          If Not Q.Eof Then
            BEGIN
            StSect:=Q.FindField('CR_LIBELLE').AsString ;
            i:=1 ;
            While StSect<>'' Do
              BEGIN
              St1:=ReadTokenSt(StSect) ; If St1='' Then St1:='999999' ;
              StSisco.Axe:='A'+IntToStr(i) ;
              St1:=BourreEtLEss(St1,AxeTofb(StSisco.Axe)) ;
              St1:=Format_String(St1,17) ;
              StSisco.AuxSect:=St1 ;
              MAJListeSISCO(StSISCO,ListePiece,IdentPiece,InfoImp) ;
              inc(i) ; If i>3 Then Break ;
              END ;
            END Else
            BEGIN
            For i:=1 To 3 Do
              BEGIN
              St1:='999999' ;
              StSisco.Axe:='A'+IntToStr(i) ;
              St1:=BourreEtLEss(St1,AxeTofb(StSisco.Axe)) ;
              St1:=Format_String(St1,17) ;
              StSisco.AuxSect:=St1 ;
              MAJListeSISCO(StSISCO,ListePiece,IdentPiece,InfoImp) ;
              END ;
//            MAJListeSISCO(StSISCO,ListePiece,IdentPiece,InfoImp) ;
            END ;
          Ferme(Q) ;
{$ELSE}
          MAJListeSISCO(StSISCO,ListePiece,IdentPiece,InfoImp) ;
{$ENDIF}
          END ;
        END ;
  END ;
END ;

Procedure RecupCptAuxSISCO(Var St,St1 :String ; Var Infos : TInfoGeneSISCO; Ident : string = 'CAU' ; InfoImp : PtTInfoImport = Nil) ;
Var Cpt,Nat,Vide,Coll,Nat1,Lib : String ;
    CptOrigine : String ;
    OkNat1 : Boolean ;
    Cpt2   : string;
BEGIN
{$IFDEF NOCONNECT}
Exit ;
{$ENDIF}
Cpt:=AnsiUpperCase(Trim(Copy(St,2,10))) ; CptOrigine:=Trim(Cpt) ;
if VH^.RecupComSx and Infos.Troncaux then // ajout me 20-03-02
   Cpt2 := Copy (Cpt,2,Length(Cpt)-1);
If VH^.RecupSISCOPGI And (InfoImp<>NIL) And (InfoImp^.LGSISCO>0) Then CptOrigine:=BourreEtLess(CptOrigine,fbGene,InfoImp^.LGSISCO) ;
Cpt:=BourreEtLess(Cpt,fbAux) ;
Cpt:=CptRemplace(Cpt,Infos.CharRT) ;
Vide:=Format_String(' ',320) ;
Coll:=TrouveCptSISCO(Cpt,Infos.FourchetteSISCO,Nat1) ;
OkNat1:=FALSE ;
If Nat1<>'' Then
  BEGIN
  If Nat1='CLI' Then
    BEGIN
    Nat:='CLI' ; Infos.AuMoinsUnClient:=TRUE ; OkNat1:=TRUE ;
    END Else If Nat1='FOU' Then
    BEGIN
    Nat:='FOU' ; Infos.AuMoinsUnFournisseur:=TRUE ; OkNat1:=TRUE ;
    END ;
  END ;
If Not OkNat1 Then
  BEGIN
  If Cpt[1] in ['C','9'] Then
    BEGIN
    Nat:='CLI' ; Infos.AuMoinsUnClient:=TRUE ;
    END Else
    BEGIN
//    Nat:='FOU' ; Infos.AuMoinsUnFournisseur:=TRUE ;
        if VH^.RecupComSx then // ajout me 30-04-2002
        begin
           If Cpt[1] in ['F','0'] Then
           begin
                Nat:='FOU' ; Infos.AuMoinsUnFournisseur:=TRUE ;
           end
           else
           begin
                if Copy (Cpt,1,2) = '40' then
                begin
                     Nat:='FOU' ; Infos.AuMoinsUnFournisseur:=TRUE ;
                end
                else
                begin
                    if Copy (Cpt,1,2) = '41' then
                    begin
                         Nat:='CLI' ; Infos.AuMoinsUnClient:=TRUE ;
                    end
                    else Nat:='DIV' ;
                end;
           end;
        end
        else
        begin
                Nat:='FOU' ; Infos.AuMoinsUnFournisseur:=TRUE ;
        end;
    END ;
  END ;
Coll:=AnsiUpperCase(Format_String(Coll,17)) ;
Coll:=CptRemplace(Coll,Infos.CharRG) ;

if VH^.RecupComSx and Infos.Troncaux then // ajout me 20-03-02
   Cpt := BourreOuTronque(Cpt2,fbaux);

// ajout me St1:='***CAU'+Format_String(Cpt,17)+Format_String(Copy(St,12,25),35)+Nat+'X'+Coll+Vide ;
Lib:=Trim(Copy(St,12,25)) ;
If VH^.RecupSISCOPGI Then Lib:=Lib+'µ'+CptOrigine ;
St1:='***'+Ident+Format_String(Cpt,17)+Format_String(Lib,35)+Nat+'X'+Coll+Vide ;
If RecupSISCO Then
  BEGIN
  If Infos.RGT<>'' Then St1:=St1+Format_String(Infos.RGT,3) Else St1:=St1+'   ' ;
  If Infos.MRT<>'' Then St1:=St1+Format_String(Infos.MRT,3) Else St1:=St1+'   ' ;
  END ;
END ;

// Si le résultat de NatureCptImpot = '---', cette fonction
// donne la nature correpondante
function GetNatureDiv(Cpt : string; Sisco : TTabCptCollSISCO) : string ;
Var R1,R2 : String ;
BEGIN
Result:='DIV' ;
If Copy(Cpt,1,2)='40' Then
  BEGIN
  If TrouveCollSISCO(Cpt,Sisco) Then Result:='COF' ;
  END Else
If Copy(Cpt,1,2)='41' Then
  BEGIN
  If TrouveCollSISCO(Cpt,Sisco) Then Result:='COC' ;
  END Else
If (Copy(Cpt,1,3)='512') Or (Copy(Cpt,1,3)='514') Then BEGIN Result:='BQE' ; END Else
If Copy(Cpt,1,2)='53' Then Result:='CAI' Else
If Copy(Cpt,1,1)='7' Then
  BEGIN
  Result:='PRO' ;
  END Else
If Copy(Cpt,1,1)='6' Then
  BEGIN
  Result:='CHA' ;
  END ;
If VH^.RecupPCL And (Result='DIV') Then
  BEGIN
  R1:=Copy(Cpt,1,2) ; R2:=Copy(Cpt,1,3) ;
  If (R1='20') Or (R1='21') Or (R1='22') Or (R1='23') Or (R1='24') Or (R1='25') Or (R1='26') Or (R1='27') Then Result:='IMO' ;
  END ;
END ;

Procedure RecupCptGenSISCO(Var St,St1,StSup :String ; Var Infos : TInfoGeneSISCO ; InfoImp : PtTInfoImport = Nil) ;
Var Cpt,Nat,NatTiers,Vide,Point,Vent,Sens,Let,Lib,Nat1 : String ;
    OkNat,OkLet,OkTiers : Boolean ;
    Coll : String ;
    CRSISCO : Boolean ;
    AuxACreer : Boolean ;
    CptOrigine : String ;
BEGIN
{$IFDEF NOCONNECT}
Exit ;
{$ENDIF}
Sens:='M' ; OkLet:=FALSE ; Let:='-' ; CRSISCO:=FALSE ; AuxACreer:=FALSE ; StSup:='' ;
Cpt:=AnsiUpperCase(Trim(Copy(St,2,10))) ; CptOrigine:='' ;
Cpt:=BourreEtLess(Cpt,fbGene) ;
Cpt:=CptRemplace(Cpt,Infos.CharRG) ;
//If VH^.RecupPCL And (Not VH^.RecupSISCOPGI) Then
If (VH^.RecupPCL or VH^.RecupSISCOPGI) Then
  BEGIN
  If (Copy(Cpt,1,2)='42') Or (Copy(Cpt,1,2)='43') Or (Copy(Cpt,1,2)='46') Then
    BEGIN
    If TrouveCollSISCO(Cpt,Infos.FourchetteSISCO) Then
      BEGIN
      CRSISCO:=TRUE ;
      If Copy(Cpt,1,2)='42' Then Nat:='COS' Else Nat:='COD' ;
      END Else
      BEGIN
      Coll:=TrouveCptSISCO(Cpt,Infos.FourchetteSISCO,Nat1) ;
      If Coll<>'' Then AuxACreer:=TRUE ;
      END ;
    END ;
  END ;
Vide:=Format_String(' ',47) ; Point:='-' ; Vent:='-----' ;
If CRSISCO Then OkNat:=TRUE Else
  BEGIN
  Nat:=NatureCptImport(Cpt,Infos.FourchetteImport) ;
  OkNat:=Nat<>'---' ;
  If Not OkNat Then Nat:=GetNatureDiv(Cpt, Infos.FourchetteSISCO) ;
  END ;
If (Nat='CHA') And Infos.AnaSurSISCO And (Not Infos.ShuntAnaCharge) Then Vent[1]:='X' ;
If (Nat='PRO') And Infos.AnaSurSISCO And (Not Infos.ShuntAnaProduit) Then Vent[1]:='X' ;
If (Nat='BQE') Or (Nat='CAI') Then Point:='X' ;
If VH^.RecupPCL Or VH^.RecupSISCOPGI Then
BEGIN
  Let:=AnsiUpperCase(Trim(Copy(St,37,1))) ;
  // Cas où l'info de lettrage n'est pas présente , on force à 'N'
  If Let = '' then Let := 'N';
  //  If Let[1] In ['M','N','O'] Then OkLet:=TRUE ;
  If VH^.RecupSISCOPGI Then
  BEGIN
    If Let[1] In ['M','O','N'] Then OkLet:=TRUE ;
  END Else
  BEGIN
    If Let[1] In ['M','O'] Then OkLet:=TRUE ;
  END ;
  if VH^.RecupPCL then
  begin
    // Dans le cas de la récup. PCL, on ne prend en compte que le paramétrage
    // du TRFPARAM pour le lettrage - CA - 15/11/2002
    if (InfoImp<>NIL) then // ajout me pour comsx
       OkLet := EstLettrableSISCO(Cpt,InfoImp^.RacineCptGenLet);
  end
  else
  begin
  If OkLet And (InfoImp<>NIL) Then
    If Not EstLettrableSISCO(Cpt,InfoImp^.RacineCptGenLet) Then OkLet:=FALSE ;
  end;
END ;
If Nat<>'DIV' Then OkLet:=FALSE ;
OkTiers:=FALSE ;
If OkLet Then
  BEGIN
  Let:='X' ; Nat:='TID' ;
  If Copy(Cpt,1,2)='40' Then BEGIN Nat:='TIC' ; OkTiers:=TRUE ; END Else
    If Copy(Cpt,1,2)='41' Then BEGIN Nat:='TID' ; ; OkTiers:=TRUE ; END
  END Else
  BEGIN
  If InfoImp<>NIL Then
  begin  { ajout me 28-01-02 pour let}
    If EstLettrableSISCO(Cpt,InfoImp^.RacineCptGenPoint) Then Point:='X'
    else
    if VH^.RecupComSx then Let:='-';
{ ajout me 28-01-02}
  end else
    Let:='-';
  END ;
If VH^.RecupPCL Then
  BEGIN
  If (Nat='IMO') Or (Nat='COC') Or ((Nat='TID') And OkTiers) Or (Nat='CHA') Then Sens:='D' ;
  If (Nat='COF') Or ((Nat='TIC') And OkTiers) Or (Nat='PRO') Then Sens:='C' ;
  END ;

if VH^.RecupComSx and Infos.Troncaux then // ajout me 20-03-02
begin
   Cpt := BourreOuTronque(Cpt,fbGene);
end;

St1:='***CGE'+Format_String(Cpt,17)+Format_String(Copy(St,12,25),35)+Nat+Let+Point+
     Vent+Vide+Sens ;
If VH^.RecupPCL And (Trim(VH^.Cpta[fbGene].Attente)='') And (Copy(Cpt,1,2)='47') Then
  BEGIN
  Lib:=AnsiUpperCase(Trim(Copy(St,12,25))) ;
  If Pos('ATTENT',Lib)<>0 Then
    BEGIN
    SetParamSoc('SO_GENATTEND',Trim(Cpt)) ; VH^.Cpta[fbGene].Attente:=Trim(Cpt) ;
    END Else
    BEGIN
    If Copy(Cpt,1,3)='471' Then VH^.Cpta[fbGene].Attente:=Trim(Cpt) ;
    END ;
  END ;
If ((VH^.RecupPCL) or (VH^.RecupSiscoPGI)) And AuxACreer Then
  BEGIN
  Coll:=TrouveCptSISCO(Cpt,Infos.FourchetteSISCO,Nat1) ;
  Coll:=AnsiUpperCase(Format_String(Coll,17)) ;
  Coll:=CptRemplace(Coll,Infos.CharRG) ;
  NatTiers:='DIV' ; If Copy(Coll,1,2)='42' Then NatTiers:='SAL' ;
  StSup:='***CAU'+Format_String(Cpt,17)+Format_String(Copy(St,12,25),35)+NatTiers+'X'+Coll+Vide ;
  Vide:=Format_String(' ',320) ;
  If (RecupSISCO or (VH^.RecupSiscoPGI)) Then
    BEGIN
    If (Infos.RGT<>'') And (NatTiers<>'SAL') Then StSup:=StSup+Format_String(Infos.RGT,3) Else StSup:=StSup+'   ' ;
    If Infos.MRT<>'' Then StSup:=StSup+Format_String(Infos.MRT,3) Else StSup:=StSup+'   ' ;
    END ;
  // CA - 02/08/2002 - Pour ne créer que l'auxiliaire dans le cas des regroupements salariés ou divers
  if StSup <> '' then St1 := StSup;
  END ;
END ;

Procedure RecupCptSISCO(Var St,St1,StSup :String ; Var Infos : TInfoGeneSISCO ; InfoImp : PtTInfoImport = Nil; Ident : string = 'CAU') ;
Var EstAux : Boolean ;
    Cpt : String ;
    i   : integer;
BEGIN
{$IFDEF NOCONNECT}
Exit ;
{$ENDIF}
Cpt:=AnsiUpperCase(Trim(Copy(St,2,10))) ;
If VH^.RecupPCL Then EstAux:=Cpt[1] In ['0','9','a'..'z','A'..'Z']
                Else EstAux:=Cpt[1] In ['0','9','C','F'] ;
if VH^.RecupComSx and ( not EstAux) then // ajout me 12-02-02
   EstAux := OkAuxiliaire (Infos.FourchetteSISCO, Cpt);

If EstAux Then RecupCptAuxSISCO(St,St1,Infos,Ident, InfoImp) Else RecupCptGenSISCO(St,St1,StSup,Infos,InfoImp) ;
// Si ce n'est pas le premier fichier importé, on ne récupère pas les infos des comptes
// V. 42 : on désactive cet partie pour récupérer tous les comptes ( cas import domix  - FQ 11741 ).
(*if not VH^.RecupComSx then
if (VH^.RecupPCL) and (InfoImp^.NumFic > 1) then St1 := '';*)
END ;

Procedure RecupSRegroupeSISCO(Var St,St1 :String ; Var Infos : TInfoGeneSISCO ; InfoImp : PtTInfoImport = Nil) ;
Var C : Char ;
    NumTL : Integer ;
    Code,Lib : String ;
BEGIN
St1:='' ;
If Not VH^.RecupSISCOPGI Then Exit ;
C:=St[3] ;
NumTL:=0 ;
If EstSerie(S3) Then Case Ord(C) Of 1,2,3 : NumTL:=Ord(C) ; END
                Else Case Ord(C) Of 1,2,3,4,5 : NumTL:=Ord(C) ; END ;
If NumTL<=0 Then Exit ;
Code:=Format_String(AnsiUpperCase(Trim(Copy(St,4,4))),17) ;
Lib:=Format_String(AnsiUpperCase(Trim(Copy(St,8,30))),35) ;
St1:='***TL'+IntToStr(NumTL-1)+Code+Lib+'SEC' ;
END ;

Procedure RecupSectionSISCO(Var St,St1 : String ; Var Infos : TInfoGeneSISCO) ;
Var Cpt,Vide,Lib,StSup : String ;
    TL : Array[1..5] Of String ;
    i,j : Integer ;
BEGIN
St1:='' ;
{$IFDEF NOCONNECT}
Exit ;
{$ENDIF}
If Infos.ShuntAnaCharge And Infos.ShuntAnaProduit Then Exit ;
Cpt:=AnsiUpperCase(Trim(Copy(St,3,15))) ;
Cpt:=BourreEtLess(Cpt,fbAxe1) ;
Cpt:=CptRemplace(Cpt,Infos.CharRS) ;
Vide:=Format_String(' ',30) ;
St1:='***SAT'+Format_String(Cpt,17)+Format_String(Copy(St,18,30),35)+'A1' ;
If VH^.RecupSISCOPGI Then
  BEGIN
  Fillchar(TL,SizeOf(TL),#0) ;
  For i:=1 To 5 Do
    BEGIN
    VH^.LgTableLibre[3,i]:=6 ;
    TL[i]:=AnsiUpperCase(Trim(Copy(St,95+(4*(i-1)),4))) ;
    TL[i]:=BourreLaDoncSurLaTable('S0'+IntToStr(i-1),TL[i]) ;
    TL[i]:=Format_String(TL[i],17) ;
    END ;
  If EstSerie(S3) Then j:=3 Else j:=5 ;
  StSup:='' ;
  For i:=1 To j Do StSup:=StSup+TL[i] ;
  St1:=St1+StSup ;
  END ;
If VH^.RecupPCL And (Trim(VH^.Cpta[fbaxe1].Attente)='') And (Copy(Cpt,1,2)='99') Then
  BEGIN
  Lib:=AnsiUpperCase(Trim(Copy(St,18,30))) ;
  If Pos('ATTENT',Lib)<>0 Then
    BEGIN
    VH^.Cpta[fbAxe1].Attente:=Trim(Cpt) ;
    END Else
    BEGIN
//    If Copy(Cpt,1,3)='471' Then VH^.Cpta[fbGene].Attente:=Trim(Cpt) ;
    END ;
  END ;
END ;

Function DupliqueJal(Var St : String) : Boolean ;
Var Lib,Lg : String ;
BEGIN
Result:=FALSE ;
If St='' Then Exit ;
If Length(St)<9 Then Exit ;
If Copy(St,45,3)='ANO' Then Exit ;
St[9]:=CharJal[1] ;
Lib:=Trim(Copy(St,10,35)) ; Lib:=Lib+'(Récup SISCO)' ;Lib:=Format_String(Lib,35) ;
St:=GoInsere(St,Lib,10,35) ;
St:=GoInsere(St,'LIB',74,3) ;
Result:=TRUE ;
END ;

Function RecupJalSISCO(Var St,St1 : String ; Var Infos : TInfoGeneSISCO ; InfoImp : PtTInfoImport = Nil) : Boolean ;
Var Jal,Lib,Nat,Gen,MS,Typ,S1,CA,SoucheSim : String ;
BEGIN
{$IFDEF NOCONNECT}
Exit ;
{$ENDIF}
Result:=TRUE ;
Jal:=AnsiUpperCase(Trim(Copy(St,3,2))) ;
If VH^.RecupCEGID And (InfoImp<>NIL) And (InfoImp.BaseOrigine='CLI') Then Jal:='Z'+Jal ;
If InfoImp<>NIL Then If Not EstJalARecuperer(Jal,InfoImp^.JalFaux) Then BEGIN Result:=FALSE ; Exit ; END ;
Jal:=CptRemplace(Jal,Infos.CharRS) ;
Lib:=AnsiUpperCase(Trim(Copy(St,5,20))) ; Gen:=AnsiUpperCase(Trim(Copy(St,25,10))) ;
Typ:=AnsiUpperCase(Trim(Copy(St,36,1))) ;
Nat:='OD' ;
If (InfoImp<>NIL) And (InfoImp^.RacineJalAN<>'') And EstJalAN(Jal,InfoImp^.RacineJalAN) Then
  BEGIN
  Nat:='ANO' ; SetParamSoc('SO_JALOUVRE',Jal) ;
  END Else
If Jal='AA' Then BEGIN Nat:='ANO' ; SetParamSoc('SO_JALOUVRE',Jal) ; END Else
 If (Jal='70') Or (Jal='VT') Or (Jal='VE') Then Nat:='VTE' Else
  If  (Jal='60') Or (Jal='AC') Then Nat:='ACH' Else
   If Jal='EC' Then Nat:='REG' Else
    If (Jal='RB') Or (Jal='OD') Or (Jal='RE') Then Nat:='OD' Else
      BEGIN
      //SG6 01.02.05
      If (Typ='T') Or (Typ='B') or (Typ='P') Then
        BEGIN
        Gen:=AnsiUpperCase(Trim(Copy(St,25,10))) ; S1:=Copy(Gen,1,3) ;
        If (S1='512') Or (S1='514') Then Nat:='BQE' Else If Copy(Gen,1,2)='53' Then Nat:='CAI' Else Gen:='' ;
        //SG6 01.02.05
        if (Infos.AnaSurSISCO) and ((Typ = 'B') or (Typ = 'P')) then Nat:='ODA'; //SG6 19/11/04 FQ 14979
        END ;
      END ;
If VH^.RecupSISCOPGI Then MS:='-' Else
  BEGIN
  MS:='LIB' ; If (Typ='A') Or (Typ='V') Then MS:='BOR' ;
  If Nat='ANO' Then MS:='-' ;
  END ;
if (Nat='ANO') and (VH^.RecupPCL) then SoucheSim:=''
else SoucheSIM:='SIM';  //If Typ='X' Then SoucheSim:='SIS' ;
If Typ='X' Then
  BEGIN
  SoucheSim:='SIS' ;
  If VH^.RecupSISCOPGI And (InfoImp<>NIL) Then
    BEGIN
    InfoImp^.JalFaux:=InfoImp^.JalFaux+Jal+';' ;
    InfoImp^.JalFaux:=InfoImp^.JalFaux+Jal+Charjal+';' ;
    Result:=FALSE ; Exit ;
    END ;
  END ;
If (Nat<>'CAI') And (Nat<>'BQE') Then
  BEGIN
  Gen:='' ;
  If Trim(Copy(St,25,10))<>'' Then CA:=BourreEtLess(AnsiUpperCase(Trim(Copy(St,25,10))),fbGene)+';'  Else CA:='' ;
  END Else CA:='' ;
If (InfoImp<>NIL) And VH^.RecupSISCOPGI Then
  BEGIN
  If (InfoImp<>NIL) And (InfoImp^.RacineJalAch<>'') And EstJalAN(Jal,InfoImp^.RacineJalAch) Then Nat:='ACH' ;
  If (InfoImp<>NIL) And (InfoImp^.RacineJalVen<>'') And EstJalAN(Jal,InfoImp^.RacineJalVen) Then Nat:='VTE' ;
  END ;
If VH^.RecupSISCOPGI Then If Typ='B' Then BEGIN CA:='***' ; InfoImp^.AuMoinsUneODA:=TRUE ; END ;
St1:='***JAL'+Format_String(Jal,3)+Format_String(Lib,35)+Format_String(Nat,3)+Format_String(Jal,3)+Format_String(SoucheSIM,3)+
     Format_String(Gen,17)+'   '+Format_String(MS,3)+Format_String(CA,200)+Format_String(' ',200) ;
END ;


Function RuptureJour(Var St : String ; OldStSISCO : TStSISCO) : Boolean ;
Var J1,J2 : Integer ;
BEGIN
Result:=FALSE ;
If VH^.RecupPCL Then Exit ;
If (Trim(St)='') Or  (Trim(OldStSISCO.Date)='') Then Exit ;
J1:=StrToInt(Copy(St,2,2)) ; J2:=StrToInt(Copy(OldStSISCO.Date,1,2)) ;
Result:=J1<>J2 ;
END ;

Function PieceEquilibree(Var IdentPiece : TIdentPiece) : Boolean ;
Var D,C : Double ;
BEGIN
Result:=FALSE ;
D:=IdentPiece.TotDP ; C:=IdentPiece.TotCP ;
If Arrondi(D-C,DecimSISCO(IdentPiece,FALSE))=0 Then If Arrondi(D,DecimSISCO(IdentPiece,FALSE))<>0 Then Result:=TRUE ;
END ;

Function RuptureOrigineSaisie(Var St : String ; Var IdentPiece : TIdentPiece ; Var ChronoEclateIncremente : Boolean) : Boolean ;
Var CodeMontantSISCO,StCodeSISCO : String ;
BEGIN
Result:=FALSE ; //CodeMontantSISCO:=Trim(Copy(St,139,1)) ;
If St[1]='G' Then CodeMontantSISCO:=Trim(Copy(St,99,1))
             Else CodeMontantSISCO:=Trim(Copy(St,139,1)) ;
If ForceTenue Then Exit ;
If IdentPiece.CodeMontant='' Then Exit ; If CodeMontantSISCO='' Then Exit ;
If CodeMontantSISCO<>Copy(IdentPiece.CodeMontant,1,1) Then
  BEGIN
  If Not ChronoEclateIncremente Then
    BEGIN
    ChronoEclateIncremente:=TRUE ;
    Inc(IdentPiece.NumChronoEclate) ;
    END ;
  Result:=TRUE ;
  END ;
END ;


Function ConvertDate(St : String) : TDateTime ;
Var D,M,Y : Word ;
BEGIN
Result:=iDate1900 ;
if Trim(St)='' then Exit ;
Y:=StrToInt(Copy(St,5,2)) ; M:=StrToInt(Copy(St,3,2)) ; D:=StrToInt(Copy(St,1,2)) ;
If Y<90 Then Y:=2000+Y Else Y:=1900+y ;
Result:=EncodeDate(Y,M,D) ;
END ;


Procedure MajExo(Var Info : TInfoGeneSISCO ; Clo : Boolean) ;
Var Q : TQuery ;
    St : String ;
    EMax : Integer ;
    YY1,YY2,MM,DD : Word ;
    stEtat : string;
BEGIN
EMax:=0 ;
Q:=OpenSQL('SELECT MAX(EX_EXERCICE) FROM EXERCICE',TRUE) ;
If Not Q.Eof Then
  BEGIN
  St:=Q.Fields[0].AsString ; If St<>'' Then EMax:=StrToInt(Q.Fields[0].AsString) ;
  END ;
Ferme(Q) ;
Inc(EMax) ;
Q:=OpenSQL('SELECT * FROM EXERCICE WHERE EX_DATEDEBUT="'+USDATETIME(Info.ExoSISCO.Deb)+'" AND EX_DATEFIN="'+USDATETIME(Info.ExoSISCO.Fin)+'" ',FALSE) ;
If Q.Eof Then
  BEGIN
  Q.Insert ; InitNew(Q) ;
  Q.FindField('EX_EXERCICE').AsString:=FormatFloat('000',EMAx) ;
  DecodeDate(Info.ExoSISCO.Deb,YY1,MM,DD) ; DecodeDate(Info.ExoSISCO.Fin,YY2,MM,DD) ;
  If YY1=YY2 Then Q.FindField('EX_LIBELLE').AsString:=IntToStr(YY1) Else Q.FindField('EX_LIBELLE').AsString:=IntToStr(YY1)+'/'+IntToStr(YY2) ;
  Q.FindField('EX_ABREGE').AsString:=Q.FindField('EX_LIBELLE').AsString ;
  Q.FindField('EX_DATEDEBUT').AsDateTime:=Info.ExoSISCO.Deb ;
  Q.FindField('EX_DATEFIN').AsDateTime:=Info.ExoSISCO.Fin ;
  Q.FindField('EX_NATEXO').AsString:='' ;
  // Dans le cas de la récup PCL, on distingue ainsi les exercices créés par le pgm de récup.
  if VH^.RecupPCL then Q.FindField('EX_ETATCPTA').AsString:='ZZZ'
  else
  begin
    If Clo Then Q.FindField('EX_ETATCPTA').AsString:='CDE' Else Q.FindField('EX_ETATCPTA').AsString:='OUV' ;
  end;
  Q.FindField('EX_ETATBUDGET').AsString:='NON' ;
  Q.FindField('EX_ETATADV').AsString:='NON' ;
  Q.FindField('EX_ETATAPPRO').AsString:='NON' ;
  Q.FindField('EX_ETATPROD').AsString:='NON' ;
  Q.Post ;
  END Else
  begin
    if VH^.RecupPCL then
    begin
      if Q.FindField('EX_ETATCPTA').AsString='ZZZ' then
      begin
        Q.Edit ;
        If Clo Then Q.FindField('EX_ETATCPTA').AsString:='CDE' Else Q.FindField('EX_ETATCPTA').AsString:='OUV' ;
        Q.Post ;
      end else
      begin
        If Clo Then stEtat:='CDE' Else stEtat:='OUV' ;
        if stEtat <> Q.FindField('EX_ETATCPTA').AsString then Info.OkEnr[4]:=False;
      end;
    end else
    begin
      If Not Clo Then
      BEGIN
        Q.Edit ;
        Q.FindField('EX_ETATCPTA').AsString:='OUV' ;
        Q.Post ;
      END ;
    end;
  end;
Ferme(Q) ;
END ;

(*
procedure MiseANiveau(InfoImp : PtTInfoImport = Nil) ;
Var St,St1,St2 : String ;
BEGIN
If InfoImp=Nil Then Exit ;
If Not VH^.RecupSISCOPGI Then Exit ;
St:=InfoImp^.RacineCptGenLet ; St2:='' ;
While St<>'' Do BEGIN St1:=ReadTokenSt(St) ; St1:=BourreEtLess(St1,fbGene) ; St2:=St2+St1+';' ; END ;
InfoImp^.RacineCptGenLet:=St2 ;
St:=InfoImp^.RacineCptGenPoint ; St2:='' ;
While St<>'' Do BEGIN St1:=ReadTokenSt(St) ; St1:=BourreEtLess(St1,fbGene) ; St2:=St2+St1+';' ; END ;
InfoImp^.RacineCptGenPoint:=St2 ;
END ;
*)
Procedure RecupSociete(i : Integer ; St : String ; Var Info : TInfoGeneSISCO ; InfoImp : PtTInfoImport = Nil) ;
Var ll1,ll2 : Integer ;
    SS,Gen,Aux1,Aux2 : String ;
    Q : TQuery ;
    _LgGen,_LgAux,_LgAna,_BourreGen,_BourreAux,_BourreAna : String ;
    TS : TOB;
BEGIN
{$IFDEF RECUPPCL}
{$IFDEF CEGID}
If (i=8) Or (i=0)Then Else Exit ;
If VH^.RecupCEGID And (InfoImp<>NIL) And (InfoImp.BaseOrigine='CLI') And (i=8) Then Exit ;
{$ENDIF}
{$ENDIF}
Case i Of
  0 : BEGIN
      Info.ExoSISCO.Deb:=ConvertDate(Copy(St,8,6)) ;
      Info.ExoSISCO.Fin:=ConvertDate(Copy(St,14,6)) ;
      Info.ExoSISCO.NombrePeriode:=0 ;
      Info.EnEuro:=Trim(Copy(St,36,3))='EUR' ;
      // ajout me
      Info.lgcpt := Trim(Copy(St,21,2));
{$IFDEF NOCONNECT}
      VH^.EnCours.Code:='001' ;
      VH^.EnCours.Deb:=Info.ExoSISCO.Deb ;
      VH^.EnCours.Fin:=Info.ExoSISCO.Fin ;
{$ENDIF}
      END ;
  1 : BEGIN
      Info.Nom:=Trim(Copy(St,3,30)) ;
      END ;
  2 : BEGIN
      Info.Adr1:=Trim(Copy(St,3,32)) ;
      Info.Adr2:=Trim(Copy(St,35,32)) ;
      Info.Adr3:=Trim(Copy(St,67,32)) ;
      END ;
  3 : BEGIN
      Info.Siret:=Trim(Copy(St,3,14)) ;
      Info.APE:=Trim(Copy(St,17,4)) ;
      Info.Tel:=Trim(Copy(St,21,20)) ;
      Info.AnaSurSISCO:=Copy(ST,54,1)='O' ;
      // ajout me
      Info.Numplan := Trim(Copy(St,41,2));
      END ;
  END ;
If Info.InitPCL
{$IFDEF RECUPPCL}
{$IFDEF CEGID}
or ((i=5) Or (i=8))
{$ENDIF}
{$ENDIF}
Then
  BEGIN
  if i<=10 Then Info.OkEnr[i]:=TRUE ;
  Case i Of
    0 : BEGIN
        if VH^.RecupPCL then
        begin
          Info.OkEnr[0] := VerifCoherenceExo ( Info.ExoSISCO.Deb, Info.ExoSISCO.Fin );
          if Info.OkEnr[0] then MajExo (Info, TRUE);
        end else MajExo(Info,TRUE) ;
        _BourreGen:='0' ;
        _BourreAux:='0' ;
        If VH^.RecupSISCOPGI Then
          BEGIN
          If LitTRFParam('000','CP_LGOK')='X' Then
            BEGIN
            _LgGen:=LitTRFParam('000','CP_LGGEN') ;
            _LgAux:=LitTRFParam('000','CP_LGAUX') ;
            If (_LgGen<>'') And (StrToInt(_LgGen)>4) Then
              BEGIN
              ll1:=StrToInt(_LgGen) ;
              _BourreGen:=LitTRFParam('000','CP_BOURREGEN') ;
              If _BourreGen='' Then _BourreGen:='0' ;
              END Else ll1:=StrToInt(Copy(St,21,2)) ;
            If (_LgAux<>'') And (StrToInt(_LgAux)>4) Then
              BEGIN
              ll2:=StrToInt(_LgAux) ;
              _BourreAux:=LitTRFParam('000','CP_BOURREAUX') ;
              If _BourreAux='' Then _BourreAux:='0' ;
              END Else ll2:=StrToInt(Copy(St,21,2)) ;
            END Else
            BEGIN
            ll1:=StrToInt(Copy(St,21,2)) ; ll2:=StrToInt(Copy(St,21,2)) ;
            END ;
          END Else
          BEGIN
          ll1:=StrToInt(Copy(St,21,2)) ; ll2:=StrToInt(Copy(St,21,2)) ;
          END ;
        If InfoImp<>NIL Then InfoImp^.LGSISCO:=StrToInt(Copy(St,21,2)) ;
        if (VH^.RecupPCL) then
        begin  // si la longueur des comptes est déjà définie dans la base, on garde cette longueur
          if (not (Valeur(GetParamSoc('SO_LGCPTEGEN')) > 0)) then 
          begin
            SetParamSoc('SO_LGCPTEGEN',ll1) ;
            SetParamSoc('SO_BOURREGEN',_BourreGen) ;
            VH^.Cpta[fbGene].Lg:=ll1 ; VH^.Cpta[fbGene].Cb:=_BourreGen[1] ;
          end;
          if (not (Valeur(GetParamSoc('SO_LGCPTEAUX')) > 0)) then
          begin
            SetParamSoc('SO_LGCPTEAUX',ll2) ;
            SetParamSoc('SO_BOURREAUX',_BourreAux) ;
            VH^.Cpta[fbAux].Lg:=ll2 ; VH^.Cpta[fbAux].Cb:=_BourreAux[1] ;
          end;
        end
        else
        begin
          SetParamSoc('SO_LGCPTEGEN',ll1) ; SetParamSoc('SO_LGCPTEAUX',ll2) ;
          SetParamSoc('SO_BOURREGEN',_BourreGen) ; SetParamSoc('SO_BOURREAUX',_BourreAux) ;
          VH^.Cpta[fbGene].Lg:=ll1 ; VH^.Cpta[fbGene].Cb:=_BourreGen[1] ;
          VH^.Cpta[fbAux].Lg:=ll2 ; VH^.Cpta[fbAux].Cb:=_BourreAux[1] ;
        end;
//        MiseANiveau(InfoImp) ;
        END ;
    1 : BEGIN
          if not VH^.RecupPCL Then  // Mis à jour à la création du dossier en PCL par DP
          begin
            SS:=Trim(Copy(St,3,30)) ; SetParamSoc('SO_LIBELLE',SS) ;
            SS:=Trim(Copy(St,33,32)) ; SetParamSoc('SO_TXTJURIDIQUE',SS) ;
          end;
        END ;
    2 : BEGIN
          if not VH^.RecupPCL Then  // Mis à jour à la création du dossier en PCL par DP
          begin
            SS:=Trim(Copy(St,3,32))  ; SetParamSoc('SO_ADRESSE1',SS) ;
            SS:=Trim(Copy(St,35,32)) ; SetParamSoc('SO_ADRESSE2',SS) ;
            SS:=Trim(Copy(St,67,32)) ; SetParamSoc('SO_ADRESSE3',SS) ;
          end;
        END ;
    3 : BEGIN
          if not VH^.RecupPCL Then // Mis à jour à la création du dossier en PCL par DP
          begin
            SS:=Trim(Copy(St,3,14))  ; SetParamSoc('SO_SIRET',SS) ;
            SS:=Trim(Copy(St,17,4))  ; SetParamSoc('SO_APE',SS) ;
            SS:=Trim(Copy(St,21,12)) ; SetParamSoc('SO_TELEPHONE',SS) ;
          end;
          If VH^.RecupSISCOPGI Then
          BEGIN
            SS:=Trim(Copy(St,56,1)) ; UpdateTRFParam(SS,'000','CP_CHARFOU') ;
            SS:=Trim(Copy(St,57,1)) ; UpdateTRFParam(SS,'000','CP_CHARCLI') ;
          END ;
//        If Info.AnaSurSISCO Then
        END ;
    4 : BEGIN
        ll1:=StrToInt(Copy(St,42,1)) ;
        if VH^.RecupPCL then
        begin
          MajExo(Info,not (ll1=0));
        end else If ll1=0 Then MajExo(Info,FALSE) ;
        END ;
    8 : BEGIN
        Gen:=BourreEtLess(Trim(Copy(St,3,10)),fbGene) ;
        If (Gen<>'') And (Gen[1] in ['6','7']) Then Exit ;
        Aux1:=BourreEtLess(Trim(Copy(St,13,10)),fbAux) ;
        Aux2:=BourreEtLess(Trim(Copy(St,23,10)),fbAux) ;
        Q:=OpenSQL('SELECT * FROM CORRESP WHERE CR_TYPE="SIS" AND CR_CORRESP="'+GEN+'" And CR_LIBELLE="'+Aux1+'" AND CR_ABREGE="'+Aux2+'" ',FALSE) ;
        If Q.Eof Then
          BEGIN
          Q.Insert ; InitNew(Q) ;
          Q.FindField('CR_TYPE').AsString:='SIS' ;
          Q.FindField('CR_CORRESP').AsString:=Gen ;
          Q.FindField('CR_LIBELLE').AsString:=Aux1 ;
          Q.FindField('CR_ABREGE').AsString:=Aux2 ;
          Q.Post ;
          END ;
        (*
        // AJOUT ME 14/01/2000 les fourchettes n'étaient pas chargées, donc pb: les comptes auxiliaires dans g_general
        If Q.Eof Then ChargeCptSISCO(Info.FourchetteSISCO,'') ;
        *)
        Ferme(Q) ;
        END ;
    9  : BEGIN
            TS := TOB.Create ('Stat',InfoImp.TOBStat,-1);
            TS.AddChampSupValeur('CODE',Trim(Copy(St,3,4)),False);
            TS.AddChampSupValeur('LIBELLE',Trim(Copy(St,7,25)),False);
         END;
    50 : BEGIN
         _BourreAna:='0' ;
         If VH^.RecupSISCOPGI Then
           BEGIN
           If LitTRFParam('000','CP_LGOK')='X' Then
             BEGIN
             _LgAna:=LitTRFParam('000','CP_LGANA') ;
             If (_LgAna<>'') And (StrToInt(_LgAna)>4) Then
               BEGIN
               ll1:=StrToInt(_LgAna) ;
               _BourreAna:=LitTRFParam('000','CP_BOURREANA') ;
               If _BourreAna='' Then _BourreAna:='0' ;
               END Else ll1:=StrToInt(Copy(St,44,2)) ;
             END Else
             BEGIN
             ll1:=StrToInt(Copy(St,44,2)) ;
             END ;
           END Else
           BEGIN
           ll1:=StrToInt(Copy(St,44,2)) ;
           END ;
        Q:=OpenSQL('SELECT * FROM AXE WHERE X_AXE="A1"',FALSE) ;
        If Not Q.Eof Then
          BEGIN
          Q.Edit ;
          Q.FindField('X_LONGSECTION').AsInteger:=ll1 ;
          Q.FindField('X_BOURREANA').AsString:=_BourreAna ;
          VH^.Cpta[fbAxe1].Lg:=ll1 ; VH^.Cpta[fbAxe1].Cb:=_BourreAna[1] ;
          Q.FindField('X_SECTIONATTENTE').AsString:=BourreEtLess(Q.FindField('X_SECTIONATTENTE').AsString,fbAxe1) ;
          Q.Post ;
          END ;
        Ferme(Q) ;
        END ;
    END ;
  END ;
END ;

Function AlimExo(Var Info : TInfoGeneSISCO ; Exo : TExoDate) : Integer ;
Var i,j,k : Integer ;
BEGIN
{$IFDEF NOCONNECT}
Exit ;
{$ENDIF}
j:=0 ; k:=0 ; Result:=0 ;
If Exo.Deb<>Info.ExoSISCO.Deb Then Result:=2 ;
If Exo.Fin<>Info.ExoSISCO.Fin Then Result:=2 ;
If Result=0 Then
  BEGIN
  Info.ExoSISCO:=Exo ;
  If Exo.Code=VH^.EnCours.Code Then Info.StPlusExo:='' Else
    If Exo.Code=VH^.Suivant.Code Then Info.StPlusExo:='+' Else
      If Exo.Code=VH^.Precedent.Code Then Info.StPlusExo:='-' Else
        For i:=1 To 5 Do
          BEGIN
          If VH^.ExoClo[i].Code=Exo.Code Then j:=i ;
          If VH^.ExoClo[i].Code<>'' Then Inc(k) ;
          END ;
  If k<>0 Then For i:=1 To k-j+1 Do Info.StPlusExo:=Info.StPlusExo+'-' ;
  END ;
END ;

Function TraiteEnr(Var NewFichier : TextFile ; ListePiece : TStringList ;
                   Var St,St1 : String ; Var IdentPiece : TIdentPiece ;
                   Var StSISCO : TStSISCO ; Var Info : TInfoGeneSISCO ;
                   Var CloseOpenAFaire : Boolean ; InfoImp : PtTInfoImport = Nil ;
                   ForceDecoupe : Boolean = FALSE) : Integer ;
Var MonoExo : Boolean ;
    Exo1,Exo2 : TExoDate ;
    StSup : String ;
    OkE : Boolean ;
    OkPiece : Integer ;
    YYSoldeAuJour : Integer ;
    PieceAvecMvt,RuptOrigineSaisie,ChronoEclateIncremente : Boolean ;
BEGIN
Result:=0 ; StSup:='' ; CloseOpenAFaire:=FALSE ;
If (InfoImp<>Nil) And (Not InfoImp^.RecupODA) And ((St[1] in ['G','w'])=TRUE) Then Exit ;
If Copy(St,1,2)='00' Then
  BEGIN
  RecupSociete(0,St,Info,InfoImp)  ;
{$IFNDEF NOCONNECT}
  If Not Info.InitPCL Then
    BEGIN
    If QuelExoDate2(Info.ExoSISCO.Deb,Info.ExoSISCO.Fin,MonoExo,Exo1,Exo2) Then
      BEGIN
      If Not MonoExo Then Result:=1 Else
        If RecupSISCO Then
          BEGIN
          END Else
          BEGIN
          If (Exo1.Code=VH^.EnCours.Code) Or ((VH^.Suivant.Code<>'') And (Exo1.Code=VH^.Suivant.Code)) Then Else Result:=2 ;
          END ;
      END Else Result:=1 ;
    If Result=0 Then BEGIN Result:=AlimExo(Info,Exo1) ; END ;
    END ;
{$ENDIF}
  END Else
{$IFNDEF NOCONNECT}
If Copy(St,1,2)='01' Then RecupSociete(1,St,Info,InfoImp) Else
If Copy(St,1,2)='02' Then RecupSociete(2,St,Info,InfoImp) Else
If Copy(St,1,2)='03' Then RecupSociete(3,St,Info,InfoImp) Else
If Copy(St,1,2)='04' Then RecupSociete(4,St,Info,InfoImp) Else
If Copy(St,1,2)='08' Then RecupSociete(8,St,Info,InfoImp) Else
If Copy(St,1,2)='09' Then RecupSociete(9,St,Info,InfoImp) Else
If Copy(St,1,2)='50' Then RecupSociete(50,St,Info,InfoImp) Else
{$ENDIF}
If Not Info.InitPCL Then
  If Copy(St,1,2)='05' Then
    BEGIN
    If RecupJalSISCO(St,St1,Info,InfoImp) Then
      BEGIN
      WriteLn(NewFichier,St1) ;
      If VH^.RecupSISCOPGI Then
         BEGIN
         If DupliqueJal(St1) Then If St1<>'' Then WriteLn(NewFichier,St1) ;
         END ;
      END ;
    END Else If Not Info.InitPCL Then
    BEGIN
    Case St[1] Of
  {$IFNDEF NOCONNECT}
      'C' : BEGIN RecupCptSISCO(St,St1,StSup,Info,InfoImp) ; If St1<>'' then WriteLn(NewFichier,St1) ; If VH^.RecupPCL And (StSup<>'') Then WriteLn(NewFichier,StSup) ; END ;    // Généraux & Auxiliaires
      'R' : BEGIN RecupSRegroupeSISCO(St,St1,Info,InfoImp) ; If St1<>'' Then WriteLn(NewFichier,St1) ; END ;
      'S' : BEGIN RecupSectionSISCO(St,St1,Info) ; If St1<>'' Then WriteLn(NewFichier,St1) ; END ; // Section

  {$ENDIF}
      'M' : BEGIN
            YYSoldeAuJour:=EcritPiece(NewFichier,ListePiece,IdentPiece,Info,PieceAvecMvt,InfoImp) ;
{$IFDEF RECUPPCL}
{$IFDEF CEGID}
            If PieceAvecMvt Then CLoseOpenAFaire:=TRUE ;
{$ENDIF}
{$ENDIF}
            If (VH^.RecupLTL Or VH^.RecupSISCOPGI ) And PieceAvecMvt And ForceDecoupe Then CLoseOpenAFaire:=TRUE ;
            OkPiece:=FaitPiece(St,ListePiece,IdentPiece,StSISCO,Info,InfoImp) ;
            If VH^.RecupPCL And (OkPiece<>0) Then BEGIN Info.PbEnr[0]:=TRUE ; END ;
            IdentPiece.NumChronoEclate:=0 ; ChronoEclateIncremente:=FALSE ;
            END ; // Mois
      'J' : BEGIN
            YYSoldeAuJour:=EcritPiece(NewFichier,ListePiece,IdentPiece,Info,PieceAvecMvt,InfoImp) ; FaitPiece(St,ListePiece,IdentPiece,StSISCO,Info,InfoImp) ;
            IdentPiece.NumChronoEclate:=0 ; ChronoEclateIncremente:=FALSE ;
            END ; // Journal
      'F' : BEGIN
            ChronoEclateIncremente:=FALSE ;
            If VH^.RecupPCL Then
              BEGIN
              OkE:=PieceEquilibree(IdentPiece) ;
              IdentPiece.FolioSISCO:='Folio '+Trim(Copy(St,2,3)) ;
              If VH^.RecupCegid Or (Arrondi(IdentPiece.TotDP-IdentPiece.TotCP,DecimSISCO(IdentPiece,FALSE))=0) Then
                IdentPiece.NoFolioSISCO:=StrToInt(Trim(Copy(St,2,3))) ;
              END Else
              BEGIN
              IdentPiece.FolioSISCO:='Folio '+Trim(Copy(St,2,3)) ; IdentPiece.NoFolioSISCO:=StrToInt(Trim(Copy(St,2,3))) ;
              END ;
            END ;// Rien à faire
      'E','G' : BEGIN
            RuptOrigineSaisie:=RuptureOrigineSaisie(St,IdentPiece,ChronoEclateIncremente) ;
            If Not RuptOrigineSaisie Then ChronoEclateIncremente:=FALSE ;
            If VH^.RecupPCL Then
              BEGIN
              OkE:=PieceEquilibree(IdentPiece) ;
//              If RuptureJour(St,StSisco) (*Or PieceEquilibree(IdentPiece)*) Or RuptureOrigineSaisie(St,IdentPiece) Then YYSoldeAuJour:=EcritPiece(NewFichier,ListePiece,IdentPiece,Info,PieceAvecMvt,InfoImp) ;
              END Else
              BEGIN
              If RuptureJour(St,StSisco) Or PieceEquilibree(IdentPiece) Or RuptureOrigineSaisie(St,IdentPiece,ChronoEclateIncremente) Then YYSoldeAuJour:=EcritPiece(NewFichier,ListePiece,IdentPiece,Info,PieceAvecMvt,InfoImp) ;
              END ;
            FaitPiece(St,ListePiece,IdentPiece,StSISCO,Info,InfoImp,RuptOrigineSaisie) ;
            END ; // Ecriture
      'p' : ; // Non traité pour l'instant
  {$IFNDEF NOCONNECT}
      'v','w' : BEGIN FaitPiece(St,ListePiece,IdentPiece,StSISCO,Info,InfoImp) ; END ; // Ecriture analytique
  {$ENDIF}
      'A' : ; // Non traité pour l'instant
      'B' : ; // Non traité pour l'instant
      END ;
    END ;
END ;

Procedure DechiffreFT(i : Integer ; St : String ; Var CptSISCO : TTabCptCollSISCO ; What : Integer) ;
Var St1,St2,St3 : String ;
    k1,k2 : Integer ;
BEGIN
K1:=Pos(',',St) ; k2:=Pos(':',St) ;
If (K1>0) And (k2>0) And (K2>K1) Then
  BEGIN
  St1:=Copy(St,1,K1-1) ; St2:=Copy(St,K1+1,K2-K1-1) ; St3:=Copy(St,K2+1,Length(St)-K2) ;
  CptSISCO[i].Cpt:=AnsiUpperCase(Trim(St1)) ;
  CptSISCO[i].Aux1:=AnsiUpperCase(Trim(St2)) ;
  CptSISCO[i].Aux2:=AnsiUpperCase(Trim(St3)) ;
  If What=0 Then CptSISCO[i].Nat:='CLI' Else CptSISCO[i].Nat:='FOU'
  END ;
END ;

Procedure chargeCptSISCO(Var CptSISCO : TTabCptCollSISCO ; ProfilPCL : String = '') ;
Var Q : TQuery ;
    i,What : Integer ;
    FourchetteVide : Boolean ;
    St,St1 : String ;
BEGIN
Fillchar(CptSISCO,SizeOf(CptSISCO),#0) ;
{$IFDEF NOCONNECT}
Exit ;
{$ENDIF}
Q:=OpenSQL('Select CR_CORRESP,CR_LIBELLE,CR_ABREGE FROM CORRESP WHERE CR_TYPE="SIS"',TRUE) ;
i:=0 ; FourchetteVide:=TRUE ;
While Not Q.Eof Do
  BEGIN
  FourchetteVide:=FALSE ;
  CptSISCO[i].Cpt:=AnsiUpperCase(Trim(Q.Fields[0].AsString)) ;
  CptSISCO[i].Aux1:=AnsiUpperCase(Trim(Q.Fields[1].AsString)) ;
  CptSISCO[i].Aux2:=AnsiUpperCase(Trim(Q.Fields[2].AsString)) ;
  Inc(i) ;
  Q.Next ;
  END ;
Ferme(Q) ;
If FourchetteVide And VH^.RecupPCL And (Not VH^.RecupSISCOPGI) Then
  BEGIN
  Q:=OpenSQL('SELECT * FROM TRFPARAM WHERE TRP_CODE="'+ProfilPCL+'" AND TRP_NOM LIKE "CP_FOURCHETTES%"',TRUE) ;
  i:=0 ;
  while not Q.EOF do
    BEGIN
    St:=Q.FindField('TRP_DATA').AsString ; What:=-1 ;
    If Q.FindField('TRP_NOM').AsString='CP_FOURCHETTESC' Then What:=0 Else
     If Q.FindField('TRP_NOM').AsString='CP_FOURCHETTESF' Then What:=1 ;
    if What<>-1 Then
      While St<>'' Do
        BEGIN
        St1:=ReadTokenSt(St) ; DechiffreFT(i,St1,CptSISCO,What) ; Inc(i) ;
        END ;
    Q.Next ;
    END ;
  Ferme(Q) ;
  END ;
END ;

Procedure CompleteInfo(Var InfoGeneSISCO : TInfoGeneSISCO) ;
BEGIN
If Not VH^.RecupSISCOPGI Then Exit ;
If LitTRFParam('000','CP_RECUPANACHA')='X' Then InfoGeneSISCO.ShuntAnaCharge:=FALSE Else InfoGeneSISCO.ShuntAnaCharge:=TRUE ;
If LitTRFParam('000','CP_RECUPANAPRO')='X' Then InfoGeneSISCO.ShuntAnaProduit:=FALSE Else InfoGeneSISCO.ShuntAnaProduit:=TRUE ;
END ;

Function TransfertSISCO(StFichier : String ; RemplaceLeFichier : Boolean ; Var NbFic : Integer ; NewRep : String = '' ;
                        InfoImp : PtTInfoImport = Nil ; ForceDecoupe : Boolean = FALSE) : Integer ;
Var Fichier,NewFichier : TextFile ;
    St,St1,StNewFichier,StNewFichier2 : String ;
    ListePiece : TStringList ;
    IdentPiece : TIdentPiece ;
    StSISCO : TStSISCO ;
    InfoGeneSISCO : TInfoGeneSISCO ;
    Pb,FichierOk : Boolean ;
    YYSoldeAuJour : Integer ;
    ProfilPCL : String ;
    OkOk : Boolean ;
    CompteurFichier : Integer ;
    PieceAvecMvt : Boolean ;
    CloseOpenAFaire : Boolean ;
BEGIN
Result:=0 ; Pb:=FALSE ; NbFic:=0 ;
If VH^.RecupSISCOPGI Then If InfoImp<>NIL Then InfoImp^.RecupODA:=TRUE ;
InitMove(1000,TraduireMemoire('Chargement du fichier en cours...')) ;
AssignFile(Fichier,StFichier) ; Reset(Fichier) ;
StNewFichier:=FileTemp('.PNM') ;
AssignFile(NewFichier,StNewFichier) ; Rewrite(NewFichier) ;
ReadLn(Fichier,St) ; // Début du fichier SISCO
ListePiece:=TStringList.Create ;
ListePiece.Sorted:=FALSE ; //ListePiece.Duplicates:=dupAccept ;
Fillchar(IdentPiece,SizeOf(IdentPiece),#0) ; IdentPiece.NumP:=1 ;
Writeln(NewFichier,StFichier) ;
Fillchar(StSISCO,SizeOf(StSISCO),#0) ;
Fillchar(InfoGeneSISCO,SizeOf(InfoGeneSISCO),#0) ;
ProfilPCL:='' ; If InfoImp<>NIL Then If InfoIMP^.ProfilPCL<>'' Then ProfilPCL:=InfoIMP^.ProfilPCL ;
ChargeCptSISCO(InfoGeneSISCO.FourchetteSISCO,ProfilPCL) ;
ChargeFourchetteCompte('SIS',InfoGeneSISCO.FourchetteImport) ;
ChargeCharRemplace(InfoGeneSISCO) ;
CompleteInfo(InfoGeneSISCO) ;
InfoGeneSISCO.LJalLu:=TStringList.Create ; CompteurFichier:=-1 ;
{$IFNDEF NOCONNECT}
InitRequete(InfoGeneSISCO.QFiche[3],3) ;
{$ENDIF}
{$IFDEF RECUPPCL}
{$IFDEF CEGID}
If InfoImp<>NIL Then If Not InfoImp^.OnFicheBase Then CompteurFichier:=0 ;
{$ENDIF}
{$ENDIF}
If (VH^.RecupLTL Or VH^.RecupSISCOPGI) And ForceDecoupe Then CompteurFichier:=0 ;
While (Not EOF(Fichier)) And (Not Pb)  do
  BEGIN
  MoveCur(FALSE) ;
  ReadLn(Fichier,St) ; If Trim(St)='' Then Continue ;
  OkOk:=TRUE ;
{$IFDEF RECUPPCL}
{$IFDEF CEGID}
  If InfoImp<>NIL Then If InfoImp^.OnFicheBase Then OkOk:=(Copy(St,1,2)='05') Or (Copy(St,1,2)='08') ;
{$ENDIF}
{$ENDIF}
  If OkOk Then Result:=TraiteEnr(NewFichier,ListePiece,St,St1,IdentPiece,StSISCO,InfoGeneSISCO,CloseOpenAFaire,InfoImp) ;
  If InfoGeneSISCO.PbEnr[0] Then Break ;
{$IFDEF RECUPPCL}
{$IFDEF CEGID}
  If CloseOpenAFaire Then CloseOpenNewFile(NewFichier,StNewFichier,RemplaceLeFichier,StFichier,NewRep,CompteurFichier)  ;
{$ENDIF}
{$ENDIF}
  If (VH^.RecupLTL Or VH^.RecupSISCOPGI) And CloseOpenAFaire And ForceDecoupe Then CloseOpenNewFile(NewFichier,StNewFichier,RemplaceLeFichier,StFichier,NewRep,CompteurFichier)  ;
  If Result<>0 Then Pb:=TRUE ;
  If HalteLa Then Break ;
  END ;
YYSoldeAuJour:=EcritPiece(NewFichier,ListePiece,IdentPiece,InfoGeneSISCO,PieceAvecMvt,InfoImp) ;
FiniMove ; Flush(NewFichier) ;
CloseFile(Fichier) ; CloseFile(NewFichier) ;
{$IFNDEF NOCONNECT}
Ferme(InfoGeneSISCO.QFiche[3]) ;
{$ENDIF}
VideStringList(ListePiece) ; ListePiece.Free ;
VideStringList(InfoGeneSISCO.LJalLu) ; InfoGeneSISCO.LJalLu.Free ;
If Result=0 Then CloseEtRenameNewFile(RemplaceLeFichier,StFichier,StNewFichier,NewRep,CompteurFichier) ;
If ForceDecoupe And VH^.RecupSISCOPGI And (CompteurFichier>=0)  Then NbFic:=CompteurFichier+1 ;
(*
AssignFile(Fichier,StFichier) ; Erase(Fichier) ;
*)
(*
If Result=0 Then
  BEGIN
  If RemplaceLeFichier Then
    BEGIN
    StNewFichier2:=NewNomFic(StFichier,'OLD') ;
    {$i-}
    AssignFile(NewFichier,StNewFichier2) ; Erase(NewFichier) ;
    {$i+}
    Hrenamefile(StFichier,StNewFichier2) ;
    Hrenamefile(StNewFichier,StFichier) ;
    END Else
    BEGIN
    If NewRep='' Then
      BEGIN
      If RecupSISCO Then StNewFichier2:=NewNomFic(StFichier,'CGE')
                    Else StNewFichier2:=NewNomFic(StFichier,'CGN') ;
      END Else
      BEGIN
      If RecupSISCO Then StNewFichier2:=NewNomFicEtDir(StFichier,'CGE',NewRep)
                    Else StNewFichier2:=NewNomFicEtDir(StFichier,'CGN',NewRep) ;
      END ;
    {$i-}
    AssignFile(NewFichier,StNewFichier2) ; Erase(NewFichier) ;
    {$i+}
    Hrenamefile(StNewFichier,StNewFichier2) ;
    END ;
  END ;
*)

InitMove(500,'') ;
Repeat
  MoveCur(FALSE) ;
  AssignFile(Fichier,StFichier) ;
  {$I-} Reset (Fichier) ; {$I+}
  FichierOk:=ioresult=0 ;
  If FichierOk Then Close(Fichier) ;
Until FichierOk ;
FiniMove ;
If VH^.RecupPCL And (InfoImp<>Nil) Then
  BEGIN
  InfoImp^.AuMoinsUnClient:=InfoGeneSISCO.AuMoinsUnClient ;
  InfoImp^.AuMoinsUnFournisseur:=InfoGeneSISCO.AuMoinsUnFournisseur ;
  InfoImp^.OkFouCli:=TRUE ;
  InfoImp^.OkFouFou:=TRUE ;
  If (InfoImp^.AuMoinsUnFournisseur) And (Not FourchetteSISCOExiste(0,InfoGeneSISCO.FourchetteSISCO)) Then InfoImp^.OkFouFou:=FALSE ;
  If (InfoImp^.AuMoinsUnClient) And (Not FourchetteSISCOExiste(1,InfoGeneSISCO.FourchetteSISCO)) Then InfoImp^.OkFouCli:=FALSE ;
  InfoImp^.PbEnr:=InfoGeneSISCO.PbEnr ;
  END ;
END ;

Function GetNomDossier(StFichier : String) : String ;
Var Fichier : TextFile ;
    St,St1 : String ;
    OkOk : Boolean ;
    i : Integer ;
    S1,S2 : String ;
BEGIN
Result:='' ; OkOk:=FALSE ; i:=0 ;
AssignFile(Fichier,StFichier) ; Reset(Fichier) ;
ReadLn(Fichier,St) ; // Début du fichier SISCO
While Not EOF(Fichier)   do
  BEGIN
  Inc(i) ;
  ReadLn(Fichier,St) ; St:=Trim(St) ;
  If St<>'' Then St1:=Copy(St,1,2) ;
  If St1='01' Then
    BEGIN
    S1:=Trim(Copy(St,3,30)) ; S2:=Trim(Copy(St,33,32)) ; Result:=S1+'/'+S2 ; Break ;
    END ;
  If i>10 Then Break ;
  END ;
CloseFile(Fichier) ;
END ;


Function InitRecupDossier(StFichier : String ; Var InfoImp : TInfoImport) : Integer ;
Var Fichier,NewFichier : TextFile ;
    St,St1,StNewFichier,StNewFichier2 : String ;
    ListePiece : TStringList ;
    IdentPiece : TIdentPiece ;
    StSISCO : TStSISCO ;
    InfoGeneSISCO : TInfoGeneSISCO ;
    Pb,FichierOk : Boolean ;
    i : Integer ;
    CloseOpenAFaire : Boolean ;
BEGIN
Result:=0 ; Pb:=FALSE ;
InitMove(1000,TraduireMemoire('Chargement du fichier en cours...')) ;
Fillchar(InfoGeneSISCO,SizeOf(InfoGeneSISCO),#0) ; InfoGeneSISCO.InitPCL:=TRUE ;
CompleteInfo(InfoGeneSISCO) ;
AssignFile(Fichier,StFichier) ; Reset(Fichier) ;
ReadLn(Fichier,St) ; // Début du fichier SISCO
While (Not EOF(Fichier)) And (Not Pb)  do
  BEGIN
  MoveCur(FALSE) ;
  ReadLn(Fichier,St) ;
  Result:=TraiteEnr(NewFichier,ListePiece,St,St1,IdentPiece,StSISCO,InfoGeneSISCO,CloseOpenAFaire,@InfoImp) ;
  If Result<>0 Then Pb:=TRUE ;
  If HalteLa Then Break ;
  END ;
FiniMove ;
CloseFile(Fichier) ;
(*
AssignFile(Fichier,StFichier) ; Erase(Fichier) ;
*)
InitMove(500,'') ;
Repeat
  MoveCur(FALSE) ;
  AssignFile(Fichier,StFichier) ;
  {$I-} Reset (Fichier) ; {$I+}
  FichierOk:=ioresult=0 ;
  If FichierOk Then Close(Fichier) ;
Until FichierOk ;
FiniMove ;
For i:=0 To 10 Do InfoImp.OkEnr[i]:=InfoGeneSISCO.OkEnr[i] ;
END ;

Procedure AlimCptResultat(Var CritEdt : TCritEdt) ;
Var Q : TQuery ;
    St : String ;
begin
{$IFDEF SPEC302}
St:='Select SO_FERMEBIL, SO_OUVREBIL, SO_RESULTAT, SO_FERMEPERTE, SO_OUVREPERTE, ' ;
St:=St+'SO_FERMEBEN, SO_OUVREBEN, SO_JALFERME, SO_JALOUVRE FROM SOCIETE WHERE SO_SOCIETE="'+V_PGI.CodeSociete+'"' ;
Q:=OpenSQL(St,TRUE) ;
If Not Q.Eof Then
   BEGIN
   CritEdt.CptLibre1:=Q.FindField('SO_FERMEPERTE').AsString ;
   CritEdt.CptLibre2:=Q.FindField('SO_FERMEBEN').AsString ;
   CritEdt.CptLibre3:=Q.FindField('SO_OUVREPERTE').AsString ;
   CritEdt.CptLibre4:=Q.FindField('SO_OUVREBEN').AsString ;
   END ;
Ferme(Q) ;
{$ELSE}
CritEdt.CptLibre1:=GetParamSoc('SO_FERMEPERTE') ;
CritEdt.CptLibre2:=GetParamSoc('SO_FERMEBEN') ;
CritEdt.CptLibre3:=GetParamSoc('SO_OUVREPERTE') ;
CritEdt.CptLibre4:=GetParamSoc('SO_OUVREBEN') ;
{$ENDIF}
end ;

Function WhereSISCO(Var CritEdt : TCritEdt ; Resultat : Boolean ; Valeur : String ; EtatRevision : TCheckBoxState) : String ;
Var StWhere,St1 : String ;
    StSaufJal,St2 : String ;
    SetType : SetttTypePiece ;
BEGIN
StWhere:=' WHERE E_DATECOMPTABLE>="'+USDateTime(CritEdt.DateDeb)+'" AND E_DATECOMPTABLE<="'+USDateTime(CritEdt.DateFin) +'"' ;
if CritEdt.Exo.Code<>'' then StWhere:=StWhere+' AND E_EXERCICE="'+CritEdt.Exo.Code+'"' ;
if CritEdt.Etab<>'' then StWhere:=StWhere+' AND E_ETABLISSEMENT="'+CritEdt.Etab+'"' ;
StWhere:=StWhere+' AND E_NUMEROPIECE>='+IntToStr(CritEdt.GL.NumPiece1) ;
StWhere:=StWhere+' AND E_NUMEROPIECE<='+IntToStr(CritEdt.GL.NumPiece2) ;
If CritEdt.Cpt1<>'' Then StWhere:=StWhere+' AND E_JOURNAL>="'+CritEdt.Cpt1+'"' ;
If CritEdt.Cpt2<>'' Then StWhere:=StWhere+' AND E_JOURNAL<="'+CritEdt.Cpt2+'"' ;
if CritEdt.GL.Sauf<>'' then
  BEGIN
  StSaufJal:=CritEdt.GL.Sauf ;
  while StSaufJal<>'' do StWhere:=StWhere+' AND E_JOURNAL<>"'+ReadTokenSt(StSaufJal)+'"' ;
  END ;
SetType:=WhatTypeEcr(Valeur,V_PGI.Controleur,EtatRevision);
St2:=WhereSupp('E_',SetType) ;
If St2<>'' Then StWhere:=StWhere+St2 ;
//StWhere:=StWhere+' AND E_QUALIFPIECE="N"' ;
If Not Resultat Then
  BEGIN
  If CritEdt.CptLibre1<>'' Then StWhere:=StWhere+' AND E_GENERAL<>"'+CritEdt.CptLibre1+'" ' ;
  If CritEdt.CptLibre2<>'' Then StWhere:=StWhere+' AND E_GENERAL<>"'+CritEdt.CptLibre2+'" ' ;
  If CritEdt.CptLibre3<>'' Then StWhere:=StWhere+' AND E_GENERAL<>"'+CritEdt.CptLibre3+'" ' ;
  If CritEdt.CptLibre4<>'' Then StWhere:=StWhere+' AND E_GENERAL<>"'+CritEdt.CptLibre4+'" ' ;
  END Else
  BEGIN
  St1:='' ;
  If CritEdt.CptLibre1<>'' Then St1:=St1+' E_GENERAL="'+CritEdt.CptLibre1+'" OR ' ;
  If CritEdt.CptLibre2<>'' Then St1:=St1+' E_GENERAL="'+CritEdt.CptLibre2+'" OR ' ;
  If CritEdt.CptLibre3<>'' Then St1:=St1+' E_GENERAL="'+CritEdt.CptLibre3+'" OR ' ;
  If CritEdt.CptLibre4<>'' Then St1:=St1+' E_GENERAL="'+CritEdt.CptLibre4+'" OR ' ;
  If St1<>'' Then
    BEGIN
    Delete(St1,Length(St1)-3,3) ; StWhere:=StWhere+' AND ('+St1+') ' ;
    END Else StWhere:='' ;
  END ;
Result:=StWhere ;
END ;

Function SelectBalSISCO(Var CritEdt : TCritEdt ; Resultat : Boolean ; Valeur : String ; EtatRevision : TCheckBoxState) : String ;
Var St,StWhere : String ;
BEGIN
Result:='' ;
St:='SELECT E_GENERAL, SUM(E_DEBIT) AS DP, SUM(E_CREDIT) AS CP FROM ECRITURE' ;
StWhere:=WhereSISCO(CritEdt,Resultat,Valeur,EtatRevision) ;
If StWhere='' Then Exit ;
St:=St+StWhere ;
St:=St+' GROUP BY E_GENERAL ' ;
Result:=St ;
END ;

// ajout me : fonctions déplacées dans utiltrans.pas
{FUNCTION AGauche (st : string; l : Integer ; C : Char) : string ;
var St1 : String ;
BEGIN
st1:=Trim(st) ;
if ((l>0) and (l<Length(St1))) then St1:=Copy(St1,1,l) ;
While Length(St1)<l Do St1:=St1+C ;
Result:=st1 ;
END ;

FUNCTION ADroite (st : string; l : Integer ; C : Char) : string ;
var St1 : String ;
BEGIN
st1:=Trim(st) ;
if ((l>0) and (l<Length(St1))) then St1:=Copy(St1,1,l) ;
While Length(St1)<l Do St1:=C+St1 ;
Result:=st1 ;
END ;}

{============================================================================}
FUNCTION STRFMONTANTSISCO ( Tot : Extended ; Long,Dec : Integer ; symbole : string3 ; Separateur : Boolean) : string ;
Var Mask : string ;
    fact : Double ;
BEGIN
If Dec>0 Then BEGIN Tot:=Arrondi(Tot,Dec) ; Fact:=IntPower(10,Dec) ; Tot:=Tot*Fact ; Dec:=0 ; END ;
Mask:=StrFMask(Dec,Symbole,Separateur) ;
Result:=FormatFloat(Mask,Tot) ;
END ;

Procedure FaitSISCODebut(Var FF : TextFile) ;
Var St,StM,StD,StA : String ;
    i : Integer ;
BEGIN
St:=FormatDateTime('dd-mm-yyyy',Date) ;
StD:=Copy(St,1,2) ; StA:=Copy(St,7,4) ;
i:=StrToInt(Copy(St,4,2)) ;
Case i Of 1 : StM:='JAN' ; 2 : StM:='FEB' ; 3 : StM:='MAR' ; 4 : StM:='APR' ;
          5 : StM:='MAY' ; 6 : StM:='JUN' ; 7 : StM:='JUL' ; 8 : StM:='AUG' ;
          9 : StM:='SEP' ; 10 : StM:='OCT' ; 11 : StM:='NOV' ; 12 : StM:='DEC' ;
          End ;
St:='***DEBUT***'+StD+'-'+StM+'-'+StA+'-000000' ;
WriteLn(FF,St) ;
END ;

Procedure FaitSISCOFin(Var FF : TextFile) ;
Var St : String ;
BEGIN
St:='***FIN***' ;
WriteLn(FF,St) ;
END ;

Procedure FaitSISCO00(Var FF : TextFile ; Var CritEdt : TCritEdt; Typarchive : string='B') ;
Var St : String ;
    Lg,Dec : Integer ;
    Mon : String ;
BEGIN
lg:=VH^.Cpta[fbGene].Lg ; If Lg>10 Then Lg:=10 ;
Dec:=V_PGI.OkDecV ; Mon:='FRF' ;
If CritEdt.Monnaie=2 Then BEGIN Dec:=V_PGI.OkDecE ; Mon:='EUR' ; END ;
St:='00'+ADroite(CritEdt.RefInterne,5,'0')+FormatDateTime('ddmmyy',CritEdt.Exo.Deb)+FormatDateTime('ddmmyy',CritEdt.Exo.Fin)+
// ajout me à la place de 'B'
//    Typarchive +FormatFloat('00',Lg)+FormatFloat('0',Dec)+'0810      GM'+Mon ; 0810 --> 1010 pour gestion euro
    Typarchive +FormatFloat('00',Lg)+FormatFloat('0',Dec)+'1010      GM'+Mon ;
WriteLn(FF,St) ;
END ;

Procedure FaitSISCOCptGen(Var FF : TextFile ; Var CritEdt : TCritEdt; bExit : boolean = True) ;
Var Q : TQuery ;
    St : String ;
    Naturecpte : string;
    Listecpte  : TStringList;
    i          : integer;
BEGIN
if bExit then Exit ;
NatureCpte := 'G';
Listecpte := TStringList.Create;

Q := OpenSql ('SELECT * from CORRESP Where CR_TYPE="SIS"', TRUE);
while not Q.EOF do
begin
     Listecpte.add (Q.FindField ('CR_CORRESP').asstring);
     Q.next;
end;
ferme (Q);
                                                              // ajout me
Q:=OpenSQL('SELECT G_GENERAL,G_LIBELLE, G_TOTDEBP, G_TOTCREP, G_NATUREGENE FROM GENERAUX',TRUE) ;
InitMove(RecordsCount(Q)-1,'') ;
While Not Q.Eof Do
  BEGIN
  MoveCur(FALSE) ;
  NatureCpte := 'G';
  For i:=0 To Listecpte.Count-1 Do
  begin
  (*
       if Q.FindField('G_GENERAL').AsString = ListeCpte[i] then
       begin
           if Q.FindField('G_NATUREGENE').AsString = 'COF' then NatureCpte := 'F'
           else
           if Q.FindField('G_NATUREGENE').AsString = 'COC' then NatureCpte := 'C';
       end;
  *)
       if BourreOuTronque(Q.FindField('G_GENERAL').AsString, fbgene) = BourreOuTronque(ListeCpte[i],fbgene) then
       begin
           if Q.FindField('G_NATUREGENE').AsString = 'COF' then NatureCpte := 'F'
           else
           if Q.FindField('G_NATUREGENE').AsString = 'COC' then NatureCpte := 'C'
           else
           if (Q.FindField('G_NATUREGENE').AsString = 'COS') or
           (Q.FindField('G_NATUREGENE').AsString = 'COD')
           then NatureCpte := 'G'; // ajout me 20-03-02
       end;
  end;
  St:='C'+AGauche(Q.FindField('G_GENERAL').AsString,10,'0')+
//  ajout me        AGauche(Q.FindField('G_LIBELLE').AsString,25,' ')+'NNG'+
          AGauche(Q.FindField('G_LIBELLE').AsString,25,' ')+'NN'+ NatureCpte +
          ADroite(StrfMontantSISCO(Q.FindField('G_TOTDEBP').AsFloat,14,CritEdt.Decimale,'',False),14,'0')+
          ADroite(StrfMontantSISCO(Q.FindField('G_TOTCREP').AsFloat,14,CritEdt.Decimale,'',False),14,'0')+
          '      ' ;
  WriteLn(FF,St) ;
  Q.Next ;
  END ;
Ferme(Q) ;
FiniMove ;
VideStringList(ListeCpte) ;
Listecpte.free;
END ;

Procedure FaitSISCOPer(Var FF : TextFile ; DD : TDateTime ; Exo : TExoDate) ;
Var St : String ;
    Y,M,D,Y1,M1,D1 : Word ;
    i : Integer ;
BEGIN
Decodedate(Exo.Deb,Y,M,D) ; Decodedate(DD,Y1,M1,D1) ; i:=12*(Y1-Y)+M1-M+1 ;
St:='M'+FormatFloat('00',i) ;
WriteLn(FF,St) ;
END ;

Procedure FaitSISCOFolio(Var FF : TextFile ; NF : Integer) ;
Var St : String ;
BEGIN
St:='F'+FormatFloat('000',NF) ;
WriteLn(FF,St) ;
END ;

Procedure FaitSISCOBalEcr(Var FF : TextFile ; Q : TQuery ; Var CritEdt : TCritEdt ; Resultat : Boolean ; DP,CP : Double ; Var NbLignes : Integer) ;
Var St : String ;
    Y,M,D : Word ;
    Lib : String ;
    MD,MC : Double ;
    Mon : String ;
    Q1 : TQuery ;
    Cpt : String ;
BEGIN
//Decodedate(Q.FindField('E_DATECOMPTABLE').AsDateTime,Y,M,D) ;
Mon:='F' ;
If CritEdt.Monnaie=2 Then BEGIN Mon:='E' ; END ;
DecodeDate(CritEdt.DateFin,Y,M,D) ; Lib:='' ;
If Resultat Then Cpt:=CritEdt.CptLibre11 Else Cpt:=Q.FindField('E_GENERAL').AsString ;
Q1:=OpenSQL('SELECT G_LIBELLE FROM GENERAUX WHERE G_GENERAL="'+Cpt+'" ',TRUE) ;
If Not Q1.Eof Then Lib:=Q1.FindField('G_LIBELLE').AsString ;
Ferme(Q1) ;
If Resultat Then BEGIN MD:=DP ; MC:=CP ; END Else
  BEGIN
  MD:=Arrondi(Q.FindField('DP').AsFloat,CritEdt.Decimale) ;
  MC:=Arrondi(Q.FindField('CP').AsFloat,CritEdt.Decimale) ;
  END ;
If MD<>0 Then
  BEGIN
  St:='E'+FormatFloat('00',D)+AGauche(Cpt,10,'0')+
          AGauche(Lib,20,' ')+'0000000001'+' '+' '+ADroite(StrfMontantSISCO(MD,13,CritEdt.Decimale,'',False),13,'0')+
//                             Stat    Eche    MP    QTe     Date Creat T    N° chèque
          ADroite('0',13,'0')+'    '+'      '+'  '+'00000000'+'      '+' '+'          '+
//        Date Rel  N   Dev    Mnt dev        Date val L
          '      '+' '+'   '+'0000000000000'+'      '+' '+Mon+ADroite('0',13,'0')+'00000000' ;
  WriteLn(FF,St) ;
  Inc(NbLignes)
  END ;
If MC<>0 Then
  BEGIN
  St:='E'+FormatFloat('00',D)+AGauche(Cpt,10,'0')+
          AGauche(Lib,20,' ')+'0000000001'+' '+' '+ADroite('0',13,'0')+ADroite(StrfMontantSISCO(MC,13,CritEdt.Decimale,'',False),13,'0')+
//         Stat    Eche    MP    QTe     Date Creat T    N° chèque
          '    '+'      '+'  '+'00000000'+'      '+' '+'          '+
//        Date Rel  N   Dev    Mnt dev        Date val L
          '      '+' '+'   '+'0000000000000'+'      '+' '+Mon+ADroite('0',13,'0')+'00000000' ;
  WriteLn(FF,St) ;
  Inc(NbLignes)
  END ;
END ;

Procedure ExportBALSISCO(StFichier : String ; Var CritEdt : TCritEdt ; Valeur : String ; EtatRevision : TCheckBoxState) ;
Var Q : TQuery ;
    St : String ;                
    Fichier : TextFile ;
    NumFolio,NbLigne : Integer ;
    D,C : Double ;
BEGIN
Assign(Fichier,StFichier) ; Rewrite(Fichier) ; D:=0 ; C:=0 ;
AlimCptResultat(CritEdt) ;
FaitSISCODebut(Fichier) ;
FaitSISCO00(Fichier,CritEdt) ;
FaitSISCOCptGen(Fichier,CritEdt) ;
FaitSISCOPER(Fichier,CritEdt.DateFin,CritEdt.Exo) ;
St:='Jco BALANCE A '+Format_String(FormatDateTime('MMM YYYY',CritEdt.DateFin),10) ; WriteLn(Fichier,St) ;
NbLigne:=0 ; NumFolio:=1 ; FaitSISCOFolio(Fichier,NumFolio) ;
St:=SelectBalSISCO(CritEdt,TRUE,Valeur,EtatRevision) ;
If St<>'' Then
  BEGIN
  Q:=OpenSQL(St,TRUE) ;
  While Not Q.Eof Do
    BEGIN
    D:=Arrondi(D+Q.FindField('DP').AsFloat,CritEdt.Decimale) ; C:=Arrondi(C+Q.FindField('CP').AsFloat,CritEdt.Decimale) ;
    Q.Next ;
    END ;
  FaitSISCOBalEcr(Fichier,Q,CritEdt,TRUE,D,C,NbLigne) ;
  Ferme(Q) ;
  END ;
D:=0 ; C:=0 ;
St:=SelectBalSISCO(CritEdt,FALSE,Valeur,EtatRevision) ;
Q:=OpenSQL(St,TRUE) ;
InitMove(RecordsCount(Q)-1,'') ;
While Not Q.Eof Do
  BEGIN
  MoveCur(FALSE) ;
//  Inc(NbLigne) ;
  If NbLigne>=90 Then BEGIN NbLigne:=1 ; Inc(NumFolio) ; FaitSISCOFolio(Fichier,NumFolio) ; END ;
  FaitSISCOBalEcr(Fichier,Q,CritEdt,FALSE,D,C,NbLigne) ;
  Q.Next ;
  END ;
FaitSISCOFin(Fichier) ;
Ferme(Q) ;
FiniMove ;
CloseFile(Fichier) ;
END ;


end.

