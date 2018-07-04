unit TOFPurge;

interface

uses Classes, StdCtrls,
{$IFDEF EAGLCLIENT}
{$ELSE}
     dbtables,FE_Main,QRS1,
{$ENDIF}
     comctrls,UTof, HCtrls, UTob,
     Ent1, Graphics, HTB97, spin,
     SysUtils, TofMeth, Forms, HPanel, Controls,
     UtilPGI, SaisUtil, ExpMvt, ImpFicU, HStatus, Vierge ;

Const Titre='Purge des écritures.' ;

type
    TOF_PURGE = class(TOF_Meth)
  private
    ExoAPurger                  : THValComboBox ;
    PeriodeDate1,PeriodeDate2   : THEdit ;
    ExoOuvert                   : TExoDate ;

    function RechercheFirstDate : TDateTime ;
    function DonnerDateFinExo (Exercice : string) : TDateTime ;
    procedure CherchePremierExoOuvert ( var Exo : TExoDate ) ;
    procedure ExoApurgerOnChange (Sender : TObject) ;
    procedure BLancerOnClick (Sender : TObject) ;
    procedure CreerNouvelleEcriture(TRC, TRD : TOB; EcrCumul : boolean; NP : Integer; Var NL : Integer) ;
    procedure TraitementPurge (SupAna : boolean ) ;
    procedure ExportEcriture(TobExo : Tob ; CptExo : integer; SupAna : boolean) ;
    procedure RecupEcritures (Exo : String; var TobEcr, TobCum : Tob ) ;
    procedure FenetreInactive ;
    Function ExoOK : Boolean ;
    procedure ChargeExercice(var TobExo: Tob);
    function ExoCorrect: boolean;
    function VerifAnalyt: boolean;

  public
    procedure OnArgument(Arguments : string); override ;
  end;

Function AucunExoCloture : boolean ;

implementation

uses HEnt1, LicUtil, HMsgBox ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 09/08/2001
Modifié le ... : 20/09/2001
Description .. : Renvoie la date de début du 1er exercice cloturé
Suite ........ : définitivement et non purgé.
Mots clefs ... : DEBUT;EXERCICE;CLOTURE
*****************************************************************}
function TOF_Purge.RechercheFirstDate : TDateTime ;
Var ReqSQL : string ;
    Q : TQuery ;
BEGIN
  ReqSQL := 'SELECT MIN(EX_DATEDEBUT) DATEMIN FROM EXERCICE WHERE EX_ETATCPTA="CDE" AND EX_NATEXO=""';
  Q := OpenSQL(ReqSQL, True) ;
  Result := Q.FindField('DATEMIN').asDateTime ;
  Ferme(Q) ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 09/08/2001
Modifié le ... :   /  /
Description .. : Renvoie les dates de début et de fin du premier exercice
Suite ........ : ouvert.
Mots clefs ... : EXERCICE;OUVERT;DATES
*****************************************************************}
procedure TOF_Purge.CherchePremierExoOuvert (Var Exo : TExoDate) ;
Var ReqSQL : string ;
    Q : TQuery ;
BEGIN
  ReqSQL := 'SELECT MIN(EX_DATEDEBUT) DATEDEB, MIN(EX_DATEFIN) DATEFIN, MIN(EX_EXERCICE) EXO FROM EXERCICE WHERE EX_ETATCPTA="OUV"' ;
  Q:=OpenSQL(ReqSQL, True) ;
  Exo.Code := Q.FindField('EXO').asString ;
  Exo.Deb := Q.FindField('DATEDEB').asDateTime ;
  Exo.Fin := Q.FindField('DATEFIN').asDateTime ;
  Ferme(Q) ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 10/08/2001
Modifié le ... :   /  /
Description .. : Renvoie la date de fin de l'exercice selectionné.
Mots clefs ... : FIN,EXERCICE
*****************************************************************}
Function TOF_Purge.DonnerDateFinExo (Exercice:string) : TDateTime ;
var Q : TQuery ;
BEGIN
  Q := OpenSQL('SELECT EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="'+Exercice+'"',true) ;
  result := Q.FindField('EX_DATEFIN').asDateTime ;
  Ferme(Q) ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 20/09/2001
Modifié le ... :   /  /
Description .. : Mise à jour de la date de fin de la période de purge en
Suite ........ : fonction de l'exercice choisi.
Mots clefs ... :
*****************************************************************}
procedure TOF_Purge.ExoApurgerOnChange (Sender : TObject) ;
BEGIN
  If ExoApurger.ItemIndex > 0 then
    BEGIN
    If PGIASK(TraduireMemoire('L''exercice choisi n''est pas le premier à purger. Tous les exercices précédents seront purgés. Confirmez-vous ?'),TraduireMemoire(Titre))=mrNo then
      BEGIN
      ExoApurger.ItemIndex := 0 ;
      Exit ;
      END ;
    END ;
  // recherche de la date de fin de l'exercice selectionné
  PeriodeDate2.Text := DateToStr(DonnerDateFinExo(ExoApurger.Value)) ;
END ;

function TOF_Purge.ExoCorrect : boolean ;
var ReqSql : string ;
Q : TQuery ;
begin
  result := false ;
  ReqSql := 'SELECT EX_EXERCICE, EX_DATEDEBUT, EX_DATEFIN, EX_NATEXO FROM EXERCICE WHERE EX_DATEDEBUT>="' + USDateTime(StrToDate(PeriodeDate1.text)) + '" ' ;
  ReqSQL := ReqSQL + 'AND EX_DATEFIN<="' + USDateTime(StrToDate(PeriodeDate2.text)) + '" AND EX_NATEXO<>"PUR"' ;
  Q := OpenSQL(ReqSQL, TRUE) ;
  result := not Q.eof ;
  Ferme(Q) ;
end ;

Function TOF_Purge.VerifAnalyt : boolean ;
var ReqSql : string ;
    Q : TQuery ;
begin
  result := false ;
  ReqSql := 'SELECT Y_GENERAL FROM ANALYTIQ WHERE Y_DATECOMPTABLE>="' + USDateTime(StrToDate(PeriodeDate1.text)) + '" ' ;
  ReqSQL := ReqSQL + 'AND Y_DATECOMPTABLE<="' + USDateTime(StrToDate(PeriodeDate2.text)) + '"' ;
  Q := OpenSql ( ReqSql, True ) ;
  If not Q.eof then
    result := PGIASK(TraduireMemoire('La période comporte de l''analytique. Voulez-vous le supprimer ?'),TraduireMemoire(Titre) )= MrYes;
  Ferme(Q) ;
end ;

procedure TOF_Purge.BLancerOnClick (Sender : TObject) ;
var SupAna : boolean ;
BEGIN
  if not ExoCorrect then
    begin
    PGIInfo(TraduireMemoire('Aucun exercice n''est à purger.'),TraduireMemoire(Titre)) ;
    Exit ;
    end ;
  if GetControlText('JOURNAL')='' then
    BEGIN
    PGIInfo(TraduireMemoire('Le journal d''à nouveau doit être renseigné.'),TraduireMemoire(Titre)) ;
    SetFocusControl('JOURNAL') ;
    Exit ;
    END ;
  if GetControlText('FICEXP')='' then
    BEGIN
    PGIInfo(TraduireMemoire('Le nom du fichier export doit être renseigné.'),TraduireMemoire(Titre)) ;
    SetFocusControl('FICEXP') ;
    Exit ;
    END ;
  SupAna := VerifAnalyt ;
  If PGIAsk(TraduireMemoire('Les écritures de la période vont être définitivement supprimées. Confirmez-vous le traitement?'),TraduireMemoire(Titre))=mrYes then
    BEGIN
    EnableControls(Ecran, False ) ;
    TraitementPurge (supAna ) ;
    EnableControls(Ecran, True ) ;
    ExoApurger.ReLoad ;
//    TFVierge(Ecran).FormShow(self) ;
    END ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payot
Créé le ...... : 13/08/2001
Modifié le ... :   /  /
Description .. : Création d'une nouvelle écriture de cumul ou mise à jour de
Suite ........ : l'écriture détail.
Mots clefs ... : ECRITURE
*****************************************************************}
procedure TOF_Purge.CreerNouvelleEcriture(TRC, TRD : TOB; EcrCumul : boolean; NP : Integer; Var NL : Integer) ;
Var TRNew : TOB ;
    DebitTr : Double ;
    DateCpt : TDateTime ;
BEGIN
  Inc(NL) ;
  TRNew := TOB.Create('ECRITURE',Nil,-1) ;
  If EcrCumul Then TRNew.Dupliquer(TRC,True,True)
              Else TRNew.Dupliquer(TRD,True,True) ;

  DateCpt := TRNew.GetValue('E_DATECOMPTABLE') ;
  TRNew.PutValue('E_EXERCICE', ExoOuvert.Code) ;
  TRNew.PutValue('E_JOURNAL', GetControlText('JOURNAL')) ;
  TRNew.PutValue('E_DATECOMPTABLE', StrToDate(GetControlText('DATEDEBUT')) ) ;
  TRNew.PutValue('E_NUMEROPIECE', NP) ;
  TRNew.PutValue('E_NUMLIGNE', NL) ;
  TRNew.PutValue('E_NATUREPIECE', 'OD') ;
  TRNew.PutValue('E_VALIDE', 'X') ;
  TRNew.PutValue('E_DATECREATION', Date) ;
  TRNew.PutValue('E_CREERPAR', 'PUR') ;
  TRNew.PutValue('E_PERIODE', GetPeriode(ExoOuvert.Deb)) ;
  TRNew.PutValue('E_SEMAINE', NumSemaine(ExoOuvert.Deb)) ;
  TRNew.PutValue('E_LIBRETEXTE8', GetControlText('REFERENCE')) ;
  TRNew.PutValue('E_LIBRETEXTE9', GetControlText('LIBELLE')) ;
  TRNew.PutValue('E_PAQUETREVISION', 1) ;
  If EcrCumul then
    BEGIN // écriture de cumul
    TRNew.PutValue('E_ECRANOUVEAU', 'OAN') ;
    TRNew.PutValue('E_QUALIFPIECE', 'N') ;
    TRNew.PutValue('E_CONFIDENTIEL', '0') ;
    TRNew.PutValue('E_ETATLETTRAGE', 'AL') ;

    // inversion des montants
    DebitTr:=TRNew.GetValue('D') ;
    TRNew.PutValue('E_DEBIT', TRNew.GetValue('C')) ;
    TRNew.PutValue('E_CREDIT', DebitTr) ;
    DebitTr:=TRNew.GetValue('DD') ;
    TRNew.PutValue('E_DEBITDEV', TRNew.GetValue('CD')) ;
    TRNew.PutValue('E_CREDITDEV', DebitTr) ;
    DebitTr:=TRNew.GetValue('DE') ;
    TRNew.PutValue('E_DEBITEURO', TRNew.GetValue('CE')) ;
    TRNew.PutValue('E_CREDITEURO', DebitTr) ;

    // zones reprise de la première écriture détail
    TRNew.PutValue('E_SOCIETE', TRD.GetValue('E_SOCIETE')) ;
    TRNew.PutValue('E_ETABLISSEMENT', TRD.GetValue('E_ETABLISSEMENT')) ;
    TRNew.PutValue('E_DEVISE', TRD.GetValue('E_DEVISE')) ;
    TRNew.PutValue('E_TAUXDEV', TRD.GetValue('E_TAUXDEV')) ;
    TRNew.PutValue('E_DATETAUXDEV', TRD.GetValue('E_DATETAUXDEV')) ;
    END
  Else
   BEGIN // écriture de détail
   if TRNew.GetValue('E_LIBREDATE') = StrToDate('01/01/1900') Then TRNew.PutValue('E_LIBREDATE', DateCpt) ;
   TRNew.PutValue('E_ECRANOUVEAU', 'H') ;
   END ;
  If EcrCumul then TRNew.InsertDB(Nil)
              else TRNew.UpdateDB ;
  TRNew.Free ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 20/09/2001
Modifié le ... :   /  /
Description .. : Export des écritures pour un exercice donné.
Mots clefs ... :
*****************************************************************}
procedure TOF_Purge.ExportEcriture (TobExo : Tob; CptExo : integer; SupAna : boolean ) ;
var CritExpMvt : TCritExpMvt ;
begin
  CritExpMvt.Date1 := StrToDate(TobExo.GetValue('EX_DATEDEBUT')) ;
  CritExpMvt.Date2 := StrToDate(TobExo.GetValue('EX_DATEFIN')) ;
  CritExpMvt.Etab := '';
  CritExpMvt.Exo := TobExo.GetValue('EX_EXERCICE');
  CritExpMvt.Format := 'CGE';
  CritExpMvt.Jal1 := '';
  CritExpMvt.Jal2 := '';
  CritExpMvt.NomFic := NewNomFic(GetControlText('FICEXP'), IntToStr(CptExo+1)) ;
  CritExpMvt.NumP1 := 0;
  CritExpMvt.NumP2 := 99999999;
  CritExpMvt.OkExport := True ;
  CritExpMvt.TypeEcr := '<<Tous>>';
  CritExpMvt.AvecODA := SupAna ;
  ExternalExportMvt('FEC', CritExpMvt) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 20/09/2001
Modifié le ... :   /  /
Description .. : Recherche des écritures pour la période donnée.
Mots clefs ... :
*****************************************************************}
procedure TOF_Purge.RecupEcritures (Exo : String; var TobEcr, TobCum : Tob ) ;
var ReqSQL, Where : string ;
    Q : TQuery ;
BEGIN
  InitMove(10000, '' ) ;
  MoveCur(False) ;
  // ecritures de détail
  ReqSQL :='SELECT E_AUXILIAIRE,E_CONFIDENTIEL,E_CREDIT,E_CREDITDEV,E_CREDITEURO,E_CREERPAR,E_DATECOMPTABLE,E_DATECREATION,' ;
  ReqSQL := ReqSQL + 'E_DATETAUXDEV,E_DEBIT,E_DEBITDEV,E_DEBITEURO,E_DEVISE,E_ECRANOUVEAU,E_ETABLISSEMENT,E_ETATLETTRAGE,E_EXERCICE,' ;
  ReqSQL := ReqSQL + 'E_GENERAL,E_JOURNAL,E_LIBREDATE,E_LIBRETEXTE8,E_LIBRETEXTE9,E_NATUREPIECE,E_NUMECHE,E_NUMEROPIECE,E_NUMLIGNE,' ;
  ReqSQL := ReqSQL + 'E_PERIODE,E_QUALIFPIECE,E_SEMAINE,E_SOCIETE,E_TAUXDEV,E_VALIDE, E_PAQUETREVISION ' ;
  ReqSQL := ReqSQL + ',G_POINTABLE FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL ' ;
  WHere := 'WHERE E_EXERCICE="'+Exo+'" AND E_QUALIFPIECE="N" AND E_ECRANOUVEAU="N" ' ;
  Where := Where + 'AND ((E_ETATLETTRAGE="AL" OR E_ETATLETTRAGE="PL") OR (E_REFPOINTAGE="" AND G_POINTABLE="X")) ' ;
  ReqSQL := ReqSQL + Where + 'ORDER BY E_GENERAL, E_AUXILIAIRE, E_DEVISE' ;
  Q := OpenSQL(ReqSQL, TRUE) ;
  Try
    TobEcr.LoadDetailDB('ECRITURE', '', '', Q, FALSE, TRUE) ;
  Finally
    Ferme(Q);
  END ;
  MoveCur(False) ;
  // ecritures de cumul
  ReqSQL := 'SELECT E_GENERAL, E_AUXILIAIRE, E_DEVISE, SUM(E_DEBIT) D, SUM(E_CREDIT) C,' ;
  ReqSQL := ReqSQL + ' SUM(E_DEBITDEV) DD, SUM(E_DEBITEURO) DE, SUM(E_CREDITDEV) CD, SUM(E_CREDITEURO) CE ' ;
  ReqSQL := ReqSQL + 'FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL ' ;
  ReqSQL := ReqSQL + Where + 'GROUP BY E_GENERAL, E_AUXILIAIRE, E_DEVISE ORDER BY E_GENERAL, E_AUXILIAIRE, E_DEVISE' ;
  Q := OpenSQL(ReqSQL, TRUE) ;
  Try
    TobCum.LoadDetailDB('ECRITURE', '', '', Q, FALSE, TRUE) ;
  Finally
    Ferme(Q);
  END ;
  FiniMove ;
END ;

procedure TOF_Purge.ChargeExercice ( var TobExo : Tob ) ;
var ReqSql : string ;
    Q : TQuery ;
begin
  ReqSql := 'SELECT EX_EXERCICE, EX_DATEDEBUT, EX_DATEFIN, EX_NATEXO FROM EXERCICE WHERE EX_DATEDEBUT>="' + USDateTime(StrToDate(PeriodeDate1.text)) + '" ' ;
  ReqSQL := ReqSQL + 'AND EX_DATEFIN<="' + USDateTime(StrToDate(PeriodeDate2.text)) + '"' ;
  Q := OpenSQL(ReqSQL, TRUE) ;
  TobExo := TOB.Create('_EXERCICE',Nil,-1) ;
  Try
    TobExo.LoadDetailDB('EXERCICE', '', '', Q, FALSE, TRUE) ;
  Finally
    Ferme(Q);
  END ;
end ;

procedure TOF_Purge.TraitementPurge (SupAna : boolean ) ;
Var ReqSQL : string ;
    TobExo, TobEcr, TobCum : TOB ;
    NumPiece, NumLigne : LongInt ;
    CptExo, CptEcr, CptCum : Integer ;
    TREcr, TRCum : TOB ;
    St : String ;
BEGIN

  if Not _BlocageMonoPoste(True) then Exit ;
  SetControlVisible('AFFTRAITEMENT',TRUE) ;
  St:=TraduireMemoire('Traitement en cours, veuillez patienter ...') ;
  SetControlCaption('AFFTRAITEMENT',St) ;

  ChargeExercice (TobExo ) ;

  NumLigne := 0 ;
  NumPiece := -1 ;

  For CptExo := 0 To TobExo.Detail.Count-1 do
    BEGIN
      St:=TraduireMemoire('Export des écritures ...') ;
      SetControlCaption('AFFTRAITEMENT',St) ;
      // création du fichier ASCII
      ExportEcriture(TobExo.Detail[CptExo],CptExo, SupAna) ;

    Try
      BeginTrans ;
      St:=TraduireMemoire('Recherche des écritures ...') ;
      SetControlCaption('AFFTRAITEMENT',St) ;
      // Récupération des écritures à purger et des ecritures de cumul
      TobEcr := TOB.Create('_ECRITURE',Nil,-1) ;
      TobCum := TOB.Create('_ECRCUMUL',Nil,-1) ;
      RecupEcritures (TobExo.Detail[CptExo].GetValue('EX_EXERCICE'), TobEcr, TobCum ) ;

      St:=TraduireMemoire('Report des écritures d''à nouveaux ...') ;
      SetControlCaption('AFFTRAITEMENT',St) ;
      InitMove(TobCum.Detail.count+TobEcr.Detail.count, '' ) ;

      // recherche du numéro de pièce en fonction du code journal s'il y a au moins une ecriture à reporter
      If (TobCum.Detail.Count > 0) And (NumPiece = -1) then NumPiece := GetNewNumJal(GetControlText('JOURNAL'),True,Date) ;

      // Report des écritures d'à nouveaux
      CptEcr := 0 ;
      For CptCum := 0 To TobCum.Detail.Count-1 do
        BEGIN
        MoveCur(False ) ;
        TRCum := TobCum.Detail[CptCum] ;
        TREcr := TobEcr.Detail[CptEcr] ;
        CreerNouvelleEcriture(TRCum, TREcr, True, NumPiece, NumLigne) ;
        While (TRCum.GetValue('E_GENERAL') = TrEcr.GetValue('E_GENERAL'))
              And (TRCum.GetValue('E_AUXILIAIRE') = TrEcr.GetValue('E_AUXILIAIRE'))
              And (TRCum.GetValue('E_DEVISE') = TrEcr.GetValue('E_DEVISE'))
              And (CptEcr <= TobEcr.Detail.count-1) Do
          BEGIN
          MoveCur(False ) ;
          CreerNouvelleEcriture(Nil, TREcr, False, NumPiece, NumLigne) ;
          Inc(CptEcr) ;
          If CptEcr <= TobEcr.Detail.count-1 Then TREcr := TobEcr.Detail[CptEcr] ;
          END ;
      END ;
      FiniMove ;
      If TobEcr <> Nil then TobEcr.Free ;
      If TobCum <> Nil then TobCum.Free ;
      CommitTrans ;
      // mise à jour de l'exercice
      TobExo.Detail[CptExo].PutValue('EX_NATEXO','PUR' ) ;
    Except
      Rollback ;
      PgiBox(TraduireMemoire('Une erreur est survenue lors du report des écritures, traitement interrompu'),TraduireMemoire(Titre)) ;
      Exit ;
    End ;
  END ;

  St:=TraduireMemoire('Mise à jour des exercices ...') ;
  SetControlCaption('AFFTRAITEMENT',St) ;
  TobExo.UpdateDB ;
  If TobExo <> Nil Then TobExo.free ;

  // suppression des écritures analytiques de la période si demandé
  if SupAna then
    begin
    St:=TraduireMemoire('Suppression des écritures analytiques...') ;
    SetControlCaption('AFFTRAITEMENT',St) ;
    InitMove(10000,'') ;
    MoveCur(False) ;
    ReqSQL := 'DELETE FROM ANALYTIQ WHERE Y_DATECOMPTABLE>="' + USDateTime(StrToDate(PeriodeDate1.text)) + '" ' ;
    ReqSQL := ReqSQL + 'AND Y_DATECOMPTABLE<="' + USDateTime(StrToDate(PeriodeDate2.text)) + '"' ;
    Try
      BeginTrans ;
      ExecuteSQL(ReqSQL) ;
      CommitTrans ;
    Except
      Rollback ;
      PgiBox(TraduireMemoire('Une erreur est survenue lors de la suppression des écritures analytiques, traitement interrompu'),TraduireMemoire(Titre)) ;
      Exit ;
    End ;
    end ;
    
  St:=TraduireMemoire('Suppression des écritures comptables ...') ;
  SetControlCaption('AFFTRAITEMENT',St) ;
  InitMove(10000,'') ;
  MoveCur(False) ;
  // suppression des écritures de la période
  ReqSQL := 'DELETE FROM ECRITURE WHERE E_DATECOMPTABLE>="' + USDateTime(StrToDate(PeriodeDate1.text)) + '" ' ;
  ReqSQL := ReqSQL + 'AND E_DATECOMPTABLE<="' + USDateTime(StrToDate(PeriodeDate2.text)) + '"' ;
  Try
    BeginTrans ;
    ExecuteSQL(ReqSQL) ;
    CommitTrans ;
  Except
    Rollback ;
    PgiBox(TraduireMemoire('Une erreur est survenue lors de la suppression des écritures comptables, traitement interrompu'),TraduireMemoire(Titre)) ;
    Exit ;
  End ;
  FiniMove ;

  SetControlVisible('AFFTRAITEMENT',FALSE) ;
  PGIInfo(TraduireMemoire('Traitement terminé'), TraduireMemoire(Titre)) ;
  _DeblocageMonoPoste(True) ;
END ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 01/08/2001
Modifié le ... : 20/09/2001
Description .. : Renvoie Vrai si aucun des exercices de la société n'est
Suite ........ : cloturé définitivement ou aucun n'est à purger.
Mots clefs ... : EXERCICE; CLOTURE
*****************************************************************}
Function AucunExoCloture : boolean ;
Var ReqSQL : String ;
BEGIN
  ReqSQL := 'SELECT EX_EXERCICE FROM EXERCICE WHERE EX_ETATCPTA="CDE" AND EX_NATEXO=""' ;
  Result := Not ExisteSQL(ReqSQL) ;
END ;

procedure TOF_Purge.FenetreInactive ;
BEGIN
//  EnableControls(Ecran,False) ;
  SetControlEnabled('BLANCER', false) ;
END ;

Function Tof_Purge.ExoOK : Boolean ;
BEGIN
  result := true ;
  If AucunExoCloture Then
    BEGIN
    PGIBox(TraduireMemoire('Aucun exercice ne peut être purgé. Le traitement est impossible.'),TraduireMemoire(Titre)) ;
    Result:= False ;
    END ;
END ;

Procedure TOF_Purge.OnArgument(Arguments : string) ;
BEGIN
  inherited ;

  If not ExoOk Then
    begin
    FenetreInactive ;
//    Exit ;
    end
  else
  If PGIAsk(TraduireMemoire('Vous devez avoir lettré et pointé tout ce qu''il est possible de faire. Voulez-vous continuer ?'),TraduireMemoire(Titre))=mrNo Then
    begin
    FenetreInactive ;
//    Exit ;
    end ;

  ExoApurger := THValComboBox(GetControl('EXOAPURGER')) ;
  PeriodeDate1 := THEdit(GetControl('PERIODE1')) ;
  PeriodeDate2 := THEdit(GetControl('PERIODE2')) ;

  CherchePremierExoOuvert (ExoOuvert) ;

  SetControlText('DATEDEBUT',DateToStr(ExoOuvert.Deb)) ;
  SetControlText('DATEFIN',DateToStr(ExoOuvert.Fin)) ;
  If PeriodeDate1 <> Nil Then PeriodeDate1.Text := DateToStr(RechercheFirstDate) ;

  If ExoApurger <> Nil Then
    BEGIN
    ExoApurger.ItemIndex := 0 ;
    PeriodeDate2.Text := DateToStr(DonnerDateFinExo(ExoApurger.Value)) ;
    ExoApurger.OnChange := ExoApurgerOnChange ;
    END ;
  TToolBarButton97(GetControl('BLANCER')).OnClick := BLancerOnClick ;

END ;

initialization
RegisterClasses([TOF_PURGE]) ;
end.
