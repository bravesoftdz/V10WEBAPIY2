{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/04/2002
Modifié le ... :   /  /
Description .. : Génération des paramètres pour les caisses CASH 2000
Mots clefs ... : 
*****************************************************************}
unit GenereBE;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, ExtCtrls, StdCtrls, ComCtrls, HRichEdt, HRichOLE, Hctrls, Mask,
  HEnt1, HPanel, UIUtil, UTOB, dbtables, ToxCash2000;

type
  TFGenereBE = class(TForm)
    CodeSite: THCritMaskEdit;
    LCodeSite: THLabel;
    NumeroCaisse: THCritMaskEdit;
    LNumeroCaisse: THLabel;
    CompteRendu: THRichEditOLE;
    Panel4: TPanel;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    CSQL: THRichEditOLE;
    LCSQL: THLabel;
    BVoir: TToolbarButton97;
    BFirst: TToolbarButton97;
    BPrecedent: TToolbarButton97;
    BSuivant: TToolbarButton97;
    BLast: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BVoirClick(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BPrecedentClick(Sender: TObject);
    procedure BSuivantClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
  private
    { Déclarations privées }
    RecentSQL : TOB ;                  // historique des requêtes SQL
    RangSQL   : Integer ;              // rang de l'élément courant dans l'historique des requêtes SQL
    procedure ChargeHistoriqueSQL ;
    procedure InsereHistoriqueSQL ( sSQL : String ) ;
    procedure AfficheHistoriqueSQL ( Indice : Integer ) ;
  public
    { Déclarations publiques }
  end;

Procedure LanceGenereBE ;

implementation

{$R *.DFM}


{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 29/03/2002
Modifié le ... : 26/04/2002
Description .. : Génération des paramètres pour les caisses CASH 2000
Mots clefs ... :
*****************************************************************}
Procedure LanceGenereBE ;
var XX  : TFGenereBE ;
    PP  : THPanel ;
BEGIN
// Affichage de l'écran de saisie
PP := FindInsidePanel ;
XX := TFGenereBE.Create(Application) ;
if PP = Nil then
   BEGIN
   try
      XX.ShowModal ;
    finally
      XX.Free ;
    end ;
   END else
   BEGIN
   InitInside(XX, PP) ;
   XX.Show ;
   END ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  ChargeHistoriqueSQL : chargement de l'historique des requêtes SQL
///////////////////////////////////////////////////////////////////////////////////////
procedure TFGenereBE.ChargeHistoriqueSQL ;
BEGIN
InsereHistoriqueSQL('Select * From ARRONDI') ;
InsereHistoriqueSQL('Select * From COMMERCIAL') ;
InsereHistoriqueSQL('Select * From DEVISE') ;
InsereHistoriqueSQL('Select * From MODEPAIE') ;
InsereHistoriqueSQL('Select * From PAYS') ;
InsereHistoriqueSQL('Select * From CODEPOST') ;
InsereHistoriqueSQL('Select * From CHOIXCOD Where CC_TYPE In ("GTM","CIV","JUR")') ;
InsereHistoriqueSQL('Select * From CHOIXEXT Where YX_TYPE Like "LT%"') ;
InsereHistoriqueSQL('Select * From TIERS Where T_NATUREAUXI="CLI"') ;
InsereHistoriqueSQL('Select * From ETABLISS') ;
InsereHistoriqueSQL('Select * From ARTICLE Where GA_TYPEARTICLE In ("PRE","FI","MAR")') ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  InsereHistoriqueSQL : insere une requête dans l'historique des requêtes SQL
///////////////////////////////////////////////////////////////////////////////////////
procedure TFGenereBE.InsereHistoriqueSQL ( sSQL : String ) ;
Var TOBL : TOB ;
BEGIN
if RecentSQL <> Nil then
   BEGIN
   TOBL := TOB.Create('', RecentSQL, -1) ;
   TOBL.AddChampSupValeur('SQL', sSQL) ;
   RangSQL := TOBL.GetIndex ;
   END ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  AfficheHistoriqueSQL : affiche une requête de l'historique des requêtes SQL
///////////////////////////////////////////////////////////////////////////////////////
procedure TFGenereBE.AfficheHistoriqueSQL ( Indice : Integer ) ;
Var TOBL : TOB ;
BEGIN
if (RecentSQL <> Nil) and (RecentSQL.Detail.Count > 0) then
   BEGIN
   if Indice = 99 then
      BEGIN
      RangSQL := RecentSQL.Detail.Count -1 ;
      END else
   if Indice = -99 then
      BEGIN
      RangSQL := 0 ;
      END else
      BEGIN
      Inc(RangSQL, Indice) ;
      if RangSQL < 0 then RangSQL := 0 ;
      if RangSQL > (RecentSQL.Detail.Count -1) then RangSQL := RecentSQL.Detail.Count -1 ;
      END ;
   TOBL := RecentSQL.Detail[RangSQL] ;
   CSQL.Text := TOBL.GetValue('SQL') ;
   END ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  FormCreate
///////////////////////////////////////////////////////////////////////////////////////
procedure TFGenereBE.FormCreate(Sender: TObject);
BEGIN
RecentSQL := TOB.Create('', Nil, -1) ;
RangSQL := 0 ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  FormShow
///////////////////////////////////////////////////////////////////////////////////////
procedure TFGenereBE.FormShow(Sender: TObject);
begin
  CSQL.Clear ;
  CompteRendu.Clear ;
  ChargeHistoriqueSQL ;
  // Appel de la fonction d'empilage dans la liste des fiches
  AglEmpileFiche(Self) ;
end ;

///////////////////////////////////////////////////////////////////////////////////////
//  FormClose
///////////////////////////////////////////////////////////////////////////////////////
procedure TFGenereBE.FormClose(Sender: TObject; var Action: TCloseAction);
BEGIN
if RecentSQL <> Nil then RecentSQL.Free ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  BValiderClick
///////////////////////////////////////////////////////////////////////////////////////
procedure TFGenereBE.BValiderClick(Sender: TObject);
Var Tox          : TOB ;
    ToxCondition : TOB ;
    ToxTable0	 : TOB ;
    NomFichier   : String ;
    NomTable0	 : String ;
    sSQL	 : String ;
    QQ           : TQuery ;
    Debut        : Integer ;
    Fin          : Integer ;
BEGIN
Inherited;
if not BValider.Enabled then Exit ;
if CodeSite.Text = '' then Exit ;
if NumeroCaisse.Text = '' then Exit ;
if CSQL.Text = '' then Exit ;
//
// Choix de la table
//
sSQL := UpperCase(Trim(CSQL.Text)) ;
if Copy(sSQL, 1, 6) <> 'SELECT' then Exit ;
Debut := Pos('FROM ', sSQL) ;
if Debut = 0 then Exit ;
Inc(Debut, 5) ;
for Fin := Debut to Length(sSQL) do if sSQL[Fin] = ' ' then Break ;
NomTable0 := Copy(sSQL, Debut, Fin-Debut) ;
//
// Creation de la TOX
//
BValider.Enabled := False ;
SourisSablier ;
Tox := Nil ;
ToxTable0 := Nil ;
QQ := OpenSQL(sSQL, True) ;
if not QQ.Eof then
begin
  Tox := TOB.Create('', Nil, -1) ;
  Tox.AddChampSup('EVENEMENT', False) ;
  ToxCondition := TOB.Create('', Tox, -1) ;
  ToxCondition.AddChampSup('CONDITION', False) ;
  ToxTable0 := TOB.Create('', ToxCondition, -1) ;
  ToxTable0.AddChampSupValeur('TOX_TABLE', NomTable0) ;
  ToxTable0.LoadDetailDB(NomTable0, '', '', QQ, False) ;
  InsereHistoriqueSQL(CSQL.Text) ;
end;
Ferme(QQ) ;
//
// Génération du fichier BE
//
if Tox <> Nil then
begin
  CompteRendu.Clear ;
  CompteRendu.lines.Add('Génération des '+ IntToStr(ToxTable0.Detail.Count) +' éléments de la table ' + NomTable0) ;
  NomFichier := ToxNomFichierBE (CodeSite.Text, NumeroCaisse.Text, 0) ;
  CompteRendu.lines.Add('dans le fichier ' + NomFichier) ;
  AvantArchivageToxPourCASH (Tox, CodeSite.Text, NumeroCaisse.Text) ;
  CompteRendu.lines.Add('Génération terminée.') ;
  Tox.Free ;
end;
SourisNormale ;
BValider.Enabled := True ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  BVoirClick
///////////////////////////////////////////////////////////////////////////////////////
procedure TFGenereBE.BVoirClick(Sender: TObject);
Var NomFichier   : String ;
BEGIN
if CodeSite.Text = '' then Exit ;
if NumeroCaisse.Text = '' then Exit ;
NomFichier := ToxNomFichierBE (CodeSite.Text, NumeroCaisse.Text, 0) ;
if FileExists(NomFichier) then CompteRendu.lines.LoadFromFile(NomFichier) ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  BFirstClick
///////////////////////////////////////////////////////////////////////////////////////
procedure TFGenereBE.BFirstClick(Sender: TObject);
BEGIN
AfficheHistoriqueSQL(-99) ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  BPrecedentClick
///////////////////////////////////////////////////////////////////////////////////////
procedure TFGenereBE.BPrecedentClick(Sender: TObject);
BEGIN
AfficheHistoriqueSQL(-1) ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  BSuivantClick
///////////////////////////////////////////////////////////////////////////////////////
procedure TFGenereBE.BSuivantClick(Sender: TObject);
BEGIN
AfficheHistoriqueSQL(1) ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  BLastClick
///////////////////////////////////////////////////////////////////////////////////////
procedure TFGenereBE.BLastClick(Sender: TObject);
BEGIN
AfficheHistoriqueSQL(99) ;
END ;

procedure TFGenereBE.FormDestroy(Sender: TObject);
begin
  // Appel de la fonction de dépilage dans la liste des fiches
  AglDepileFiche ;
end;

end.
