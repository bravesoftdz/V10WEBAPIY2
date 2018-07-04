unit TrtRecupSuspect;

interface

uses
  SysUtils, Classes, Graphics, Controls, Forms, Dialogs,Windows,
  paramsoc,
  HTB97, Grids, Hctrls, ExtCtrls, HPanel, StdCtrls, Mask, UIUtil,
  Hent1, SaisUtil, HSysMenu, UTOB, hmsgbox, Aglinit, utobdebug, UtilGC, KpmgUtil, Dicobtp,
  EntGc     // cinClientKpmg
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS},dbtables{BDE}{$ELSE},uDbxDataSet, DBGrids, HDB,
  TntExtCtrls, TntGrids, TntStdCtrls{$ENDIF}
  {$ENDIF}
  ;

const InError = 'ERREUR';

function EntreeTrtRecupSuspect (Action : TActionFiche; StParSuspect,StFichier :string): string;
function LoadLesTobParSuspect (StParSuspect : string; TobEntete, TobLigne : TOB): boolean;

Const SAF_Objects     : integer = 0;
      SAF_Libelle     : integer = 1;
      SAF_Offset      : integer = 2;
      SAF_Longueur    : integer = 3;
      SAF_Champ       : integer = 4;
      SAF_Formule     : integer = 5;
      SAF_Test        : integer = 6;
      SAF_Erreur      : integer = 7;

// permet d'obtenir la traduction
// ajouter ds cette liste tous les libellés figurant ds la liste "LibEntrepriseParticulier"
// non traduit
      TexteMessage: array[1..10] of string 	= (
        {1}        'Civilité Contact'
        {2}        ,'Nom Contact'
        {3}        ,'Prénom Contact'
        {4}        ,'Fonction Contact'
        {5}        ,'Téléphone Contact'
        {6}        ,'Minitel'
        {7}        ,'Publipostage'
        {8}        ,'E-mail Contact'
        {9}        ,'Publipostage Contact'
        {10}       ,'eMailing Contact'
        );

  LibEntrepriseParticulier : array[1..38,1..3] of String 	= (
   ('RSU_JURIDIQUE'                 , 'Abréviat. postale'   , 'Civilité'          ),
   ('RSU_LIBELLE'                   , 'Raison Sociale'      , 'Nom'               ),
   ('RSU_PRENOM'                    , 'Suite Rs'            , 'Prénom'            ),
   ('RSU_ADRESSE1'                  , 'Adresse 1'           , 'Adresse 1'         ),
   ('RSU_ADRESSE2'                  , 'Adresse 2'           , 'Adresse 2'         ),
   ('RSU_ADRESSE3'                  , 'Adresse 3'           , 'Adresse 3'         ),
   ('RSU_CODEPOSTAL'                , 'Code postal'         , 'Code postal'       ),
   ('RSU_VILLE'                     , 'Ville'               , 'Ville'             ),
   ('RSU_PAYS'                      , 'Pays'                , 'Pays'              ),
   ('RSU_NATIONALITE'               , 'Nationalité'         , 'Nationalité'       ),
   ('RSU_LANGUE'                    , 'Langue'              , 'Langue'            ),
   ('RSU_TELEPHONE'                 , 'Téléphone'           , 'Tél domicile'      ),
   ('RSU_FAX'                       , 'Fax'                 , 'Tél bureau'        ),
   ('RSU_TELEX'                     , 'Tél portable'        , 'Tél portable'      ),
   ('RSU_RVA'                       , 'Site Web'            , 'E-mail'            ),
   ('RSU_SIRET'                     , 'Code SIRET'          , ''                  ),
   ('RSU_NIF'                       , 'Code NIF       '     , ''                  ),
   ('RSU_EAN'                       , 'Code EAN'            , ''                  ),
   ('RSU_APE'                       , 'Code NAF'            , ''                  ),
   ('RSU_JOURNAISSANCE'             , ''                    , 'Jour de naissance' ),
   ('RSU_MOISNAISSANCE'             , ''                    , 'Mois de naissance' ),
   ('RSU_ANNEENAISSANCE'            , ''                    , 'Année de naissance'),
   ('RSU_SEXE'                      , ''                    , 'Sexe'              ),
   ('RSU_COMMENTAIRE'               , 'Commentaire'         , 'Commentaire'       ),
   ('RSU_FORMEJURIDIQUE'            , 'Forme juridique'     , ''                  ),
   ('RSU_REGION'                    , 'Région'              , 'Région'            ),
   ('RSU_PUBLIPOSTAGE'              , 'Publipostage'        , 'Publipostage'      ),
   ('RSU_EMAILING'                  , 'eMailing'            , 'eMailing'          ),
   ('RSU_BLOCNOTE'                  , 'Bloc-notes'          , 'Bloc-notes'        ),
   ('RSU_CONTACTCIVILITE'           , 'Civilité Contact'    , 'Civilité Contact'  ),
   ('RSU_CONTACTNOM'                , 'Nom Contact'         , 'Nom Contact'       ),
   ('RSU_CONTACTPRENOM'             , 'Prénom Contact'      , 'Prénom Contact'    ),
   ('RSU_CONTACTFONCTION'           , 'Fonction Contact'    , 'Fonction Contact'  ),
   ('RSU_CONTACTTELEPH'             , 'Téléphone Contact'   , 'Téléphone Contact' ),
   ('RSU_CONTACTRVA'                , 'E-mail Contact'      , 'E-mail Contact'    ),
   ('RSU_CONTACTPUBLI'              , 'Publipostage Contact', 'Publipostage Contact'),
   ('RSU_ENSEIGNE'                  , 'Enseigne'            , ''                  ),
   ('RSU_CONTACTEMLG'               , 'eMailing Contact'    , 'eMailing Contact'  ));


type
  TFTrtRecupSuspect = class(TForm)
    HMTrad: THSystemMenu;
    MsgBox: THMsgBox;
    PBOUTON: TDock97;
    Pgeneral: TPanel;
    PCorps: THPanel;
    TRSS_PARSUSPECT: THLabel;
    RSS_PARSUSPECT: THCritMaskEdit;
    TRSS_FICHIER: THLabel;
    RSS_FICHIER: THCritMaskEdit;
    TRSS_SEPARATEUR: THLabel;
    RSS_SEPARATEUR: THValComboBox;
    ToolWindow971: TToolWindow97;
    BNouveau: TToolbarButton97;
    BValider: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    OpenDialogButton: TOpenDialog;
    bSupprimer: TToolbarButton97;
    RSS_Apercu: TMemo;
    RSS_FormatFichier: TRadioGroup;
    RSS_DEBUT: THNumEdit;
    RSS_FIN: THNumEdit;
    TRSS_FIN: THLabel;
    TRSS_DEBUT: THLabel;
    RSS_LIBELLE: TMaskEdit;
    TRSS_LIBELLE: THLabel;
    RSS_SEPTEXTE: THCritMaskEdit;
    TRSS_SEPTEXTE: THLabel;
    PList: TPanel;
    PFleche: TPanel;
    PGRID: TPanel;
    BFlecheDroite: TToolbarButton97;
    BFlecheGauche: TToolbarButton97;
    GChamp: THGrid;
    RSS_PARTICULIER: THValComboBox;
    TRSS_PARTICULIER: THLabel;
    TRSS_LONGENREG: THLabel;
    RSS_LONGENREG: THNumEdit;
    Boffset: TToolbarButton97;
    TLChamp: THGrid;
    GCAPERCU: THGrid;
    // Evenements de la Forme
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    // Evenements de l'entete
    procedure RSS_PARSUSPECTExit(Sender: TObject);
    procedure RSS_FICHIERExit(Sender: TObject);
    procedure RSS_SEPARATEURExit(Sender: TObject);
    // Evenements de la Grid GCHAMP
    procedure GChampCellEnter(Sender: TObject; var ACol, ARow: integer;
                              var Cancel: boolean);
    procedure GChampCellExit(Sender: TObject; var ACol, ARow: integer;
                             var Cancel: boolean);
    // Evènements de sélection de champ
    procedure BFlecheDroiteClick(Sender: TObject);
    procedure BFlecheGaucheClick(Sender: TObject);
    procedure BNouveauClick(Sender: TObject);
    procedure RSS_FICHIERElipsisClick(Sender: TObject);
    procedure BSupprimerClick(Sender: TObject);
    procedure RSS_FICHIERChange(Sender: TObject);
    procedure RSS_FormatFichierClick(Sender: TObject);
    procedure RSS_ApercuClick(Sender: TObject);
    procedure BOffsetClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BAideClick(Sender: TObject);
    procedure RSS_PARTICULIERExit(Sender: TObject);
    function ControlFormule ( LaFormule, LeChamp : String ; ARow : Integer ) : String;
    procedure FormResize(Sender: TObject);
    procedure FormKeyDown ( Sender: TObject; var Key: Word; Shift: TShiftState );
    procedure GcApercuCellEnter ( Sender : Tobject ; var ACol, ARow: integer ; var Cancel: boolean );
    procedure MajTestFormule;
    procedure GChampPostDrawCell  ( ACol, ARow : Longint; Canvas : TCanvas ; AState : TGridDrawState ) ;

  private
    { Déclarations privées }
    LesColChamp         : string ;
    iTableLigne         : integer;
    bModif              : boolean;
    AffMessErr          : boolean;
    // Gestion des Données
    function ChargeParametrage: boolean;
    function  ControleModif : boolean;
    // Initialisations
    Procedure InitialiseEntete;
    procedure InitialiseSeparateur;
    procedure InitialiseListeChamp;
    Procedure InitialiseGrid;
    // Manipulation de l'entete
    procedure AffecteEntete;
    function ControleEntete  : boolean ;
    // Manipulation des lignes
    procedure FormateZoneSaisie (ACol,ARow : Longint) ;
    procedure TraiterNombre (ACol, ARow : integer; stCol : string);
    // Validation
    procedure EnregistreTrtRecupSuspect;
    procedure RenseigneDetail;
    procedure RenseigneEntete;
    procedure ValideTrtRecupSuspect;
    procedure ApercuFichier;
    function PositionDebut : extended;
    procedure AjouteApercu ( StrData : String ; Premiere : boolean = False );
  public
    (* Déclarations publiques *)
    Action              : TActionFiche;
    Retour              : String;
    TobParSuspect       : TOB;
    TobParSuspectLig    : TOB;
  end;

var
  FTrtRecupSuspect: TFTrtRecupSuspect;

implementation
uses
   CbpMCD
   ,CbpEnumerator
;

{$R *.DFM}

function EntreeTrtRecupSuspect (Action : TActionFiche; StParSuspect,StFichier :string): string;
var FF : TFTrtRecupSuspect;
    PPANEL  : THPanel ;
begin
  SourisSablier;
  FF := TFTrtRecupSuspect.Create(Application) ;
  FF.Action := Action ;
  FF.AffMessErr := True;
  if FF.Action = TaModif then
  begin
    FF.RSS_PARSUSPECT.Text := StParSuspect;
    FF.RSS_FICHIER.Text := StFichier;
  end
  else
    FF.RSS_FICHIER.Text := GetSynRegKey('FichierSuspect',StFichier,TRUE);

  PPANEL := FindInsidePanel ; // permet de savoir si la forme dépend d'un PANEL
  if PPANEL = Nil then        // Le PANEL est le premier ecran affiché
     begin
      try
        FF.ShowModal ;
        Result:=FF.Retour ; //Retourne le nouveau paramètre saisie
      finally
        FF.Free ;
      end ;
     SourisNormale ;
     end else
     begin
     InitInside (FF, PPANEL);
     FF.Show ;
     end ;
end ;

procedure TFTrtRecupSuspect.FormCreate (Sender: TObject);
begin
  iTableLigne           := PrefixeToNum ('RRL') ;
  TobParSuspect         := TOB.Create ('PARSUSPECT', Nil, -1) ;
  TobParSuspectLig      := TOB.Create ('description', Nil, -1) ;
  GChamp.PostDrawCell   := GChampPostDrawCell ;
end;

procedure TFTrtRecupSuspect.FormShow (Sender: TObject);
begin
  InitialiseGrid;
  AffecteGrid (GChamp, Action) ;
  InitialiseEntete;
  bModif := False;
end;

{==============================================================================================}
{======================================= Initialisations ======================================}
{==============================================================================================}

Procedure TFTrtRecupSuspect.InitialiseEntete;
begin
  // TJ 02/08/2007
  // SI client KPMG, alors mise en place pour gestion des mises à jour par le code tiers.
  if GetParamSocSecur('SO_AFCLIENT', 0, True) = cInClientKPMG then
  begin
    RSS_PARTICULIER.Items.Add('Mise à jour par code tiers');
    RSS_PARTICULIER.Values.Add('MAJ');
    TRSS_PARTICULIER.Caption  := 'Ent. / Part. / Maj';
  end;
TobParSuspect.InitValeurs;
TobParSuspectLig.InitValeurs;
RSS_SEPTEXTE.Enabled := False;
RSS_SEPTEXTE.Color := clActiveBorder;
RSS_SEPTEXTE.Text := '';
RSS_LIBELLE.Text := '';
RSS_FormatFichier.ItemIndex := 1;     //FQ10505
RSS_PARTICULIER.ItemIndex := 0;
GChamp.ColWidths[SAF_Longueur] := 70;
InitialiseSeparateur;                 //FQ10505
RSS_SEPARATEUR.Value := #9;           //FQ10505
RSS_LONGENREG.Value := 0;
GChamp.VidePile(False);
GChamp.FixedRows := 1;
if (ChargeParametrage)   then
    begin
    Action := taModif;
    RSS_PARSUSPECT.Enabled := False;
    RSS_FormatFichier.enabled := False;
    if RSS_FICHIER.CanFocus then RSS_FICHIER.SetFocus;
    end else
    begin
    Action := taCreat;
    RSS_PARSUSPECT.Enabled := True;
    RSS_FormatFichier.enabled := true;
    if (RSS_PARSUSPECT.Text <> '') then
    RSS_LIBELLE.SetFocus
    else RSS_PARSUSPECT.SetFocus;
    end;
InitialiseListeChamp;
ApercuFichier;
end;

Procedure TFTrtRecupSuspect.InitialiseGrid;
var acol,arow:integer;
        Cancel : boolean;
begin
LesColChamp :='RRL_LIBELLE;RRL_OFFSET;RRL_LONGUEUR;RRL_CHAMP;RRL_FORMULE';
GChamp.ColWidths[0] := 12;
GChamp.ColWidths[SAF_Libelle] := 110;
GChamp.TabStops[SAF_Libelle] := False;
GChamp.ColAligns[SAF_Libelle]:=taLeftJustify;
if RSS_FormatFichier.ItemIndex = 0 then GChamp.ColWidths[SAF_Longueur] := 50
else GChamp.ColWidths[SAF_Longueur] := 0;
GChamp.ColTypes[SAF_Longueur] := 'R';
GChamp.ColAligns[SAF_Longueur]:=taRightJustify;
GChamp.ColFormats[SAF_Longueur] := '0';
GChamp.ColWidths[SAF_Offset] := 50;
GChamp.ColTypes[SAF_offset] := 'R';
GChamp.ColFormats[SAF_offset] := '0';
GChamp.ColAligns[SAF_offset]:=taRightJustify;
GChamp.ColWidths[SAF_Champ] := 0;
GChamp.ColLengths [SAF_Champ] := -1;
GChamp.ColTypes[SAF_Champ] := 'I';
GChamp.TabStops[SAF_Champ] := False;
GChamp.FixedCols := 1;
GChamp.ColWidths[SAF_Formule] := 300;
GChamp.TabStops[SAF_Formule] := True;
GChamp.ColAligns[SAF_Formule]:=taLeftJustify;
GChamp.ColWidths[SAF_Test] := 300;
GChamp.TabStops[SAF_Test] := False;
GChamp.ColAligns[SAF_Test]:=taLeftJustify;
GChamp.ColWidths[SAF_Erreur]  := -1;


Arow := 1; Acol :=1; Cancel:=false;
GChampCellEnter(Self,Acol,Arow,Cancel);
end;

function TFTrtRecupSuspect.ChargeParametrage: boolean;
var ind : integer;
begin
  result := False;
  if (LoadLesTobParSuspect (RSS_PARSUSPECT.Text, TobParSuspect, TobParSuspectLig)) then
  begin
    GChamp.rowcount := TobParSuspectLig.detail.count+1 ;
    for ind:=0 to TobParSuspectLig.detail.count-1 do
    begin
    TobParSuspectLig.detail[ind].PutLigneGrid (GChamp, ind+1, False, False, LesColChamp);
    end;
    GChamp.row := 1;
    AffecteEntete;
    retour:= RSS_PARSUSPECT.Text;
    result := True;
  end;
  RSS_LIBELLE.Text := TobParSuspect.GetValue ('RSS_LIBELLE');
  HMTrad.ResizeGridColumns (GChamp);
  GChamp.ColWidths [0] := 10;
end;

function LoadLesTobParSuspect (StParSuspect : string; TobEntete, TobLigne : TOB): boolean;
var Q : TQuery;
begin
result := false;
TobEntete.InitValeurs;
TobLigne.InitValeurs;
TobLigne.cleardetail;
if StParSuspect = '' then  exit;
if (TobEntete.SelectDB ('"' + StParSuspect + '"', Nil)) then
  begin
    Q := OpenSQL('SELECT * FROM PARSUSPECTLIG WHERE RRL_PARSUSPECT="'+StParSuspect+'" ORDER BY RRL_OFFSET',True) ;
    if not Q.Eof then
    begin
    TobLigne.LoadDetailDB ('PARSUSPECTLIG', '', '', Q, False);
    result := true;
    end;
    Ferme (Q);
  end;
end;

procedure TFTrtRecupSuspect.InitialiseListeChamp;
var iChamp,iTable,iColLib,ARow:integer;
Mcd : IMCDServiceCOM;
Table     : ITableCOM ;
FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  TLChamp.VidePile(False);
  if RSS_PARTICULIER.Value <> 'MAJ' then
  begin
    if (RSS_PARTICULIER.Value='PAR') then iColLib :=3 else iColLib:=2;
    for iChamp:=Low(LibEntrepriseParticulier) to High(LibEntrepriseParticulier) do
    begin
      if (LibEntrepriseParticulier [iChamp,iColLib]='') then continue;
      if not ((TLChamp.RowCount = 2) and (TLChamp.Cells[0,1] = '')) then   //1ère ligne vide
         TLChamp.RowCount := TLChamp.RowCount + 1;
      ARow := TLChamp.RowCount -1;
      TLChamp.Cells[0,ARow]:=TraduireMemoire(LibEntrepriseParticulier [iChamp,iColLib]);
      TLChamp.Cells[1,ARow]:=LibEntrepriseParticulier [iChamp,1];
    end;

    table := Mcd.getTable(mcd.PrefixetoTable('RSC'));
    FieldList := Table.Fields;
    FieldList.Reset();
    While FieldList.MoveNext do
    begin
      if ((copy((FieldList.Current as IFieldCOM).name,1,7)='RSC_RSC') or (copy((FieldList.Current as IFieldCOM).name,1,8)='RSC_6RSC'))
      and ((FieldList.Current as IFieldCOM).libelle<>'')
      and ((FieldList.Current as IFieldCOM).libelle[1]<>'.') and ((FieldList.Current as IFieldCOM).libelle<>'Error')  then
      begin
         if not ((TLChamp.RowCount = 2) and (TLChamp.Cells[0,1] = '')) then   //1ère ligne vide
            TLChamp.RowCount := TLChamp.RowCount + 1;
        ARow := TLChamp.RowCount -1;
        TLChamp.Cells[0,ARow]:=(FieldList.Current as IFieldCOM).libelle;
        TLChamp.Cells[1,ARow]:=(FieldList.Current as IFieldCOM).name;
      end;
    end;


    //Ajout des commerciaux 2 et 3
    // FQ 10467
    if GereCommercial then
    begin
      TLChamp.RowCount := TLChamp.RowCount + 1;
      ARow := TLChamp.RowCount -1;
      TLChamp.Cells[0,ARow]:='Commercial 1';
      TLChamp.Cells[1,ARow]:='RSU_REPRESENTANT';
      TLChamp.RowCount := TLChamp.RowCount + 1;
      ARow := TLChamp.RowCount -1;
      TLChamp.Cells[0,ARow]:='Commercial 2';
      TLChamp.Cells[1,ARow]:='RSC_REPRESENTANT2';
      TLChamp.RowCount := TLChamp.RowCount + 1;
      ARow := TLChamp.RowCount -1;
      TLChamp.Cells[0,ARow]:='Commercial 3';
      TLChamp.Cells[1,ARow]:='RSC_REPRESENTANT3';
    end;


    //ajout des champs pour traitements spéciaux
    TLChamp.RowCount := TLChamp.RowCount + 1;
    ARow := TLChamp.RowCount -1;
    TLChamp.Cells[0,ARow]:='*Adresse*';
    TLChamp.Cells[1,ARow]:='*Adresse*';

    TLChamp.RowCount := TLChamp.RowCount + 1;
    ARow := TLChamp.RowCount - 1;
    TLChamp.Cells[0,ARow]:='*Raison sociale*';
    TLChamp.Cells[1,ARow]:='*Raison sociale*';
  end
  else
  begin
    table := Mcd.getTable(mcd.PrefixetoTable('T'));
    FieldList := Table.Fields;
    FieldList.Reset();
    While FieldList.MoveNext do
    begin
      if (copy((FieldList.Current as IFieldCOM).name, 1, 2) = 'T_') and ((FieldList.Current as IFieldCOM).libelle <> '')
      and (Copy((FieldList.Current as IFieldCOM).libelle, 1, 1) <> '.') and ((FieldList.Current as IFieldCOM).libelle <> 'Error')  then
      begin
         if not ((TLChamp.RowCount = 2) and (TLChamp.Cells[0,1] = '')) then   //1ère ligne vide
            TLChamp.RowCount  := TLChamp.RowCount + 1;
        ARow                  := TLChamp.RowCount -1;
        TLChamp.Cells[0,ARow] := (FieldList.Current as IFieldCOM).libelle;
        TLChamp.Cells[1,ARow] := (FieldList.Current as IFieldCOM).name;
      end;
    end;

    table := Mcd.getTable(mcd.PrefixetoTable('YTC'));
    FieldList := Table.Fields;
    FieldList.Reset();
    While FieldList.MoveNext do
    begin
      if (copy((FieldList.Current as IFieldCOM).name, 1, 4) = 'YTC_') and ((FieldList.Current as IFieldCOM).libelle <> '')
      and (Copy((FieldList.Current as IFieldCOM).libelle, 1, 1) <> '.') and ((FieldList.Current as IFieldCOM).libelle <> 'Error') then
      begin
        if ((FieldList.Current as IFieldCOM).name = 'YTC_AUXILIAIRE') or ((FieldList.Current as IFieldCOM).name = 'YTC_TIERS') then
          Continue;
        if not ((TLChamp.RowCount = 2) and (TLChamp.Cells[0,1] = '')) then   //1ère ligne vide
          TLChamp.RowCount  := TLChamp.RowCount + 1;
        ARow                  := TLChamp.RowCount -1;
        TLChamp.Cells[0,ARow] := (FieldList.Current as IFieldCOM).libelle;
        TLChamp.Cells[1,ARow] := (FieldList.Current as IFieldCOM).name;
      end;
    end;

    table := Mcd.getTable(mcd.PrefixetoTable('RPR'));
    FieldList := Table.Fields;
    FieldList.Reset();
    While FieldList.MoveNext do
    begin
      if (copy((FieldList.Current as IFieldCOM).name, 1, 4) = 'RPR_') and ((FieldList.Current as IFieldCOM).libelle <> '')
      and (Copy((FieldList.Current as IFieldCOM).libelle, 1, 1) <> '.') and ((FieldList.Current as IFieldCOM).libelle <> 'Error')  then
      begin
        if ((FieldList.Current as IFieldCOM).name = 'RPR_AUXILIAIRE') then
          Continue;
        if not ((TLChamp.RowCount = 2) and (TLChamp.Cells[0,1] = '')) then   //1ère ligne vide
          TLChamp.RowCount  := TLChamp.RowCount + 1;
        ARow                  := TLChamp.RowCount -1;
        TLChamp.Cells[0,ARow] := (FieldList.Current as IFieldCOM).libelle;
        TLChamp.Cells[1,ARow] := (FieldList.Current as IFieldCOM).name;
      end;
    end;

  end;

  TLChamp.SortGrid(0,false);

end;

{==============================================================================================}
{================================ Manipulation de l'entete ====================================}
{==============================================================================================}

procedure TFTrtRecupSuspect.AffecteEntete;
var ind : integer;
begin
  RSS_LIBELLE.Text := TobParSuspect.GetValue ('RSS_LIBELLE');
  if RSS_FICHIER.Text = '' then
  RSS_FICHIER.Text := TobParSuspect.GetValue ('RSS_FICHIER');
  RSS_PARTICULIER.value := TobParSuspect.GetValue ('RSS_SUFFIXE');
  if TobParSuspect.GetValue ('RSS_TYPEENREG') = 'X' then
  begin
      RSS_FormatFichier.ItemIndex := 0;   // Longueur Fixe
      RSS_SEPARATEUR.enabled := False;
      RSS_SEPARATEUR.Color := clActiveBorder;
      RSS_LONGENREG.Value := TobParSuspect.GetValue ('RSS_LONGENREG');
      GChamp.ColWidths [SAF_Longueur] := 70;
      TRSS_DEBUT.caption := TraduireMemoire('Début Champ');
      GChamp.Cells [SAF_Offset,0] := TraduireMemoire('Début');
  end else
  begin
      RSS_FormatFichier.ItemIndex := 1;   // Délimité par un séparateur
      RSS_SEPARATEUR.enabled := True;
      RSS_SEPARATEUR.Color := clWindow;
      RSS_LONGENREG.enabled := False;

      //pb TAB
      if TobParSuspect.GetValue('RSS_SEPARATEUR') = 'TAB' then
        TobParSuspect.PutValue('RSS_SEPARATEUR', #9);
      with RSS_SEPARATEUR do
        for ind:=0 to Items.Count-1 do
          if (Values[ind] = TobParSuspect.GetValue ('RSS_SEPARATEUR')) then  ItemIndex := ind;
      GChamp.ColWidths [SAF_Longueur] := 0;
      TRSS_DEBUT.caption := TraduireMemoire('Position Champ');
      GChamp.Cells [SAF_Offset,0] := TraduireMemoire('Position');
      if RSS_SEPARATEUR.Value = 'AUT' then
          begin
          RSS_SEPTEXTE.Text := TobParSuspect.GetValue ('RSS_SEPTEXTE');
          RSS_SEPTEXTE.enabled := True;
          RSS_SEPTEXTE.Color := clWindow;
          end else
          begin
          RSS_SEPTEXTE.Enabled := False;
          RSS_SEPTEXTE.Color := clActiveBorder;
          RSS_SEPTEXTE.Text := TobParSuspect.GetValue ('RSS_SEPARATEUR');
          if RSS_SEPTEXTE.Text='' then RSS_SEPTEXTE.Text:=RSS_SEPARATEUR.value;
          end;
  end;
end;

function TFTrtRecupSuspect.ControleEntete : boolean ;
var Index, LgDet : integer;
begin
  Result := False;
  if RSS_PARSUSPECT.Text = '' then
    begin
    MsgBox.Execute (0, Caption, '') ;
    if RSS_PARSUSPECT.CanFocus then RSS_PARSUSPECT.SetFocus;
    exit;
    end;

  if (RSS_FormatFichier.ItemIndex = 1) and (RSS_SEPTEXTE.Text = '') then
    begin
      MsgBox.Execute (6, Caption, '');
      if (RSS_SEPARATEUR.text = 'AUT') then
        begin
        if RSS_SEPTEXTE.CanFocus then RSS_SEPTEXTE.SetFocus
        end
      else if RSS_SEPARATEUR.CanFocus then RSS_SEPARATEUR.SetFocus;
      exit;
    end;
  if (RSS_FormatFichier.ItemIndex = 0) then
    begin
      LgDet := 0;
      for Index := 0 to TobParSuspectLig.Detail.Count - 1 do
          LgDet := LgDet + TobParSuspectLig.Detail[Index].GetValue('RRL_LONGUEUR');
      if (RSS_LONGENREG.Value <> 0) and (LgDet > RSS_LONGENREG.Value) then
          begin
            MsgBox.Execute (17, Caption, '') ;
          end
    end;
  Result := True;
end;

{==============================================================================================}
{=============================== Manipulation des lignes ======================================}
{==============================================================================================}

procedure TFTrtRecupSuspect.FormateZoneSaisie (ACol, ARow : Longint) ;
Var St, StC : string ;
begin
St := GChamp.Cells [ACol, ARow];
StC := St ;
if (((ACol = SAF_Longueur) or (ACol = SAF_Offset)) and (St <> '0')) then
   StC := StrF00 (Valeur(St), 0);
GChamp.Cells [ACol, ARow] := StC ;
end ;

procedure TFTrtRecupSuspect.TraiterNombre (ACol, ARow : integer; stCol : String);
begin
  if ((GChamp.Cells [ACol, ARow] <> '') and (GChamp.Cells [ACol, ARow] <> '0')) then
  begin
    if TOB(GChamp.Objects[SAF_Objects, ARow])<>nil then
    TOB(GChamp.Objects[SAF_Objects, ARow]).PutValue (stCol, StrToInt (GChamp.Cells [ACol, ARow]));
    bModif := true;
  end;
end;

{==============================================================================================}
{======================================= Gestion des Données ==================================}
{==============================================================================================}

function TFTrtRecupSuspect.ControleModif : boolean;
var mrResultat : integer;
begin
Result := True;
if (RSS_LIBELLE.Text <> TobParSuspect.GetValue ('RSS_LIBELLE')) or
   (RSS_PARTICULIER.Value <> TobParSuspect.GetValue ('RSS_SUFFIXE')) then bModif := True ;
if (bModif = True) or (TobParSuspectLig.IsOneModifie) or (TobParSuspect.IsOneModifie) then
    begin
    mrResultat := MsgBox.Execute (15, Caption, '');
    if mrResultat = mrCancel then Result := False
    else
        begin
        Result := True;
        if mrResultat = mrYes then
            begin
            if ControleEntete then ValideTrtRecupSuspect
            else Result := False;
            end;
        end;
    end;

end;

{==============================================================================================}
{================================= Validation =================================================}
{==============================================================================================}

procedure TFTrtRecupSuspect.EnregistreTrtRecupSuspect;
begin
  try
    BeginTrans;
    TobParSuspect.SetAllModifie (True);
    TobParSuspect.InsertOrUpdateDB(False);
    ExecuteSQL ('DELETE FROM PARSUSPECTLIG WHERE RRL_PARSUSPECT="'+ RSS_PARSUSPECT.Text+'"');
    TobParSuspectLig.InsertDB(nil);
    SaveSynRegKey('FichierSuspect',ExtractFilePath(RSS_FICHIER.Text),TRUE);
    CommitTrans;
  except
    on E: Exception do
    begin
      RollBack;
    end;
  end;

end;

procedure TFTrtRecupSuspect.RenseigneDetail;
var i_ind :integer;
begin
if TobParSuspectLig.Detail.Count > 0 then
for i_ind := 1 to GChamp.RowCount-1 do
begin
  TOB(GChamp.Objects[SAF_Objects,i_ind]).PutValue ('RRL_PARSUSPECT', RSS_PARSUSPECT.Text);
  if (RSS_FormatFichier.ItemIndex = 1) and (GChamp.Cells[SAF_Offset, i_ind] = '0') then
    TOB(GChamp.Objects[SAF_Objects,i_ind]).PutValue ('RRL_OFFSET', i_ind)
  else TOB(GChamp.Objects[SAF_Objects,i_ind]).PutValue('RRL_OFFSET',valeurI(GChamp.Cells[SAF_Offset,i_ind]));
  TOB(GChamp.Objects[SAF_Objects,i_ind]).PutValue('RRL_LONGUEUR',valeurI(GChamp.Cells[SAF_Longueur,i_ind]));
  TOB(GChamp.Objects[SAF_Objects,i_ind]).PutValue('RRL_FORMULE',GChamp.Cells[SAF_Formule,i_ind]);
end;
end;

procedure TFTrtRecupSuspect.RenseigneEntete;
begin
TobParSuspect.PutValue ('RSS_PARSUSPECT', RSS_PARSUSPECT.Text);
TobParSuspect.PutValue ('RSS_FICHIER', RSS_FICHIER.Text);
if RSS_FormatFichier.ItemIndex = 0 then
  TobParSuspect.PutValue ('RSS_TYPEENREG', 'X')
else
  begin
  TobParSuspect.PutValue ('RSS_TYPEENREG', '-');
  RSS_LONGENREG.Value := 0;
  end;

//modif pour pb découpage avec TAB
if RSS_SEPARATEUR.Value = #9 then
  TobParSuspect.PutValue ('RSS_SEPARATEUR', 'TAB')
else
  TobParSuspect.PutValue ('RSS_SEPARATEUR', RSS_SEPARATEUR.Value);

TobParSuspect.PutValue ('RSS_LONGENREG', RSS_LONGENREG.Value);
TobParSuspect.PutValue ('RSS_LIBELLE', RSS_LIBELLE.Text);
TobParSuspect.PutValue ('RSS_SEPTEXTE', RSS_SEPTEXTE.Text);
TobParSuspect.PutValue ('RSS_SUFFIXE',RSS_PARTICULIER.Value);
Retour := RSS_PARSUSPECT.Text;  //Retour du code paramètre description
end;

procedure TFTrtRecupSuspect.ValideTrtRecupSuspect;
begin

  RenseigneDetail;
  RenseigneEntete;
  EnregistreTrtRecupSuspect;
end;

{==============================================================================================}
{=============================== Evènements de la Form ========================================}
{==============================================================================================}

procedure TFTrtRecupSuspect.FormClose (Sender: TObject; var Action: TCloseAction);
var index :integer;
begin
TobParSuspect.Free;
TobParSuspect := Nil ;
TobParSuspectLig.cleardetail;
for Index := TobParSuspectLig.Detail.Count - 1 Downto 0 do
 TobParSuspectLig.Detail[Index].Free ;
TobParSuspectLig.Free;
TobParSuspectLig := Nil ;
if IsInside(Self) then Action := caFree ;
end;

procedure TFTrtRecupSuspect.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
CanClose := ControleModif;
end;

procedure TFTrtRecupSuspect.BNouveauClick(Sender: TObject);
var bNouveau : boolean;
begin
bNouveau := ControleModif;
if bNouveau = True then
    begin
      Action := TaCreat;
      RSS_PARSUSPECT.Text := '';
      RSS_FICHIER.Text := '';
      RSS_PARSUSPECT.Enabled := True;
      TobParSuspectLig.ClearDetail;
      InitialiseGrid;
      InitialiseEntete;
      bModif := False;
    end;
end;

{==============================================================================================}
{========================= Evenement de l'Entete  =============================================}
{==============================================================================================}

procedure TFTrtRecupSuspect.RSS_FICHIERElipsisClick(Sender: TObject);
begin
  if OpenDialogButton.Execute then
    if OpenDialogButton.FileName <> '' then
    begin
      RSS_FICHIER.Text := OpenDialogButton.Filename;
    end;
end;

procedure TFTrtRecupSuspect.RSS_FICHIERChange(Sender: TObject);
begin
bModif := True;
RSS_FormatFichier.enabled := true;
AffMessErr := True;
ApercuFichier;
end;

procedure TFTrtRecupSuspect.RSS_FICHIERExit (Sender: TObject);
begin
  if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
end;

procedure TFTrtRecupSuspect.RSS_PARSUSPECTExit (Sender: TObject);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if RSS_PARSUSPECT.Text <> '' then
    begin
    RSS_PARSUSPECT.Text := Trim(UpperCase(RSS_PARSUSPECT.Text));
    if Action = TaCreat then
      begin
      if ExisteSQL ('SELECT 1 FROM PARSUSPECT WHERE RSS_PARSUSPECT="'+RSS_PARSUSPECT.Text+'"') then
        begin
        PGIInfo('Ce code de description de fichier existe déjà.', '');
        RSS_PARSUSPECT.Text := '';
        RSS_PARSUSPECT.SetFocus;
        Exit;
        end
      else ChargeParametrage;
      end
    else ChargeParametrage;
    end;
end;

procedure TFTrtRecupSuspect.RSS_SEPARATEURExit (Sender: TObject);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if not bModif then
  PGIInfoAF('Cette modification sera prise en compte lors du prochain rechargement de la description', '');

if (RSS_SEPARATEUR.Value = 'AUT') then
    begin
    RSS_SEPTEXTE.Enabled := True;
    RSS_SEPTEXTE.Color := clWindow;
    RSS_SEPTEXTE.Text := '';
    RSS_SEPTEXTE.setfocus;
    end else
    begin
    RSS_SEPTEXTE.Enabled := False;
    RSS_SEPTEXTE.Color := clActiveBorder;
    RSS_SEPTEXTE.Text := RSS_SEPARATEUR.value;
    end;
    bModif := True;
end;



{==============================================================================================}
{=============================== Evènements de la Grid ========================================}
{==============================================================================================}

procedure TFTrtRecupSuspect.GChampCellEnter (Sender: TObject; var ACol,
                                           ARow: integer; var Cancel: boolean);
begin
  if (Action = taConsult) or Cancel then Exit;
  if RSS_FormatFichier.ItemIndex = 0 then
  begin
    if (GChamp.Col = SAF_Libelle) or ( (GChamp.col <> SAF_Offset) and
       ( (GChamp.Cells [SAF_Offset, GChamp.Row] = '') or (GChamp.Cells [SAF_Offset, GChamp.Row] = '0') ) ) then
    begin
      GChamp.Col := SAF_Offset ;
    end;
  end
  {
    else
      GChamp.Col := SAF_Offset;
      }
  else
  begin
    GCAPERCU.Col :=  StrtoInt(Gchamp.Cells[SAF_Offset, Gchamp.Row]);
    RSS_DEBUT.Value := GCAPERCU.Col;

  end;

end;

procedure TFTrtRecupSuspect.GChampCellExit (Sender: TObject; var ACol, ARow: integer; var Cancel: boolean);
var
  Lerreur               : String;

begin
  if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
  FormateZoneSaisie (ACol, ARow);
  if ACol = SAF_Longueur then TraiterNombre (ACol, ARow, 'RRL_LONGUEUR');
  if ACol = SAF_Offset then TraiterNombre (ACol, ARow, 'RRL_OFFSET');
  if ACol = SAF_Formule then
  begin
{    if Trim_(GChamp.Cells[ACol, ARow]) = '' then
    begin
      Gchamp.Cells[SAF_Test, ARow] := '';
      exit;
    end;
 }
    Lerreur             := ControlFormule(GChamp.Cells[ACol, ARow], GChamp.Cells[SAF_Champ, ARow], ARow);
    if Lerreur <> '' then
    begin
      pgibox('La formule n''est pas valide !', 'Formule');
      GChamp.Col := ACol;
      GChamp.Row := ARow;
    end
    else
    begin
      if GChamp.cells[SAF_Formule, ARow] <> '' then
        begin
        if TOB(GChamp.Objects[SAF_Objects, ARow])<>nil then
        begin
          TOB(GChamp.Objects[SAF_Objects, ARow]).PutValue('RRL_FORMULE',GChamp.Cells[SAF_Formule,ARow]);
          TOB(GChamp.Objects[SAF_Objects, ARow]).SetAllModifie(True);
          bModif := true;
        end;
      end;
    end;
  end;
  MajTestFormule;
end;

{==============================================================================================}
{=============================== Evènements de sélection de champ =============================}
{==============================================================================================}

procedure TFTrtRecupSuspect.BFlecheDroiteClick(Sender: TObject);
var ARow : integer;
    StChamp,StLibelle :string;
    TOBL : Tob;
begin
  StChamp := TLChamp.Cells[1,TLChamp.Row];
  StLibelle := TLChamp.Cells[0,TLChamp.Row];
  ARow := GChamp.Row;
  TOBL:=TobParSuspectLig.FindFirst(['RRL_CHAMP'],[StChamp],False);
  if (TOBL = nil) then
  begin
    GChamp.CacheEdit ;
    GChamp.SynEnabled := False ;
    if (ARow <= TobParSuspectLig.Detail.Count) then
    begin
      inc (Arow);
      GChamp.InsertRow (ARow);
      HMTrad.ResizeGridColumns (GChamp) ;
      GChamp.Row := ARow;
    end;
    GChamp.MontreEdit; GChamp.SynEnabled := true;
    GChamp.Rows[ARow].Clear;
    TOBL := TOB.Create ('PARSUSPECTLIG', TobParSuspectLig, ARow-1) ;
    TOBL.InitValeurs;
    TOBL.PutValue ('RRL_PARSUSPECT', RSS_PARSUSPECT.Text);
    TOBL.PutValue ('RRL_CHAMP', StChamp);
    TOBL.PutValue ('RRL_LIBELLE', StLibelle);
    TOBL.PutLigneGrid (GChamp, ARow, False, False, LesColChamp);
  end;
end;

procedure TFTrtRecupSuspect.BFlecheGaucheClick(Sender: TObject);
var ARow : integer;
begin
  ARow := GChamp.Row ;
  if (ARow < 1) or (ARow > TobParSuspectLig.Detail.Count)
  or (GChamp.Cells [SAF_Libelle, ARow] = '') then Exit;
  bmodif := true;
  GChamp.CacheEdit; GChamp.SynEnabled := False;
  if TOB(GChamp.Objects[SAF_Objects, ARow])<>nil then TOB(GChamp.Objects[SAF_Objects, ARow]).Free;
  if TobParSuspectLig.Detail.Count = 0 then GChamp.Rows [ARow].Clear
  else GChamp.DeleteRow (ARow);
  if (ARow > 1) then GChamp.Row := ARow -1 ;
  GChamp.MontreEdit; GChamp.SynEnabled := true;
end;

procedure TFTrtRecupSuspect.InitialiseSeparateur;
var
  StringList: TStrings;
begin
  StringList := TStringList.Create;
  try
    with StringList do begin
      Add(TraduireMemoire('Virgule'));
      Add(TraduireMemoire('Point-virgule'));
      Add(TraduireMemoire('Tabulation'));
      Add(TraduireMemoire('Barre verticale'));
      Add(TraduireMemoire('Autre séparateur'));
    end;

    with RSS_SEPARATEUR do begin
      Items.Assign(StringList);
    end;
    with StringList do begin
      clear;
      Add(',');
      Add(';');
      Add(#9);
      Add('|');
      Add('AUT');
    end;
    with RSS_SEPARATEUR do begin
      Values.Assign(StringList);
      ItemIndex := 0;
      RSS_SEPTEXTE.Text := Values[0];
    end;

  finally
    StringList.free;
  end;
end;

procedure TFTrtRecupSuspect.BSupprimerClick(Sender: TObject);
var
    Select : string;
    mrResultat : integer;
begin
  if (RSS_PARSUSPECT.Text = '') then Exit;
  mrResultat := MsgBox.Execute (16, Caption, '');
  if (mrResultat = mrCancel) or (mrResultat = mrNo) then Exit;
  BeginTrans;
  Select := 'Delete from PARSUSPECT where RSS_PARSUSPECT="' + RSS_PARSUSPECT.Text + '"';
  ExecuteSQL(Select);
  Select := 'Delete from PARSUSPECTLIG where RRL_PARSUSPECT="' + RSS_PARSUSPECT.Text + '"';
  ExecuteSQL(Select);
  try
      CommitTrans;
      TobParSuspectLig.ClearDetail;
      RSS_FICHIER.Text := '';
      RSS_PARSUSPECT.Text := '';
      InitialiseEntete;
  except
      RollBack;
  end;
end;

procedure TFTrtRecupSuspect.ApercuFichier;
var Fichier : textfile;
    stEnreg : string;
    i_ind : integer;

begin
  if RSS_FormatFichier.ItemIndex = 1 then
    GCAPERCU.Visible := True
  else
    GCAPERCU.Visible := False;


  if (RSS_FICHIER.Text <> '') and fileexists (RSS_FICHIER.Text) then
  begin
    AssignFile (Fichier, RSS_FICHIER.Text);
    {$i-}
    Reset (Fichier);
    {$i+}
    if IoResult=0 then
    begin
      readln (Fichier, stEnreg);
      RSS_Apercu.Lines.Text := stEnreg;


      if RSS_FormatFichier.enabled then
      begin
        RSS_FormatFichier.ItemIndex := 0;
        with RSS_SEPARATEUR do
          for i_ind:=0 to Items.Count-1 do
            if (pos(Values[i_ind],stEnreg) > 0) then
            begin
            ItemIndex := i_ind;
            RSS_FormatFichier.ItemIndex := 1;
            RSS_SEPTEXTE.Text := Values[i_ind];
            end;
      end;

      if RSS_FormatFichier.ItemIndex = 1 then
        AjouteApercu ( stEnreg, True );

      if (RSS_LONGENREG.Value = 0) then
        RSS_LONGENREG.Value :=  Length (stEnreg);
      readln (Fichier, stEnreg);

      i_ind := 0;
      while not EOF(Fichier) and (i_ind < 25) do
      begin
        RSS_Apercu.Lines.Append(stEnreg);
        //partie ajoutée pour afficher une grille pour apercu
        if RSS_FormatFichier.ItemIndex = 1 then
          AjouteApercu ( stEnreg );

        if (RSS_LONGENREG.Value < Length (stEnreg)) then
           RSS_LONGENREG.Value :=  Length (stEnreg);
        readln (Fichier, stEnreg);inc(i_ind);
      end;
      RSS_Apercu.SelStart := 0; RSS_Apercu.SelLength := 1;
      if (RSS_LONGENREG.Value <> 0) then
      begin
      RSS_LONGENREG.Enabled := False;
      RSS_LONGENREG.Color := clActiveBorder;
      end;
      CloseFile(Fichier);

      MajTestFormule;
      
    end
    else
    begin
      if AffMessErr = True then
      begin
        MsgBox.Execute (18, Caption, '');
        AffMessErr := False;
      end;
      RSS_Apercu.clear;


      if action = TaCreat then RSS_FICHIER.Text := ExtractFilePath(RSS_FICHIER.Text);
    end;
  end else
      RSS_Apercu.clear;
end;

procedure TFTrtRecupSuspect.RSS_FormatFichierClick(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
  if RSS_FormatFichier.ItemIndex = 0 then
  begin
    RSS_LONGENREG.Enabled := True;
    RSS_LONGENREG.Color := clWindow;
    RSS_SEPARATEUR.enabled := False;
    RSS_SEPARATEUR.Color := clActiveBorder;
    GChamp.ColWidths [SAF_Longueur] := 50;
    GChamp.TabStops[SAF_Longueur] := True;
    GChamp.Cells [SAF_Offset,0] := TraduireMemoire('Début');
    TRSS_DEBUT.caption := TraduireMemoire('Début champ');
    GCAPERCU.Visible := False;
  end else
  begin
    RSS_SEPARATEUR.enabled := True;
    RSS_SEPARATEUR.Color := clWindow;
    RSS_LONGENREG.Enabled := False;
    RSS_LONGENREG.Color := clActiveBorder;
    GChamp.ColWidths [SAF_Longueur] := -1;
    GChamp.TabStops[SAF_Longueur] := False;
    GChamp.TabStops[SAF_Formule] := true;
    GChamp.Cells [SAF_Offset,0] := TraduireMemoire('Position');
    TRSS_DEBUT.caption := TraduireMemoire('Position champ');
    GCAPERCU.Visible := True;
  end;
  HMTrad.ResizeGridColumns (GChamp);
  GChamp.ColWidths [0] := 10;
  bModif := True;
end;

procedure TFTrtRecupSuspect.RSS_ApercuClick(Sender: TObject);
begin
RSS_DEBUT.value := PositionDebut;
RSS_FIN.value := RSS_Apercu.SelLength;
end;

procedure TFTrtRecupSuspect.BOffsetClick(Sender: TObject);
var ARow : integer;
begin
    GChamp.CacheEdit; GChamp.SynEnabled := false;
    if (GChamp.Row < 1) or (TobParSuspectLig.Detail.Count = 0)  then Exit;
    if RSS_FormatFichier.ItemIndex = 1 then
    begin
      GChamp.Cells [SAF_Offset, GChamp.Row] := IntToStr(GCAPERCU.Col);
    end else
    begin
        GChamp.Cells [SAF_Offset, GChamp.Row] := StrF00 (RSS_DEBUT.value, 0);
        GChamp.Cells [SAF_Longueur, GChamp.Row] := StrF00 (RSS_FIN.value, 0);
        TraiterNombre (SAF_Longueur, GChamp.Row, 'RRL_LONGUEUR');
    end;
    TraiterNombre (SAF_Offset, GChamp.Row, 'RRL_OFFSET');
    ARow := GChamp.Row;inc (Arow);
    if (ARow <= TobParSuspectLig.Detail.Count) then GChamp.GotoRow(ARow);
    GChamp.MontreEdit; GChamp.SynEnabled := true;
end;

function TFTrtRecupSuspect.PositionDebut: extended;
var stEnreg : string;
    I,Index,IStart,LongEnr : integer;
begin
    stEnreg := RSS_Apercu.Lines[0];
    LongEnr := Length (stEnreg); I:=0; IStart := RSS_Apercu.SelStart;
    while ((LongEnr > 0) and (IStart > LongEnr)) do
    begin
    IStart := IStart -2 - LongEnr;  inc (I);
    stEnreg := RSS_Apercu.Lines[I];
    LongEnr := Length (RSS_Apercu.Lines[I]);
    end;

    if RSS_FormatFichier.ItemIndex = 1 then
    begin
      index := 1;
      for I := 1 to LongEnr  do
      begin
        if  (stEnreg[I] = RSS_SEPTEXTE.text) then inc(Index);
        if (I = IStart) then break;
      end;
      result := Index;
    end else
    result := IStart+1;
end;

procedure TFTrtRecupSuspect.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

procedure TFTrtRecupSuspect.RSS_PARTICULIERExit(Sender: TObject);
begin
  InitialiseListeChamp;
  if (RSS_PARTICULIER.value <> TobParSuspect.GetValue ('RSS_SUFFIXE')) then bModif := True;
end;




function TFTrtRecupSuspect.ControlFormule(LaFormule, LeChamp : String ; ARow : Integer ): String;
var
  Pos1 : integer;
  Pos2 : integer;
  ChampRemp : String;
  StrRes : String;                                           
  Prefixe : String;
  Itable : integer;
  I : integer;
  LeType : String;
  Offset : String;
  StrChamp : String;
  Longeur : String;
	Mcd : IMCDServiceCOM;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

  Result                := '';
  GChamp.Cells[SAF_Erreur, ARow] := '0';

  if LaFormule = '' then exit ;
  if LaFormule = '{}' then exit;
  if LaFormule = '{-}' then exit;

  GChamp.Cells[SAF_Test, ARow] := '';
	Letype := mcd.getField(LeChamp).tipe;

  Pos1                  := Pos('{', LaFormule);
  While Pos1 <> 0 do
  begin
    Pos2                := Pos('}', LaFormule);
    if Pos2 < Pos1 then
    begin
      Result            := 'Formule incorecte';
      Break;
    end;

    ChampRemp           := Copy(LaFormule, Pos1, Pos2-Pos1+1);
    if (LeType = 'DOUBLE') or (LeType = 'INTEGER') then
    begin
      OffSet            := Copy(ChampRemp, 2, Length(ChampRemp)-2);
      if GCAPERCU.RowCount > 1 then
      begin
        if GCAPERCU.Row > 1 then
          i             := GCAPERCU.Row
        else
          i             := 1;
        LaFormule       := StringReplace_(LaFormule, ChampRemp, '(' + GCAPERCU.Cells[StrToInt(Offset), i] + ')', [rfReplaceAll]);
      end
      else
        LaFormule       := StringReplace_(LaFormule, ChampRemp, '(123)', [rfReplaceAll]);
    end
    else
    begin
      OffSet            := Copy(ChampRemp, 2, Length(ChampRemp)-2);
      if GCAPERCU.RowCount > 1 then
      begin
         if GCAPERCU.Row > 1 then
          i             := GCAPERCU.Row
        else
          i             := 1;
       StrChamp         := GCAPERCU.Cells[StrToInt(Offset), i];
       StrChamp         := StringReplace_(StrChamp, '"', '',[rfReplaceAll]);
       LaFormule        := StringReplace_(LaFormule, ChampRemp, '("'+StrChamp+'")', [rfReplaceAll]);
      end
      else
      begin

        Pos2            := Pos('-', ChampRemp);
        OffSet          := Copy(ChampRemp, 2, Pos2-2);
        Longeur         := Copy(ChampRemp, pos2+1, length(ChampRemp)-Pos2-1);
        StrChamp        := Copy(RSS_Apercu.Lines[1], StrToInt(OffSet), StrToInt(Longeur));

        LaFormule       := StringReplace_(LaFormule, ChampRemp, '("'+StrChamp+'")', [rfReplaceAll]);
      end;
    end;
    Pos1                := Pos('{', LaFormule);
  end;

  Pos1                  := Pos('EXECUTESQL', uppercase_(LaFormule));
  If Pos1 <> 0 Then
    Result              := 'EXECUTESQL interdit';

  if Result = '' then
    StrRes              := TheFormVM.GetExprValue(LaFormule)
  else
    Exit;

  if ( LeType = 'DATE' )
  and ( StrRes <> '' ) then
  begin
    if not IsValidDate(StrRes) then
      Result            := 'Type incompatible, date attendue';
  end;

  If LeType = 'BOOLEAN' then
  begin
    if (StrRes <> 'X') and (StrRes <> '-') then
      Result            := 'Type incompatible, booléen attendu';
  end;

  if ((LeType = 'DOUBLE') or (LeType = 'INTEGER'))
  and ( StrRes <> '' ) then
  begin
    if not IsNumeric(StrRes) then
      Result            := 'Type incompatible, numérique attendu'

  end;

  if Result = '' then
    GChamp.Cells[SAF_Test, ARow] := StrRes
  else
    GChamp.Cells[SAF_Erreur, ARow] := '1';

end;

procedure TFTrtRecupSuspect.AjouteApercu(StrData : String; Premiere: boolean = False);
var
  Pos1 : integer;
  Index : integer;
  StrVal : String;
  LaLigne : integer;

begin
  Index := 1;
  LaLigne := GCAPERCU.RowCount;

  if not Premiere then
  begin
    GCAPERCU.RowCount := LaLigne + 1;
    GCAPERCU.Cells[0, LaLigne] := IntToStr(LaLigne);
  end;

  Pos1 := Pos(RSS_SEPARATEUR.value, StrData);
  while Pos1 <> 0 do
  begin
    GCAPERCU.ColCount := Index + 1;

    StrVal := Copy(StrData, 1, Pos1 -1);
    StrData := Copy(StrData, Pos1+1, length(StrData));
    if Premiere then
      GCAPERCU.Cells[Index, 0] := StrVal 
    else
      GCAPERCU.Cells[Index, LaLigne] := StrVal;
    GCAPERCU.AutoResizeColumn(Index, 30);

    Inc(Index);
    Pos1 := Pos(RSS_SEPARATEUR.Value, StrData);
  end;

  GCAPERCU.ColCount := Index + 1;
  if Premiere then
    GCAPERCU.Cells[Index, 0] := StrData
  else
    GCAPERCU.Cells[Index, LaLigne] := StrData;
  GCAPERCU.AutoResizeColumn(Index, 30);


  if LaLigne = 2 then
  begin
    GCAPERCU.FixedCols := 1;
    GCAPERCU.FixedRows := 1;
  end;

end;

procedure TFTrtRecupSuspect.FormResize(Sender: TObject);
var
  i : integer;

begin
  if GCAPERCU.ColCount = 1 then
    exit;
    
  for i := 1 to GCAPERCU.ColCount-1 do
    GCAPERCU.AutoResizeColumn(i, 30);


end;


procedure TFTrtRecupSuspect.FormKeyDown ( Sender: TObject; var Key: Word; Shift: TShiftState );
var
  StrOffset : String;
  StrNvl : String;

begin
//  if ( Key = 83 ) and ( ssCtrl in Shift ) then      // CTRl-Q
  if Key = 45 then    // Insert (45)
  begin
    if RSS_FormatFichier.ItemIndex = 1 then
    begin
      StrOffset := '{' + IntToStr(GCAPERCU.Col) + '}';
      StrNvl := GChamp.Cells[SAF_Formule, GChamp.Row] + StrOffset;
      GChamp.Cells[SAF_Formule, GChamp.Row] := StrNvl;
    end;
  end;

end;


procedure TFTrtRecupSuspect.GcApercuCellEnter ( Sender: Tobject ; var ACol, ARow: integer ; var Cancel: boolean );
begin
  RSS_DEBUT.Value := GCAPERCU.Col;
  if ARow <> GCAPERCU.Row then
    MajTestFormule;
end;


procedure TFTrtRecupSuspect.MajTestFormule;
var
  i : integer;
  Lerreur               : String;

begin
  for i := 1 to GChamp.RowCount -1 do
  begin
    if Trim(GChamp.Cells[SAF_Formule, i]) <> '' then
    begin
      Lerreur           := ControlFormule(GChamp.Cells[SAF_Formule, i], GChamp.Cells[SAF_Champ, i], i);
      if  Lerreur <> '' then
      begin
         GChamp.Cells[SAF_Test, i] := Lerreur ;
      end;
    end
    else
    begin
      if RSS_FormatFichier.ItemIndex = 1 then
      begin
        Lerreur         := ControlFormule('{'+GChamp.Cells[SAF_offset, i]+'}', GChamp.Cells[SAF_Champ, i], i);
        if Lerreur <> '' then
        begin
          GChamp.Cells[SAF_Test, i] := Lerreur;
        end;
      end
      else
      begin
        if VALEURI(GChamp.Cells[SAF_Longueur, i]) <> 0 then   //FQ10505
        begin
          Lerreur       := ControlFormule('{'+GChamp.Cells[SAF_offset, i]+'-'+GChamp.Cells[SAF_Longueur, i]+'}', GChamp.Cells[SAF_Champ, i], i);
          if Lerreur <> '' then
          begin
            GChamp.Cells[SAF_Test, i] := Lerreur;
          end;
        end;
      end;


    end;
  end;

end;


procedure TFTrtRecupSuspect.GChampPostDrawCell  ( ACol, ARow : Longint; Canvas : TCanvas ; AState : TGridDrawState ) ;
var
   R : Trect ;
begin
  if ( ACol = SAF_Test )
  and ( arow > 0 ) then
  begin
     if GChamp.Cells[SAF_Erreur, Arow] = '1' then
     begin
        Canvas.Font.Color := clRed ;
//     GChamp.Invalidate
        R := GChamp.CellRect(ACol,arow);
        Canvas.TextRect(R,R.left,R.top,GChamp.Cells[Acol, Arow]);
     end;
//     Canvas.Brush.Color := clred ;
     //Canvas.TextRect(Rect,rect.Left,rect.Top, GChamp.Cells[Acol, Arow]) ;
  end
  else
     Canvas.Font.Color := clBlack ;


{  else
  begin
     GChamp.Canvas.Font.Color := clBlack ;
     GChamp.Canvas.Brush.Color := clred ;
     GChamp.Canvas.TextRect(Rect,rect.Left,rect.Top, GChamp.Cells[Acol, Arow]) ;
  end;
}
end;


end.
