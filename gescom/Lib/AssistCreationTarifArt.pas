{***********UNITE*************************************************
Auteur  ...... : Dominique BROSSET
Créé le ...... : 24/04/2001
Modifié le ... : 19/07/2001
Description .. : Assistant de création de tarif article pour un groupe d'articles
Mots clefs ... : TARIF;ARTICLE
*****************************************************************}
unit AssistCreationTarifArt;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, HTB97, StdCtrls, Hctrls, ComCtrls, ExtCtrls,
  UTOB, HPanel, Mask, HEnt1,
{$IFDEF EAGLCLIENT}
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$endIF}
  TarifUtil, UtilGC, TiersUtil
  ,EntGC // mcd14/05/03
  ;

procedure Assist_CreationTarifArt (TobA : TOB);

type
  TFAssistCreationTarifArt = class(TFAssist)
    TabSheet1: TTabSheet;
    PTITRE: THPanel;
    GBCREATION: TGroupBox;
    TINTRO: THLabel;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    PBevel2: TBevel;
    PBevel1: TBevel;
    PBevel3: TBevel;
    GBPRIX: TGroupBox;
    PRIXCOEF: THNumEdit;
    TVALEURPRIX: THLabel;
    TARRONDIPRIX: THLabel;
    ARRONDIPRIX: THValComboBox;
    GBPERIODE: TGroupBox;
    PanelFin: TPanel;
    TTextFin1: THLabel;
    TTextFin2: THLabel;
    PBevel4: TBevel;
    TRecap: THLabel;
    ListRecap: TListBox;
    DATEFIN: THCritMaskEdit;
    DATEDEBUT: THCritMaskEdit;
    TDATEDEBUT: THLabel;
    TDATEFIN: THLabel;
    HRecap: THMsgBox;
    HMsgErr: THMsgBox;
    PType: TPanel;
    RTARIFHT: TRadioButton;
    RTARIFTTC: TRadioButton;
    PType2: TPanel;
    TARIFTIERS: THValComboBox;
    RTARIFTIERS: TRadioButton;
    TIERS: THCritMaskEdit;
    RTIERS: TRadioButton;
    RARTICLE: TRadioButton;
    RPRIXNET: TRadioButton;
    RPV: TRadioButton;
    RPA: TRadioButton;
    TRPV: THLabel;
    Hdevise: THLabel;
    DEVISE: THValComboBox;
    HDepot: THLabel;
    DEPOT: THValComboBox;
    GBREMISE: TGroupBox;
    TREMISE: THLabel;
    REMISE: THCritMaskEdit;
    T_CASCADEREMISE: THLabel;
    CASCADEREMISE: THValComboBox;
    HLabel1: THLabel;
    REMISE_RESULTANTE: THCritMaskEdit;
    GBLIBELLE: TGroupBox;
    CBTYPE: TCheckBox;
    CBARTICLE: TCheckBox;
    CBTARIFTIERS: TCheckBox;
    CBTIERS: TCheckBox;
    CBMODECALCUL: TCheckBox;
    CBREMISE: TCheckBox;
    CBDEBUT: TCheckBox;
    CBFIN: TCheckBox;
    CBDEVISE: TCheckBox;
    CBDEPOT: TCheckBox;
    SEP: TEdit;
    LSEP: TLabel;
    LIBRE: THCritMaskEdit;
    HLIBRE: THLabel;
    TRPRIXNET: THLabel;
    procedure FormShow(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure DATEDEBUTExit(Sender: TObject);
    procedure DATEFINExit(Sender: TObject);
    procedure RARTICLEClick(Sender: TObject);
    procedure RTARIFTIERSClick(Sender: TObject);
    procedure REMISEChange(Sender: TObject);
    procedure RPRIXNETClick(Sender: TObject);
    procedure RTIERSClick(Sender: TObject);
    procedure RPVClick(Sender: TObject);
    procedure RPAClick(Sender: TObject);
  private
    { Déclarations privées }
    TobTarif, TobArticle : TOB ;
//    function  SpaceStr ( nb : integer) : string;
//    function  ExtractLibelle ( St : string) : string;
    procedure ListeRecap;
// Traitement
    procedure TraiterTarif ;
    procedure CreationTarif ;
    procedure CreationTarifDetail  (TobA : TOB; LIMT : longint);
    function ConstructionLibelle (stCodeArticle : string) : string;
  public
    { Déclarations publiques }
  end;

var
  FAssistCreationTarifArt: TFAssistCreationTarifArt;

implementation

{$R *.DFM}

procedure Assist_CreationTarifArt (TobA : TOB);
var
   Fo_Assist : TFAssistCreationTarifArt;
begin
     Fo_Assist := TFAssistCreationTarifArt.Create (Application);
     Fo_Assist.TobArticle:= TobA;
     Try
         Fo_Assist.ShowModal;
     Finally
         Fo_Assist.free;
     end;
end;

{=========================================================================================}
{============================= Evenements de la forme ====================================}
{=========================================================================================}
procedure TFAssistCreationTarifArt.FormShow(Sender: TObject);
begin
  inherited;
bAnnuler.Visible := True;
bSuivant.Enabled := True;
bFin.Visible := True;
bFin.Enabled := False;
CASCADEREMISE.Value := 'MIE';
DEVISE.Value := V_PGI.DevisePivot;
  //mcd 07/05/03 cacher dépot pr GI/GA
if (ctxAffaire in V_PGI.PGIContexte) and  not(ctxGCAFF in V_PGI.PGIContexte) then
   begin
   CbDepot.Visible :=False;
   CbDepot.Enabled :=False;
   Depot.Visible :=False;
   Depot.Enabled :=False;
   HDepot.Visible :=False;
   HDepot.Enabled :=False;
   end;
end;

{=========================================================================================}
{================================= Récapitulatif =========================================}
{=========================================================================================}
(* function TFAssistCreationTarifArt.SpaceStr ( nb : integer) : string;
Var St_Chaine : string ;
    i_ind : integer ;
begin
St_Chaine := '' ;
for i_ind := 1 to nb do St_Chaine:=St_chaine+' ';
Result:=St_Chaine;
end;

function TFAssistCreationTarifArt.ExtractLibelle ( St : string) : string;
Var St_Chaine : string ;
    i_pos : integer ;
begin
Result := '';
i_pos := Pos ('&', St);
if i_pos > 0 then
    begin
    St_Chaine := Copy (St, 1, i_pos - 1) + Copy (St, i_pos + 1, Length(St));
    end else St_Chaine := St;
Result := St_Chaine + ' : ';
end; *)

procedure TFAssistCreationTarifArt.ListeRecap;
var stType : string;
begin
ListRecap.Items.Clear;
ListRecap.Items.Add (PTITRE.Caption);
ListRecap.Items.Add ('');
if RTARIFHT.Checked then stType := 'HT '
else stType := 'TTC ';
if RARTICLE.Checked then
    begin
    ListRecap.Items.Add ('Création tarif ' + stType +
                            ExtractLibelle (RARTICLE.Caption));
    end else if RTARIFTIERS.Checked then
        begin
        ListRecap.Items.Add ('Création tarif ' + stType +  'Article/' +
                                ExtractLibelle (RTARIFTIERS.Caption));
        ListRecap.Items.Add (SpaceStr (4) + ExtractLibelle (RTARIFTIERS.Caption) +
                                SpaceStr (1) + TARIFTIERS.Value);
        end else
        begin
        ListRecap.Items.Add ('Création tarif ' + stType +  'Article/' +
                                ExtractLibelle (RTIERS.Caption));
        ListRecap.Items.Add (SpaceStr (4) + ExtractLibelle (RTIERS.Caption) +
                                SpaceStr (1) + TIERS.Text);
        end;

if RPRIXNET.Checked then
    begin
    ListRecap.Items.Add (SpaceStr (4) + 'Affectation d''un Prix net ' +
                            PRIXCOEF.Text);
    end else
        begin
        if RPV.Checked then
            begin
            ListRecap.Items.Add (SpaceStr (4) + 'Calcul ' + ExtractLibelle (RPV.Caption));
            end else
            begin
            ListRecap.Items.Add (SpaceStr (4) + 'Calcul ' + ExtractLibelle (RPA.Caption));
            end;
        ListRecap.Items.Add (SpaceStr (4) + 'Coefficient multiplicateur : ' + PRIXCOEF.Text);
        ListRecap.Items.Add (SpaceStr (4) + 'Arrondi selon la méthode : ' + ARRONDIPRIX.Value);
        end;

if REMISE.Text <> '' then ListRecap.Items.Add (SpaceStr (4) + 'Remise : ' + REMISE.Text);
ListRecap.Items.Add (SpaceStr (4) + 'Mode Remise : ' + RechDom ('GCCASCADEREMISE', CASCADEREMISE.Value, FALSE));

ListRecap.Items.Add ('');

ListRecap.Items.Add (SpaceStr (4) + 'Date de début : ' + DATEDEBUT.Text);
ListRecap.Items.Add (SpaceStr (4) + 'Date de fin : ' + DATEFIN.Text);
ListRecap.Items.Add (SpaceStr (4) + 'Devise : ' + DEVISE.Value);
ListRecap.Items.Add (SpaceStr (4) + 'Dépôt : ' + DEPOT.Value);

ListRecap.Items.Add ('');

ListRecap.Items.Add ('Libellé');
ListRecap.Items.Add (ConstructionLibelle (''));

end;

{=========================================================================================}
{========================== Evenements Liés aux Boutons ==================================}
{=========================================================================================}
procedure TFAssistCreationTarifArt.bSuivantClick(Sender: TObject);
Var Onglet : TTabSheet;
    St_NomOnglet : String;
    i_NumEcran : integer;
begin
Onglet := P.ActivePage;
  inherited;
st_NomOnglet := Onglet.Name;
i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;
if i_NumEcran = 0 then
    begin
    if (RTARIFTIERS.Checked) and (TARIFTIERS.Value = '') then
        begin
        HShowMessage ('0;'+PTITRE.Caption+';'+HMsgErr.Mess[1]+';W;O;O;O;','','') ;
        RestorePage ;
        P.ActivePage := PreviousPage;
        PChange (nil);
        TARIFTIERS.SetFocus;
        end;
    if RTIERS.Checked then
        begin
        if TIERS.Text = '' then
            begin
            HShowMessage ('0;'+PTITRE.Caption+';'+HMsgErr.Mess[2]+';W;O;O;O;','','') ;
            RestorePage ;
            P.ActivePage := PreviousPage;
            PChange (nil);
            TIERS.SetFocus;
            end else
        if TiersAuxiliaire(TIERS.Text) = '' then
            begin
            HShowMessage ('0;'+PTITRE.Caption+';'+HMsgErr.Mess[4]+';W;O;O;O;','','') ;
            RestorePage ;
            P.ActivePage := PreviousPage;
            PChange (nil);
            TIERS.SetFocus;
            end;
        end;
    end else if i_NumEcran = 1 then
        begin
        if (PRIXCOEF.Value = 0.0) and (REMISE.Text = '') then
            begin
            HShowMessage ('0;'+PTITRE.Caption+';'+HMsgErr.Mess[3]+';W;O;O;O;','','') ;
            RestorePage ;
            P.ActivePage := PreviousPage;
            PChange (nil);
            PRIXCOEF.SetFocus;
            end;
        end;
if (bSuivant.Enabled) then bFin.Enabled := False else bFin.Enabled := True;
if bFin.Enabled then ListeRecap;
end;

procedure TFAssistCreationTarifArt.bPrecedentClick(Sender: TObject);
begin
  inherited;
RestorePage ;
PChange(nil) ;
if (bSuivant.Enabled) then bFin.Enabled := False else bFin.Enabled := True;
end;

procedure TFAssistCreationTarifArt.bFinClick(Sender: TObject);
begin
  inherited;
TraiterTarif ;
Close ;
end;

{=========================================================================================}
{============================= Evenements Onglet 1 =======================================}
{=========================================================================================}
procedure TFAssistCreationTarifArt.RARTICLEClick(Sender: TObject);
begin
  inherited;
TARIFTIERS.Enabled := False;
TIERS.Enabled := False;
CBTARIFTIERS.Enabled := False;
CBTIERS.Enabled := False;
CBTARIFTIERS.Checked := False;
CBTIERS.Checked := False;
TARIFTIERS.Value := '';
TIERS.Text := '';
CBARTICLE.Enabled := True;
end;

procedure TFAssistCreationTarifArt.RTARIFTIERSClick(Sender: TObject);
begin
  inherited;
TIERS.Enabled := False;
TARIFTIERS.Enabled := True;
CBARTICLE.Enabled := False;
CBTIERS.Enabled := False;
CBARTICLE.Checked := False;
CBTIERS.Checked := False;
CBTARIFTIERS.Enabled := True;
TIERS.Text := '';
end;

procedure TFAssistCreationTarifArt.RTIERSClick(Sender: TObject);
begin
  inherited;
TIERS.Enabled := True;
TARIFTIERS.Enabled := False;
CBARTICLE.Enabled := False;
CBTARIFTIERS.Enabled := False;
CBARTICLE.Checked := False;
CBTARIFTIERS.Checked := False;
CBTIERS.Enabled := True;
TARIFTIERS.Value := '';
end;

procedure TFAssistCreationTarifArt.REMISEChange(Sender: TObject);
begin
  inherited;
REMISE_RESULTANTE.Text := FloatToStr(RemiseResultante (REMISE.Text));
end;

{=========================================================================================}
{============================= Evenements Onglet 2 =======================================}
{=========================================================================================}
procedure TFAssistCreationTarifArt.RPRIXNETClick(Sender: TObject);
begin
  inherited;
ARRONDIPRIX.Enabled := False;
end;

procedure TFAssistCreationTarifArt.RPVClick(Sender: TObject);
begin
  inherited;
ARRONDIPRIX.enabled := True;
end;

procedure TFAssistCreationTarifArt.RPAClick(Sender: TObject);
begin
  inherited;
ARRONDIPRIX.enabled := True;
end;

{=========================================================================================}
{============================= Evenements Onglet 3 =======================================}
{=========================================================================================}
procedure TFAssistCreationTarifArt.DATEDEBUTExit(Sender: TObject);
var stErr : string;
begin
    inherited;
    if not IsValidDate(DATEDEBUT.Text) and (DATEDEBUT.Enabled) then
        begin
        stErr:='"'+ DATEDEBUT.Text + '" ' + HMsgErr.Mess[0];
        HShowMessage('0;'+PTITRE.Caption+';'+StErr+';W;O;O;O;','','') ;
        DATEDEBUT.SetFocus;
        end;
end;

procedure TFAssistCreationTarifArt.DATEFINExit(Sender: TObject);
var stErr : string;
begin
    inherited;
    if not IsValidDate(DATEFIN.Text) and (DATEFIN.Enabled) then
        begin
        stErr:='"'+ DATEFIN.Text + '" ' + HMsgErr.Mess[0];
        HShowMessage('0;'+PTITRE.Caption+';'+StErr+';W;O;O;O;','','') ;
        DATEFIN.SetFocus;
        end;
end;

{=========================================================================================}
{================================== Traitement ===========================================}
{=========================================================================================}
procedure TFAssistCreationTarifArt.TraiterTarif ;
Var ioerr : TIOErr ;
    TOBJnal : TOB ;
    QQ : TQuery ;
    NumEvt,NbEnreg : integer ;
begin
NbEnreg:=TobArticle.detail.count ;
if NbEnreg <= 0 then exit;

NumEvt:=0 ;
TOBJnal:=TOB.Create('JNALEVENT', Nil, -1) ;
TOBJnal.PutValue('GEV_TYPEEVENT', 'TAR');
TOBJnal.PutValue('GEV_LIBELLE', PTITRE.Caption);
TOBJnal.PutValue('GEV_DATEEVENT', Date);
TOBJnal.PutValue('GEV_UTILISATEUR', V_PGI.User);
ioerr := Transactions (CreationTarif, 2);
QQ:=OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT',True,-1,'',true) ;
if Not QQ.EOF then NumEvt:=QQ.Fields[0].AsInteger ;
Ferme(QQ) ;
Inc(NumEvt) ;
TOBJnal.PutValue('GEV_NUMEVENT', NumEvt);
Case ioerr of
        oeOk  : begin
                TOBJnal.PutValue('GEV_ETATEVENT', 'OK');
                ListRecap.Items.Add('');
                ListRecap.Items.Add(IntToStr (NbEnreg) + ' ' + HRecap.Mess[4]);
                TOBJnal.PutValue('GEV_BLOCNOTE', ListRecap.Items.Text);
                end;
    oeUnknown : begin
                MessageAlerte(HRecap.Mess[2]) ;
                ListRecap.Items.Add('');
                ListRecap.Items.Add(HRecap.Mess[2]);
                TOBJnal.PutValue('GEV_ETATEVENT', 'ERR');
                end ;
    oeSaisie  : begin
                MessageAlerte(HRecap.Mess[3]) ;
                ListRecap.Items.Add('');
                ListRecap.Items.Add(HRecap.Mess[3]);
                TOBJnal.PutValue('GEV_ETATEVENT', 'ERR');
                end ;
   end ;
TOBJnal.InsertDB(Nil) ;
TOBJnal.Free ;
end;

procedure TFAssistCreationTarifArt.CreationTarif ;
Var i_ind : integer ;
    TSql (*, TSqlDim *) : TQuery ;
    // TobArtDim : TOB;
    LIMaxTarif : Longint;
begin
TSql := OpenSQL ('SELECT MAX(GF_TARIF) FROM TARIF', TRUE,-1,'',true) ;
if TSql.EOF then LIMaxTarif := 1 else LIMaxTarif := TSql.Fields[0].AsInteger + 1 ;
Ferme(TSql) ;

//TobArtDim := Tob.Create ('', nil, -1);
TobTarif := Tob.Create ('', nil, -1);
for i_ind := 0 to TobArticle.detail.count - 1 do
    begin
(*     if TobArticle.Detail[i_ind].GetValue ('GA_STATUTART') = 'GEN' then
        begin
        TSqlDim := OpenSQL ('SELECT GA_ARTICLE, GA_PAHT, GA_PVHT, GA_PVTTC FROM ARTICLE WHERE GA_STATUTART="DIM" AND ' +
                                ' GA_CODEARTICLE="' + TobArticle.Detail[i_ind].GetValue ('GA_CODEARTICLE') + '"', TRUE) ;
        TSqlDim.First;
        while not TSqlDim.Eof do
            begin
            TobArtDim.SelectDb ('', TSqlDim);
            CreationTarifDetail (TobArtDim, LIMaxTarif) ;
            LIMaxTarif := LIMaxTarif + 1;
            TSqlDim.Next;
            end;
        Ferme (TSqlDim);
        end else *)
        begin
        CreationTarifDetail (TobArticle.Detail[i_ind], LIMaxTarif) ;
        LIMaxTarif := LIMaxTarif + 1;
        end;
    end;
//TobArtDim.Free;
TobTarif.SetAllModifie (True); // DBR Fiche 10095
TobTarif.InsertOrUpdateDB(True) ;
end;

procedure TFAssistCreationTarifArt.CreationTarifDetail (TobA : TOB;
                                                        LIMT : Longint) ;
var TobT : Tob;
    Prix, dRemise : double;
begin
TobT := Tob.Create ('TARIF', TobTarif, -1);
TobT.PutValue ('GF_TARIF', LIMT);

TobT.PutValue ('GF_ARTICLE', TobA.GetValue ('GA_ARTICLE'));

if RTARIFTIERS.Checked = True then
    begin
    TobT.PutValue ('GF_TARIFTIERS', TARIFTIERS.Value);
    end else if RTIERS.Checked = True then
        begin
        TobT.PutValue ('GF_TIERS', TIERS.Text);
        TobT.PutValue ('GF_NATUREAUXI', 'CLI');
        end;

TobT.PutValue ('GF_LIBELLE', ConstructionLibelle (TobA.GetValue ('GA_ARTICLE')));

if RPRIXNET.Checked = True then
    begin
    TobT.PutValue ('GF_PRIXUNITAIRE', PRIXCOEF.Value);
    Prix := PRIXCOEF.Value;
    end else if RPA.Checked = True then
        begin
        Prix := TobA.GetValue ('GA_PAHT') * PRIXCOEF.Value;
        Prix := ArrondirPrix (ARRONDIPRIX.Value, Prix);
        Tobt.PutValue ('GF_PRIXUNITAIRE', Prix) ;
        end else
        begin
        if RTARIFHT.Checked = True then Prix := TobA.GetValue ('GA_PVHT') * PRIXCOEF.Value
        else Prix := TobA.GetValue ('GA_PVTTC') * PRIXCOEF.Value;
        Prix := ArrondirPrix (ARRONDIPRIX.Value, Prix);
        Tobt.PutValue ('GF_PRIXUNITAIRE', Prix) ;
        end;

TobT.PutValue ('GF_QUALIFPRIX', 'GRP');

TobT.PutValue ('GF_BORNEINF', -999999);
TobT.PutValue ('GF_BORNESUP', 999999);

TobT.PutValue ('GF_DATEDEBUT', StrToDate(DATEDEBUT.Text));
TobT.PutValue ('GF_DATEFIN', StrToDate(DATEFIN.Text));

if DEVISE.Value = '' then DEVISE.Value := V_PGI.DevisePivot;
TobT.PutValue ('GF_DEVISE', DEVISE.Value);

if DEPOT.Value <> '' then TobT.PutValue ('GF_DEPOT', DEPOT.Value);

TobT.PutValue ('GF_SOCIETE', V_PGI.CodeSociete) ;

dRemise := RemiseResultante (REMISE.Text);
TobT.PutValue ('GF_REMISE', dRemise) ;
TobT.PutValue ('GF_CALCULREMISE', REMISE.Text) ;

CalcPriorite (TobT);
TobT.PutValue ('GF_CASCADEREMISE', CASCADEREMISE.Value);
TobT.PutValue ('GF_MODECREATION', 'AUT');
TobT.PutValue ('GF_FERME', '-');
TobT.PutValue ('GF_QUANTITATIF', '-');

if RTARIFHT.Checked = True then TobT.PutValue ('GF_REGIMEPRIX', 'HT')
else TobT.PutValue ('GF_REGIMEPRIX', 'TTC');
end;

function TFAssistCreationTarifArt.ConstructionLibelle (stCodeArticle : string): string;
begin
Result := LIBRE.Text;
if CBTYPE.Checked = True then
    begin
    if Result <> '' then Result := Result + SEP.Text;
    if RTARIFHT.Checked = True then Result := Result + 'TARIF HT'
    else Result := Result + 'TARIF TTC';
    end;
if CBARTICLE.Checked = True then
    begin
    if Result <> '' then Result := Result + SEP.Text;
    if stCodeArticle = '' then Result := Result + 'Code Article'
    else Result := Result + stCodeArticle;
    end;
if CBTARIFTIERS.Checked = True then
    begin
    if Result <> '' then Result := Result + SEP.Text;
    Result := Result + RechDom ('TTTARIFCLIENT', TARIFTIERS.Value, False);
    end;
if CBTIERS.Checked = True then
    begin
    if Result <> '' then Result := Result + SEP.Text;
    Result := Result + TIERS.Text;
    end;
if (CBMODECALCUL.Checked = True) and (PRIXCOEF.Value <> 0) then
    begin
    if Result <> '' then Result := Result + SEP.Text;
    if RPRIXNET.Checked = True then Result := Result + 'Prix Net'
    else begin
        if RPV.Checked = True then Result := Result + 'Prix de vente * '
        else Result := Result + 'Prix d''achat * ';
        Result := Result + FloatToStr (PRIXCOEF.Value);
        if ARRONDIPRIX.Value <> '' then
            Result := Result + '- arrondi : ' + RechDom ('GCCODEARRONDI', ARRONDIPRIX.Value, False);
        end;
    end;
if CBREMISE.Checked = True then
    begin
    if Result <> '' then Result := Result + SEP.Text;
    Result := Result + 'Remise : ' + REMISE_RESULTANTE.Text;
    end;
if CBDEBUT.Checked = True then
    begin
    if Result <> '' then Result := Result + SEP.Text;
    Result := Result + 'du ' + DATEDEBUT.Text;
    end;
if CBFIN.Checked = True then
    begin
    if (Result <> '') and (CBDEBUT.Checked <> True) then Result := Result + SEP.Text;
    Result := Result + 'au ' + DATEFIN.Text;
    end;
if CBDEVISE.Checked = True then
    begin
    if Result <> '' then Result := Result + SEP.Text;
    Result := Result + DEVISE.Value;
    end;
if CBDEPOT.Checked = True then
    begin
    if Result <> '' then Result := Result + SEP.Text;
    Result := Result + DEPOT.Value;
    end;
if Result = '' then Result := 'Tarif article'
end;

end.
