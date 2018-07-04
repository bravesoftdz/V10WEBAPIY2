unit AssistDupTarif;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, StdCtrls, HPanel, ComCtrls, HSysMenu, hmsgbox, HTB97, ExtCtrls,
{$IFDEF EAGLCLIENT}
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  HEnt1, Hctrls, Mask, AGLInitGc, UTob, TarifUtil, UtilGC;

  procedure Assist_DupTarif (TOBT : TOB);

type
  TFAssistDupTarif = class(TFAssist)
    TabSheet1: TTabSheet;
    PTITRE: THPanel;
    TINTRO: THLabel;
    TabSheet6: TTabSheet;
    PanelFin: TPanel;
    TTextFin1: THLabel;
    TTextFin2: THLabel;
    TRecap: THLabel;
    ListRecap: TListBox;
    ART: THCritMaskEdit;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    TabSheet5: TTabSheet;
    GroupBox3: TGroupBox;
    TCGPRIX: THLabel;
    TCBDATE: THLabel;
    CB_PRIX: TCheckBox;
    CB_DATE: TCheckBox;
    PBevel2: TBevel;
    GBPRIX: TGroupBox;
    TTYPEMAJPRIX: THLabel;
    TVALEURPRIX: THLabel;
    TARRONDIPRIX: THLabel;
    TYPEMAJPRIX: THValComboBox;
    VALEURPRIX: THNumEdit;
    ARRONDIPRIX: THValComboBox;
    GBREMISE: TGroupBox;
    TTYPEMAJREMISE: THLabel;
    TVALEURREMISE: THLabel;
    TYPEMAJREMISE: THValComboBox;
    VALEURREMISE: THNumEdit;
    DATEEFFET: THCritMaskEdit;
    TDATEEFFET: THLabel;
    GBPERIODE: TGroupBox;
    TTYPEMAJPERIODE: THLabel;
    TNBREJOUR: THLabel;
    TDATEDEBUT: THLabel;
    TDATEFIN: THLabel;
    TYPEMAJPERIODE: THValComboBox;
    NBREJOUR: THNumEdit;
    UpDownJour: TUpDown;
    CBDATEDEBUT: TCheckBox;
    CBDATEFIN: TCheckBox;
    DATEFIN: THCritMaskEdit;
    DATEDEBUT: THCritMaskEdit;
    GroupBox2: TGroupBox;
    TGF_ARTICLE: THLabel;
    TGF_TIERS: THLabel;
    TGF_DEPOT: THLabel;
    TGF_TARIFARTICLE: THLabel;
    TGF_TARIFTIERS: THLabel;
    GF_ARTICLE: THCritMaskEdit;
    GF_TIERS: THCritMaskEdit;
    GF_DEPOT: THValComboBox;
    GF_TARIFARTICLE: THValComboBox;
    GF_TARIFTIERS: THValComboBox;
    HMsgErr: THMsgBox;
    HRecap: THMsgBox;
    CBBasculeTTC: TCheckBox;
    CBBasculeHT: TCheckBox;
    TIERS: THCritMaskEdit;
    procedure FormShow(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure GF_ARTICLEElipsisClick(Sender: TObject);
    procedure GF_TIERSElipsisClick(Sender: TObject);
    procedure CB_PRIXClick(Sender: TObject);
    procedure CB_DATEClick(Sender: TObject);
    procedure TYPEMAJPRIXChange(Sender: TObject);
    procedure TYPEMAJREMISEChange(Sender: TObject);
    procedure TYPEMAJPERIODEChange(Sender: TObject);
    procedure CBDATEDEBUTClick(Sender: TObject);
    procedure CBDATEFINClick(Sender: TObject);
    procedure DATEEFFETExit(Sender: TObject);
    procedure DATEDEBUTExit(Sender: TObject);
    procedure DATEFINExit(Sender: TObject);
  private
    { Déclarations privées }
    TOBTarif : TOB ;
    procedure TraiterTarif ;
    procedure MajTarif ;
    procedure MajTarifDetail  (TOBD : TOB);
    procedure ModifPrix (Var TOBD : TOB; ModifDateEffet : Boolean) ;
    procedure ModifRemise (Var TOBD : TOB; ModifDateEffet : Boolean) ;
    procedure ModifPeriode (Var TOBD : TOB) ;
    procedure ListeRecap;
  public
    { Déclarations publiques }
  end;

var
  FAssistDupTarif: TFAssistDupTarif;
  i_NumEcran : integer;

implementation

{$R *.DFM}

procedure Assist_DupTarif (TOBT : TOB);
var
   Fo_Assist : TFAssistDupTarif;
Begin
     Fo_Assist := TFAssistDupTarif.Create (Application);
     Fo_Assist.TOBTarif:= TOBT;
     Try
         Fo_Assist.ShowModal;
     Finally
         Fo_Assist.free;
     End;
end;

procedure TFAssistDupTarif.FormShow(Sender: TObject);
var
   stArt, stTiers, stDep, stTarArt, stTarTiers : string;
   i_ind1 : integer;
   bMulArt, bMulTiers, bMulDep, bMulTarArt, bMulTarTiers : boolean;
   bTTC, bHT : boolean;
   TobTemp : TOB;

begin
  inherited;
bMulArt := False;
bMulTiers := False;
bMulDep := False;
bMulTarArt := False;
bMulTarTiers := False;
bTTC := False;
bHT := False;
bFin.Visible := True;
bFin.Enabled := False;
for i_ind1 := 0 to TOBTarif.Detail.Count - 1 do
    begin
    TobTemp := TOBTarif.Detail[i_ind1];
    if i_ind1 = 0 then
        begin
        stArt := TobTemp.GetValue('GF_ARTICLE');
        stTiers := TobTemp.GetValue('GF_TIERS');
        stDep := TobTemp.GetValue('GF_DEPOT');
        stTarArt := TobTemp.GetValue('GF_TARIFARTICLE');
        stTarTiers := TobTemp.GetValue('GF_TARIFTIERS');
        if TobTemp.GetValue ('GF_REGIMEPRIX') = 'TTC' then bTTC := True;
        if TobTemp.GetValue ('GF_REGIMEPRIX') = 'HT' then bHT := True;
        Continue;
        end;
    if stArt <> TobTemp.GetValue('GF_ARTICLE') then bMulArt := True;
    if stTiers <> TobTemp.GetValue('GF_TIERS') then bMulTiers := True;
    if stDep <> TobTemp.GetValue('GF_DEPOT') then bMulDep := True;
    if stTarArt <> TobTemp.GetValue('GF_TARIFARTICLE') then bMulTarArt := True;
    if stTarTiers <> TobTemp.GetValue('GF_TARIFTIERS') then bMulTarTiers := True;
    if (TobTemp.GetValue ('GF_REGIMEPRIX') <> 'TTC') then bTTC := False;
    if (TobTemp.GetValue ('GF_REGIMEPRIX') <> 'HT') then bHT := False;
    end;
GF_ARTICLE.Text := '<<Inchangé>>';
GF_TIERS.Text := '<<Inchangé>>';
GF_DEPOT.Items.Insert(0,'<<Inchangé>>');
GF_DEPOT.Values.Insert(0, '');
GF_DEPOT.ItemIndex := 0;
GF_TARIFARTICLE.Items.Insert(0,'<<Inchangé>>');
GF_TARIFARTICLE.Values.Insert(0, '');
GF_TARIFARTICLE.ItemIndex := 0;
GF_TARIFTIERS.Items.Insert(0,'<<Inchangé>>');
GF_TARIFTIERS.Values.Insert(0, '');
GF_TARIFTIERS.ItemIndex := 0;
if (bMulArt) or (stArt = '') then
    begin
    GF_ARTICLE.Enabled := False;
    if stArt <> '' then bMulTarArt := True;
    end;
if (bMulTiers) or (stTiers = '') then
    begin
    GF_TIERS.Enabled := False;
    if stTiers <> '' then bMulTarTiers := True;
    end;
if bMulDep then GF_DEPOT.Enabled := False;
if (bMulTarArt) or (stTarArt = '') then GF_TARIFARTICLE.Enabled := False;
if (bMulTarTiers) or (stTarTiers = '') then GF_TARIFTIERS.Enabled := False;
If (ctxAffaire in V_PGI.PGIContexte) then begin
   GF_DEPOT.Enabled := False;
   GF_DEPOT.visible := False;
   TGF_DEPOT.visible := False;
   end;

if (not bMulArt) and (stArt = '') then GBPRIX.Visible := false;

CBBasculeTTC.Checked := False;
CBBasculeHT.Checked := False;
CBBasculeHT.Enabled := bTTC; // on a que des TTC donc on peut basculer en HT
CBBasculeTTC.Enabled := bHT; // on a que des HT donc on peut basculer en TTC
end;

procedure TFAssistDupTarif.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if (Screen.ActiveControl = GF_ARTICLE) or (Screen.ActiveControl = GF_TIERS)then
    Case Key of
        VK_DELETE : BEGIN
                    Key := 0 ;
                    THCritMaskEdit(Screen.ActiveControl).Text := '<<Inchangé>>';
                    END;
        END;
end;

procedure TFAssistDupTarif.bSuivantClick(Sender: TObject);
Var Onglet : TTabSheet;
    St_NomOnglet : String;
begin
  inherited;
Onglet := P.ActivePage;
st_NomOnglet := Onglet.Name;
i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;
if i_NumEcran <= 5 then NBREJOUR.Value:=StrToFloat(NBREJOUR.Text);
if (CB_PRIX.Checked = False) and (i_NumEcran = 3) then
    BEGIN
    RestorePage ;
    Onglet := NextPage ;
    if Onglet = nil then P.SelectNextPage(True) else BEGIN P.ActivePage := Onglet ; PChange(nil) ; END ;
    NBREJOUR.Value:=StrToFloat (NBREJOUR.Text);
    END;
st_NomOnglet := Onglet.Name;
i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;
if (CB_DATE.Checked = False) and (i_NumEcran = 4) then
    BEGIN
    RestorePage ;
    Onglet := NextPage ;
    if Onglet = nil then P.SelectNextPage(True) else BEGIN P.ActivePage := Onglet ; PChange(nil) ; END ;
    END;
if (bSuivant.Enabled) then bFin.Enabled := False else bFin.Enabled := True;
if bFin.Enabled then ListeRecap;
end;

procedure TFAssistDupTarif.bPrecedentClick(Sender: TObject);
Var Onglet : TTabSheet;
    St_NomOnglet : String;
begin
  inherited;
Onglet := P.ActivePage;
st_NomOnglet := Onglet.Name;
i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;
if i_NumEcran = 4 then NBREJOUR.Value:=StrToFloat(NBREJOUR.Text);
if (CB_DATE.Checked = False) and (i_NumEcran = 4) then
    BEGIN
    RestorePage ;
    Onglet := PreviousPage ;
    if Onglet = nil then P.SelectNextPage(True) else BEGIN P.ActivePage := Onglet ; PChange(nil) ; END ;
    END;
st_NomOnglet := Onglet.Name;
i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;
if (CB_PRIX.Checked = False) and (i_NumEcran = 3) then
    BEGIN
    RestorePage ;
    Onglet := PreviousPage ;
    if Onglet = nil then P.SelectNextPage(True) else BEGIN P.ActivePage := Onglet ; PChange(nil) ; END ;
    END;
if (bSuivant.Enabled) then bFin.Enabled := False else bFin.Enabled := True;
end;

procedure TFAssistDupTarif.bFinClick(Sender: TObject);
var
    i_ind1, i_compteur : integer;
    TOBTemp : TOB;
    Q : TQuery ;
    MaxTarif : Longint ;

begin
  inherited;
Q := OpenSQL('SELECT MAX(GF_TARIF) FROM TARIF', TRUE,-1,'',true);
if Q.EOF then MaxTarif := 0 else MaxTarif := Q.Fields[0].AsInteger;
Ferme(Q) ;
i_compteur := 1;
for i_ind1 := 0 to TOBTarif.Detail.Count - 1 do
    begin
    TOBTemp := TOBTarif.Detail[i_ind1];
    if TOBTemp.GetValue('GF_REGIMEPRIX') <> 'GLO' then
        begin
        if TOBTemp.GetValue ('GF_ARTICLE') <> '' then
            if GF_ARTICLE.Text <> '<<Inchangé>>' then TOBTemp.PutValue('GF_ARTICLE', ART.Text);
        end
        else begin
        if TOBTemp.GetValue ('GF_TARIFARTICLE') <> '' then
            if GF_TARIFARTICLE.Text <> '<<Inchangé>>' then TOBTemp.PutValue('GF_TARIFARTICLE', GF_TARIFARTICLE.Value);
        end;
    if TOBTemp.GetValue ('GF_TIERS') <> '' then
        if GF_TIERS.Text <> '<<Inchangé>>'   then TOBTemp.PutValue('GF_TIERS', TIERS.Text);
    if TOBTemp.GetValue ('GF_TARIFTIERS') <> '' then
        if GF_TARIFTIERS.Text <> '<<Inchangé>>'   then TOBTemp.PutValue('GF_TARIFTIERS', GF_TARIFTIERS.Value);
    if GF_DEPOT.Text <> '<<Inchangé>>' then TOBTemp.PutValue('GF_DEPOT', GF_DEPOT.Value);

    if CBBasculeTTC.Checked = true then TOBTemp.PutValue ('GF_REGIMEPRIX', 'TTC');
    if CBBasculeHT.Checked = true then TOBTemp.PutValue ('GF_REGIMEPRIX', 'HT');

    if TOBTemp.IsOneModifie then
    begin
        TOBTemp.PutValue('GF_TARIF', MaxTarif + i_compteur);
        Inc (i_compteur);
        TOBTemp.SetAllModifie(True);
    end;
    end;
if TOBTarif.IsOneModifie then
    begin
    BeginTrans;
    TOBTarif.InsertDB(nil);
    Try
        CommitTrans;
    Except
        RollBack ;
    end;
    end;
TraiterTarif ;
Close;
end;

procedure TFAssistDupTarif.GF_ARTICLEElipsisClick(Sender: TObject);
begin
  inherited;
ART.DataType := 'GCARTICLE';
ART.Text := '';
DispatchRecherche(ART, 1, '', 'GA_CODEARTICLE=' + Trim(Copy(ART.Text, 1, 18)), '');
if ART.Text <> '' then GF_ARTICLE.Text := copy(ART.Text, 0, Length(ART.Text) - 1);
end;

procedure TFAssistDupTarif.GF_TIERSElipsisClick(Sender: TObject);
begin
  inherited;
TIERS.DataType := 'GCTIERSCLI';
TIERS.Text := '';
DispatchRecherche(TIERS, 2, 'T_NATUREAUXI="CLI"','T_TIERS=' + Trim(GF_TIERS.Text), '');
if TIERS.Text <> '' then GF_TIERS.Text := TIERS.Text;
end;

{=========================================================================================}
{============================= Evenements Onglet 1 =======================================}
{=========================================================================================}
procedure TFAssistDupTarif.CB_PRIXClick(Sender: TObject);
begin
  inherited;
bSuivant.Enabled:=(CB_PRIX.Checked) or (CB_DATE.Checked);
end;

procedure TFAssistDupTarif.CB_DATEClick(Sender: TObject);
begin
  inherited;
bSuivant.Enabled:=(CB_PRIX.Checked) or (CB_DATE.Checked);
end;

{=========================================================================================}
{============================= Evenements Onglet 2 =======================================}
{=========================================================================================}
procedure TFAssistDupTarif.TYPEMAJPRIXChange(Sender: TObject);
begin
  inherited;
if TYPEMAJPRIX.Value <> '' then
    BEGIN
    VALEURPRIX.Enabled := True;
    VALEURPRIX.Color := clWindow;
    VALEURPRIX.TabStop := True;
    if TYPEMAJPRIX.Value = 'P03' then VALEURPRIX.NumericType:= ntPercentage
                                 else VALEURPRIX.NumericType:= ntGeneral;
    if TYPEMAJPRIX.Value = 'P01' then
        begin
        VALEURPRIX.Decimals := 2;
        VALEURPRIX.Masks.PositiveMask := '#,##0.00';
        end else
        begin
        VALEURPRIX.Decimals := 3;
        VALEURPRIX.Masks.PositiveMask := '#,##0.000';
        end;
    END else
    BEGIN
    VALEURPRIX.Enabled := False;
    VALEURPRIX.Color := clBtnFace;
    VALEURPRIX.TabStop := False;
    END;
end;

procedure TFAssistDupTarif.TYPEMAJREMISEChange(Sender: TObject);
begin
  inherited;
if TYPEMAJREMISE.Value <> '' then
    BEGIN
    VALEURREMISE.Enabled := True;
    VALEURREMISE.Color := clWindow;
    VALEURREMISE.TabStop := True;
    END else
    BEGIN
    VALEURREMISE.Enabled := False;
    VALEURREMISE.Color := clBtnFace;
    VALEURREMISE.TabStop := False;
    END;
end;

procedure TFAssistDupTarif.DATEEFFETExit(Sender: TObject);
var stErr : string;
begin
  inherited;
  if not IsValidDate(DATEEFFET.Text) then
  begin
    stErr:='"'+ DATEEFFET.Text + '" ' + HMsgErr.Mess[0];
    HShowMessage('0;'+PTITRE.Caption+';'+StErr+';W;O;O;O;','','') ;
    DATEEFFET.SetFocus;
  end;
end;

{=========================================================================================}
{============================= Evenements Onglet 3 =======================================}
{=========================================================================================}

procedure TFAssistDupTarif.TYPEMAJPERIODEChange(Sender: TObject);
begin
  inherited;
if (TYPEMAJPERIODE.Value = 'D02') or (TYPEMAJPERIODE.Value = 'D03') then
    BEGIN
    NBREJOUR.Enabled := True;
    NBREJOUR.Color := clWindow;
    NBREJOUR.TabStop := True;
    UpDownJour.Enabled := True;
    CBDATEDEBUT.Enabled := False;
    CBDATEDEBUT.TabStop := False;
    CBDATEFIN.Enabled := False;
    CBDATEFIN.TabStop := False;
    DATEDEBUT.Enabled := False;
    DATEDEBUT.Color := clBtnFace;
    DATEDEBUT.TabStop := False;
    DATEFIN.Enabled := False;
    DATEFIN.Color := clBtnFace;
    DATEFIN.TabStop := False;
    END else if TYPEMAJPERIODE.Value = 'D01' then
    BEGIN
    NBREJOUR.Enabled := False;
    NBREJOUR.Color := clBtnFace;
    NBREJOUR.TabStop := False;
    UpDownJour.Enabled := False;
    CBDATEDEBUT.Enabled := True;
    CBDATEDEBUT.TabStop := True;
    CBDATEFIN.Enabled := True;
    CBDATEFIN.TabStop := True;
    if CBDATEDEBUT.Checked then
        BEGIN
        DATEDEBUT.Enabled := True;
        DATEDEBUT.Color := clWindow;
        DATEDEBUT.TabStop := True;
        END else
        BEGIN
        DATEDEBUT.Enabled := False;
        DATEDEBUT.Color := clBtnFace;
        DATEDEBUT.TabStop := False;
        END;
    if CBDATEFIN.Checked then
        BEGIN
        DATEFIN.Enabled := True;
        DATEFIN.Color := clWindow;
        DATEFIN.TabStop := True;
        END else
        BEGIN
        DATEFIN.Enabled := False;
        DATEFIN.Color := clBtnFace;
        DATEFIN.TabStop := False;
        END;
    END else
    BEGIN
    NBREJOUR.Enabled := False;
    NBREJOUR.Color := clBtnFace;
    NBREJOUR.TabStop := False;
    UpDownJour.Enabled := False;
    CBDATEDEBUT.Enabled := False;
    CBDATEDEBUT.TabStop := False;
    CBDATEFIN.Enabled := False;
    CBDATEFIN.TabStop := False;
    DATEDEBUT.Enabled := False;
    DATEDEBUT.Color := clBtnFace;
    DATEDEBUT.TabStop := False;
    DATEFIN.Enabled := False;
    DATEFIN.Color := clBtnFace;
    DATEFIN.TabStop := False;
    END;
end;

procedure TFAssistDupTarif.CBDATEDEBUTClick(Sender: TObject);
begin
  inherited;
if CBDATEDEBUT.Checked then
    BEGIN
    DATEDEBUT.Enabled := True;
    DATEDEBUT.Color := clWindow;
    DATEDEBUT.TabStop := True;
    END else
    BEGIN
    DATEDEBUT.Enabled := False;
    DATEDEBUT.Color := clBtnFace;
    DATEDEBUT.TabStop := False;
    END;
end;

procedure TFAssistDupTarif.CBDATEFINClick(Sender: TObject);
begin
  inherited;
if CBDATEFIN.Checked then
    BEGIN
    DATEFIN.Enabled := True;
    DATEFIN.Color := clWindow;
    DATEFIN.TabStop := True;
    END else
    BEGIN
    DATEFIN.Enabled := False;
    DATEFIN.Color := clBtnFace;
    DATEFIN.TabStop := False;
    END;
end;

procedure TFAssistDupTarif.DATEDEBUTExit(Sender: TObject);
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

procedure TFAssistDupTarif.DATEFINExit(Sender: TObject);
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

procedure TFAssistDupTarif.TraiterTarif ;
Var ioerr : TIOErr ;
    TOBJnal : TOB ;
    QQ : TQuery ;
    NumEvt,NbEnreg : integer ;
BEGIN
NbEnreg:=TOBTarif.detail.count ;
if NbEnreg <= 0 then exit;

NumEvt:=0 ;
TOBJnal:=TOB.Create('JNALEVENT', Nil, -1) ;
TOBJnal.PutValue('GEV_TYPEEVENT', 'TAR');
TOBJnal.PutValue('GEV_LIBELLE', PTITRE.Caption);
TOBJnal.PutValue('GEV_DATEEVENT', Date);
TOBJnal.PutValue('GEV_UTILISATEUR', V_PGI.User);
ioerr := Transactions (MajTarif, 2);
QQ:=OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT',True,-1,'',true) ;
if Not QQ.EOF then NumEvt:=QQ.Fields[0].AsInteger ;
Ferme(QQ) ;
Inc(NumEvt) ;
TOBJnal.PutValue('GEV_NUMEVENT', NumEvt);
Case ioerr of
        oeOk  : BEGIN
                TOBJnal.PutValue('GEV_ETATEVENT', 'OK');
                ListRecap.Items.Add('');
                ListRecap.Items.Add(IntToStr (NbEnreg) + ' ' + HRecap.Mess[4]);
                TOBJnal.PutValue('GEV_BLOCNOTE', ListRecap.Items.Text);
                END;
    oeUnknown : BEGIN
                MessageAlerte(HRecap.Mess[2]) ;
                ListRecap.Items.Add('');
                ListRecap.Items.Add(HRecap.Mess[2]);
                TOBJnal.PutValue('GEV_ETATEVENT', 'ERR');
                END ;
    oeSaisie  : BEGIN
                MessageAlerte(HRecap.Mess[3]) ;
                ListRecap.Items.Add('');
                ListRecap.Items.Add(HRecap.Mess[3]);
                TOBJnal.PutValue('GEV_ETATEVENT', 'ERR');
                END ;
   END ;
TOBJnal.InsertDB(Nil) ;
TOBJnal.Free ;
END;

procedure TFAssistDupTarif.MajTarif ;
Var i_ind : integer ;
    TOBD : TOB ;
BEGIN
for i_ind := 0 to TOBTarif.detail.count -1 do
    BEGIN
    TOBD := TOBTarif.Detail[i_ind] ;
    MajTarifDetail (TOBD) ;
    END;
TOBTarif.UpdateDB(False) ;
end;

procedure TFAssistDupTarif.MajTarifDetail (TOBD : TOB) ;
Var ModifDateEffet : Boolean ;
BEGIN
ModifDateEffet:= True;
if CB_PRIX.checked then
    BEGIN
    if TYPEMAJPRIX.Value <> '' then
        BEGIN
        ModifPrix (TOBD, ModifDateEffet);
        ModifDateEffet := False;
        END;
    if TYPEMAJREMISE.Value <> '' then
        BEGIN
        ModifRemise (TOBD, ModifDateEffet);
        //ModifDateEffet := False;
        END;
    END;
if CB_DATE.checked then
    BEGIN
    if TYPEMAJPERIODE.Value <> '' then
        BEGIN
        ModifPeriode (TOBD);
        END;
    END;
END;

procedure TFAssistDupTarif.ModifPrix (Var TOBD : TOB; ModifDateEffet : Boolean) ;
Var Valeur, Prix : Double;
    CodeArr, CodeDevise : String ;
BEGIN
if TOBD.GetValue ('GF_REGIMEPRIX')='GLO' then exit;
Valeur := VALEURPRIX.Value;
CodeArr := ARRONDIPRIX.Value;
CodeDevise := TOBD.GetValue ('GF_DEVISE');
if ModifDateEffet then
    BEGIN
    TOBD.PutValue ('GF_PRIXANCIEN', TOBD.GetValue ('GF_PRIXUNITAIRE'));
    TOBD.PutValue ('GF_REMISEANCIEN', TOBD.GetValue ('GF_REMISE'));
    TOBD.PutValue ('GF_DATEEFFET', StrToDate (DATEEFFET.Text));
    END;
if TYPEMAJPRIX.Value = 'P01' then
    BEGIN
    TOBD.PutValue ('GF_PRIXUNITAIRE', Valeur) ;
    END else if TYPEMAJPRIX.Value = 'P02' then
        BEGIN
        Prix := TOBD.GetValue ('GF_PRIXUNITAIRE') * Valeur;
        Prix := ArrondirPrix (CodeArr, Prix);
        TOBD.PutValue ('GF_PRIXUNITAIRE', Prix) ;
        END else if TYPEMAJPRIX.Value = 'P03' then
            BEGIN
            Prix := (TOBD.GetValue ('GF_PRIXUNITAIRE') * (1 + Valeur));
            Prix := ArrondirPrix (CodeArr, Prix);
            TOBD.PutValue ('GF_PRIXUNITAIRE', Prix) ;
            END;
END;

procedure TFAssistDupTarif.ModifRemise (Var TOBD : TOB; ModifDateEffet : Boolean) ;
Var Valeur, Remise : Double;
    Calcremise : string ;
BEGIN
Valeur:= strToFloat(VALEURREMISE.Text);
if ModifDateEffet then
    BEGIN
    TOBD.PutValue ('GF_PRIXANCIEN', TOBD.GetValue ('GF_PRIXUNITAIRE'));
    TOBD.PutValue ('GF_REMISEANCIEN', TOBD.GetValue ('GF_REMISE'));
    TOBD.PutValue ('GF_DATEEFFET', StrToDate (DATEEFFET.Text));
    END;
if TYPEMAJREMISE.Value = 'R01' then
    BEGIN
    TOBD.PutValue ('GF_REMISE', Valeur) ;
    TOBD.PutValue ('GF_CALCULREMISE', FloatToStr (Valeur)) ;
    END else if TYPEMAJREMISE.Value = 'R02' then
        BEGIN
        CalcRemise := TOBD.GetValue ('GF_CALCULREMISE') + '+' + Trim(FloatToStr (Valeur));
        Remise := RemiseResultante (CalcRemise);
        TOBD.PutValue ('GF_REMISE', Remise) ;
        TOBD.PutValue ('GF_CALCULREMISE', CalcRemise) ;
        END
END;

procedure TFAssistDupTarif.ModifPeriode (Var TOBD : TOB) ;
Var DateD, DateF : TDateTime ;
    nbj : integer ;
BEGIN
if TYPEMAJPERIODE.Value = 'D01' then
    BEGIN
    if CBDATEDEBUT.Checked then DateD:=StrToDate(DATEDEBUT.Text) else DateD:=TOBD.GetValue('GF_DATEDEBUT') ;
    if CBDATEFIN.Checked then Datef:=StrToDate(DATEFIN.Text) else DateF:=TOBD.GetValue('GF_DATEFIN') ;
    if DateD <= DateF then
        BEGIN
        TOBD.PutValue ('GF_DATEDEBUT', DateD);
        TOBD.PutValue ('GF_DATEFIN', DateF);
        END;
    END else if TYPEMAJPERIODE.Value = 'D02' then
        BEGIN
        DateD:=TOBD.GetValue('GF_DATEDEBUT') ;
        DateF:=TOBD.GetValue('GF_DATEFIN') ;
        nbj := StrToInt (NBREJOUR.Text) ;
        DateF:=DateF + nbj;
        if DateF > IDate2099 then  DateF := IDate2099 ;
        if DateD <= DateF then
            BEGIN
            TOBD.PutValue ('GF_DATEDEBUT', DateD);
            TOBD.PutValue ('GF_DATEFIN', DateF);
            END;
        END else if TYPEMAJPERIODE.Value = 'D03' then
        BEGIN
        DateD:=TOBD.GetValue('GF_DATEDEBUT') ;
        DateF:=TOBD.GetValue('GF_DATEFIN') ;
        nbj := StrToInt (NBREJOUR.Text) ;
        DateF:=DateF - nbj;
        if DateD <= DateF then
            BEGIN
            TOBD.PutValue ('GF_DATEDEBUT', DateD);
            TOBD.PutValue ('GF_DATEFIN', DateF);
            END;
        END;
END;

procedure TFAssistDupTarif.ListeRecap;
Var st_chaine : string;
BEGIN
ListRecap.Items.Clear;
ListRecap.Items.Add (PTITRE.Caption);
ListRecap.Items.Add ('');
if CB_PRIX.checked then
    BEGIN
    st_chaine := HRecap.Mess[0];
    ListRecap.Items.Add (ExtractLibelle (CB_PRIX.Caption) + st_chaine);
//    ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (TDATEEFFET.Caption) + DATEEFFET.Text);
    if GBPRIX.Visible then
    begin
        ListRecap.Items.Add ( SpaceStr(4) + ExtractLibelle (GBPRIX.Caption));
        ListRecap.Items.Add (SpaceStr(8) + ExtractLibelle (TTYPEMAJPRIX.Caption) + TYPEMAJPRIX.Text);
        if TYPEMAJPRIX.Value <> '' then
        BEGIN
            ListRecap.Items.Add (SpaceStr(8) + ExtractLibelle (TVALEURPRIX.Caption) + VALEURPRIX.Text);
            ListRecap.Items.Add (SpaceStr(8) + ExtractLibelle (TARRONDIPRIX.Caption) + ARRONDIPRIX.Text);
        END;
    end;
    ListRecap.Items.Add ( SpaceStr(4) + ExtractLibelle (GBREMISE.Caption));
    ListRecap.Items.Add (SpaceStr(8) + ExtractLibelle (TTYPEMAJREMISE.Caption) + TYPEMAJREMISE.Text);
    if TYPEMAJREMISE.Value <> '' then
        BEGIN
        ListRecap.Items.Add (SpaceStr(8) + ExtractLibelle (TVALEURREMISE.Caption) + VALEURREMISE.Text);
        END;
    END else
    BEGIN
    st_chaine := HRecap.Mess[1];
    ListRecap.Items.Add (ExtractLibelle (CB_PRIX.Caption) + st_chaine);
    END;
ListRecap.Items.Add ('');
if CB_DATE.checked then
    BEGIN
    st_chaine := HRecap.Mess[0];
    ListRecap.Items.Add (ExtractLibelle (CB_DATE.Caption) + st_chaine);
    ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (TTYPEMAJPERIODE.Caption) + TYPEMAJPERIODE.Text);
    if TYPEMAJPERIODE.Value <> '' then
        if TYPEMAJPERIODE.Value = 'D01' then
            BEGIN
            if CBDATEDEBUT.Checked then
                BEGIN
                st_chaine := HRecap.Mess[0];
                ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (CBDATEDEBUT.Caption) + St_Chaine);
                ListRecap.Items.Add (SpaceStr(8) + ExtractLibelle (TDATEDEBUT.Caption) + DATEDEBUT.Text);
                END else
                BEGIN
                st_chaine := HRecap.Mess[1];
                ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (CBDATEDEBUT.Caption) + St_Chaine);
                END;
            if CBDATEFIN.Checked then
                BEGIN
                st_chaine := HRecap.Mess[0];
                ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (CBDATEFIN.Caption) + St_Chaine);
                ListRecap.Items.Add (SpaceStr(8) + ExtractLibelle (TDATEFIN.Caption) + DATEFIN.Text);
                END else
                BEGIN
                st_chaine := HRecap.Mess[1];
                ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (CBDATEFIN.Caption) + St_Chaine);
                END;
            END else
            BEGIN
            ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (TNBREJOUR.Caption) + NBREJOUR.Text);
            END;
    END else
    BEGIN
    st_chaine := HRecap.Mess[1];
    ListRecap.Items.Add (ExtractLibelle (CB_DATE.Caption) + st_chaine);
    END;
if CBBasculeTTC.Checked = True then
    begin
    ListRecap.Items.Add ('');
    ListRecap.Items.Add (HRecap.Mess [5]);
    end;
if CBBasculeHT.Checked = True then
    begin
    ListRecap.Items.Add ('');
    ListRecap.Items.Add (HRecap.Mess [6]);
    end;
ListRecap.Items.Add ('');
END;

end.
