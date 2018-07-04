unit AssistMajTarifArt;

interface

uses
     Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
     assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls,
{$IFDEF EAGLCLIENT}
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
{$IFDEF BTP}
     UtilArticle,
{$ENDIF}
     ParamSoc, HPanel, UTOB, HEnt1, TarifUtil, UtilGC, FactUtil, Mask;

procedure Assist_MajTarifArt (TOBArticle : TOB);

type       
  TFAssistTarifArt = class(TFAssist)
    TabSheet1: TTabSheet;
    PTITRE: THPanel;
    TabSheet2: TTabSheet;
    ListRecap: TListBox;
    PanelFin: TPanel;
    TTextFin1: THLabel;
    TTextFin2: THLabel;
    TRecap: THLabel;
    TabSheet3: TTabSheet;
    GBParamTTC: TGroupBox;
    TARRONDITTC: THLabel;
    TTYPEMAJPRIXTTC: THLabel;
    TVALEURPRIXTTC: THLabel;
    ARRONDITTC: THValComboBox;
    TYPEMAJPRIXTTC: THValComboBox;
    VALEURPRIXTTC: THNumEdit;
    GBParamHT: TGroupBox;
    TARRONDIHT: THLabel;
    TTYPEMAJPRIXHT: THLabel;
    TVALEURPRIXHT: THLabel;
    ARRONDIHT: THValComboBox;
    TYPEMAJPRIXHT: THValComboBox;
    VALEURPRIXHT: THNumEdit;
    GBParamBase: TGroupBox;
    TARRONDIBASE: THLabel;
    TTYPEMAJPRIXBASE: THLabel;
    TVALEURPRIXBASE: THLabel;
    ARRONDIBASE: THValComboBox;
    TYPEMAJPRIXBASE: THValComboBox;
    VALEURPRIXBASE: THNumEdit;
    TINTRO: THLabel;
    GroupBox2: TGroupBox;
    Label1: TLabel;
    CALCPASAUTOPV: TCheckBox;
    HT: TRadioButton;
    TTC: TRadioButton;
    CALCAUTOPV: TCheckBox;
    ECART: THNumEdit;
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure TYPEMAJPRIXHTChange(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure TYPEMAJPRIXTTCChange(Sender: TObject);
    procedure TYPEMAJPRIXBASEChange(Sender: TObject);

  private
    { Déclarations privées }
    TOBArt : TOB ;
    bParamHT, bParamTTC, bParamPxBase : boolean;
    procedure CreerListeRecap;
    procedure MajTarif ;
    procedure ModifPrix (TOBD : TOB ; NomChamp,TypeMAJ,CodeArr : string ; Valeur,PrixRef : double) ;
    procedure TraiterTarif ;
    procedure VerifParam;
    procedure RecherchePxAffecte(TOBD : TOB ; CodeBase : string ; PrixRef : double);
  public
    { Déclarations publiques }
  end;

var
  FAssistTarifArt: TFAssistTarifArt;

implementation

{$R *.DFM}

procedure Assist_MajTarifArt (TOBArticle : TOB);
var
   Fo_Assist : TFAssistTarifArt;
Begin
     Fo_Assist := TFAssistTarifArt.Create (Application);
     Fo_Assist.TOBArt:= TOBArticle;
     Try
         Fo_Assist.ShowModal;
     Finally
         Fo_Assist.free;
     End;
end;

{=========================================================================================}
{============================= Evenements de la form =====================================}
{=========================================================================================}

procedure TFAssistTarifArt.FormShow(Sender: TObject);
var DecPrix : string;
    i : integer;
begin
  inherited;
bAnnuler.Visible := True;
bFin.Visible := True;
bFin.Enabled := False;
  DecPrix := '';
  for i := 1 to GetParamSoc('SO_DECPRIX') do DecPrix := DecPrix + '0';
  VALEURPRIXTTC.Masks.PositiveMask := '#,##0.' + DecPrix;
  VALEURPRIXHT.Masks.PositiveMask := '#,##0.' + DecPrix;
  VALEURPRIXBASE.Masks.PositiveMask := '#,##0.' + DecPrix;
end;

procedure TFAssistTarifArt.bSuivantClick(Sender: TObject);
begin
  inherited;
if (bSuivant.Enabled) then bFin.Enabled := False else bFin.Enabled := True;
CreerListeRecap;
end;

procedure TFAssistTarifArt.CreerListeRecap;
var StSpace,stChaine : string;
    stTypePrix : string;
begin
StSpace:='   ';
stChaine:=' non';
VerifParam;
ListRecap.Clear;
if not bParamPxBase then ListRecap.Items.Add(ExtractLibelle(GBParamBase.Caption)+stChaine)
else begin
     ListRecap.Items.Add(ExtractLibelle(GBParamBase.Caption));
     ListRecap.Items.Add(StSpace + ExtractLibelle(TTYPEMAJPRIXBASE.Caption) + TYPEMAJPRIXBASE.Text);
     ListRecap.Items.Add(StSpace + ExtractLibelle(TVALEURPRIXBASE.Caption) + VALEURPRIXBASE.Text);
     ListRecap.Items.Add(StSpace + ExtractLibelle(TARRONDIBASE.Caption) + ARRONDIBASE.Text);
     end;
ListRecap.Items.Add('');
if not bParamHT then ListRecap.Items.Add(ExtractLibelle(GBParamHT.Caption)+stChaine)
else begin
     ListRecap.Items.Add(ExtractLibelle(GBParamHT.Caption));
     ListRecap.Items.Add(StSpace + ExtractLibelle(TTYPEMAJPRIXHT.Caption) + TYPEMAJPRIXHT.Text);
     ListRecap.Items.Add(StSpace + ExtractLibelle(TVALEURPRIXHT.Caption) + VALEURPRIXHT.Text);
     ListRecap.Items.Add(StSpace + ExtractLibelle(TARRONDIHT.Caption) + ARRONDIHT.Text);
    end;
ListRecap.Items.Add('');
if not bParamTTC then ListRecap.Items.Add(ExtractLibelle(GBParamTTC.Caption)+stChaine)
else begin
     ListRecap.Items.Add(ExtractLibelle(GBParamTTC.Caption));
     ListRecap.Items.Add(StSpace + ExtractLibelle(TTYPEMAJPRIXTTC.Caption) + TYPEMAJPRIXTTC.Text);
     ListRecap.Items.Add(StSpace + ExtractLibelle(TVALEURPRIXTTC.Caption) + VALEURPRIXTTC.Text);
     ListRecap.Items.Add(StSpace + ExtractLibelle(TARRONDITTC.Caption) + ARRONDITTC.Text);
     end;
ListRecap.Items.Add('');

if HT.Checked = true then stTypePrix := 'HT'
else stTypePrix := 'TTC';

if CALCAUTOPV.Checked then
    begin
    ListRecap.Items.Add ('Mise A jour des Prix de vente ' + stTypePrix + ' ' +
                         ExtractLibelle (CALCAUTOPV.Caption));
    ListRecap.Items.Add('');
    end;
if (CALCPASAUTOPV.Checked) and (StrToFloat (ECART.Text) > 0.0) then
    begin
    ListRecap.Items.Add ('Mise A jour des Prix de vente ' + stTypePrix);
    ListRecap.Items.Add (ExtractLibelle (CALCPASAUTOPV.Caption) + ' à ' + ECART.Text + ' %');
    ListRecap.Items.add ('Entre le prix de la fiche et le prix recalculé');
    end;
if (not CALCAUTOPV.Checked) and
   ((not CALCPASAUTOPV.Checked) or (StrToFloat (ECART.Text) <= 0.0)) then
    begin
    ListRecap.Items.Add ('Pas de mise a jour des Prix de vente');
    end;
end;

procedure TFAssistTarifArt.bPrecedentClick(Sender: TObject);
begin
  inherited;
if (bSuivant.Enabled) then bFin.Enabled := False else bFin.Enabled := True;
end;

procedure TFAssistTarifArt.bFinClick(Sender: TObject);
begin
  inherited;
TraiterTarif ;
Close ;
end;

procedure TFAssistTarifArt.TYPEMAJPRIXBASEChange(Sender: TObject);
begin
  inherited;
if TYPEMAJPRIXBASE.Value <> '' then
    Begin
    VALEURPRIXBASE.Enabled := True;
    VALEURPRIXBASE.Color := clWindow;
    VALEURPRIXBASE.TabStop := True;
    if TYPEMAJPRIXBASE.Value = 'P03' then VALEURPRIXBASE.NumericType:= ntPercentage
                                 else VALEURPRIXBASE.NumericType:= ntGeneral;
    end else
    Begin
    VALEURPRIXBASE.Enabled := False;
    VALEURPRIXBASE.Color := clBtnFace;
    VALEURPRIXBASE.TabStop := False;
    end;
end;

procedure TFAssistTarifArt.TYPEMAJPRIXHTChange(Sender: TObject);
begin
  inherited;
if TYPEMAJPRIXHT.Value <> '' then
    Begin
    VALEURPRIXHT.Enabled := True;
    VALEURPRIXHT.Color := clWindow;
    VALEURPRIXHT.TabStop := True;
    if TYPEMAJPRIXHT.Value = 'P03' then VALEURPRIXHT.NumericType:= ntPercentage
                                 else VALEURPRIXHT.NumericType:= ntGeneral;
    end else
    Begin
    VALEURPRIXHT.Enabled := False;
    VALEURPRIXHT.Color := clBtnFace;
    VALEURPRIXHT.TabStop := False;
    end;
end;

procedure TFAssistTarifArt.TYPEMAJPRIXTTCChange(Sender: TObject);
begin
  inherited;
if TYPEMAJPRIXTTC.Value <> '' then
    Begin
    VALEURPRIXTTC.Enabled := True;
    VALEURPRIXTTC.Color := clWindow;
    VALEURPRIXTTC.TabStop := True;
    if TYPEMAJPRIXTTC.Value = 'P03' then VALEURPRIXTTC.NumericType:= ntPercentage
                                 else VALEURPRIXTTC.NumericType:= ntGeneral;
    end else
    Begin
    VALEURPRIXTTC.Enabled := False;
    VALEURPRIXTTC.Color := clBtnFace;
    VALEURPRIXTTC.TabStop := False;
    end;
end;


{=========================================================================================}
{================================== Traitement ===========================================}
{=========================================================================================}

procedure TFAssistTarifArt.MajTarif ;
Var i_ind : integer ;
    TOBD : TOB ;
    stTypePrix : string;
    dPrix : double;
    stSavAutoTTC, stSavAutoHT : string;
begin

  if not bParamHT           and
     not bParamTTC          and
     not bParamPxBase       and
     not CALCAUTOPV.Checked and
     not CALCPASAUTOPV.Checked then exit;

  for i_ind := 0 to TOBArt.detail.count -1 do
  begin
    TOBD := TOBArt.Detail[i_ind] ;
    //
    //Modification du prix de base
    if bParamPxBase then
    begin
      ModifPrix (TOBD,'GA_PAHT',TYPEMAJPRIXBASE.Value,ARRONDIBASE.Value,VALEURPRIXBASE.Value,-1);
      RecalculPrPV(TOBD,Nil);
    end;

    if bParamHT then
    begin
      ModifPrix (TOBD,'GA_PVHT',TYPEMAJPRIXHT.Value,ARRONDIHT.Value,VALEURPRIXHT.Value,-1);
      RecalculCoefMarg (TOBD);
      RecalculPvTTC (TOBD,Nil);
    end;

    if bParamTTC then
    begin
      ModifPrix (TOBD,'GA_PVTTC',TYPEMAJPRIXTTC.Value,ARRONDITTC.Value,VALEURPRIXTTC.Value,-1);
      recalculCoef(TOBD);
    end;

    //gestion des Check HT et TTC.
    If (Not bParamPxBase) And (Not BParamHT) And (Not BParamTTC) then RecalculPrPV(TOBD,Nil);

    {*
    if bParamHT and ((TOBD.GetValue('GA_CALCPRIXHT')='AUC')or(TOBD.GetValue('GA_CALCPRIXHT')='')) then
    begin
       ModifPrix (TOBD,'GA_PVHT',TYPEMAJPRIXHT.Value,ARRONDIHT.Value,VALEURPRIXHT.Value,-1);
       RecherchePxAffecte(TOBD,'HT',TOBD.GetValue('GA_PVHT'));
    end
    else
      if bParamHT and (TOBD.GetValue('GA_CALCPRIXHT')<>'AUC')and(TOBD.GetValue('GA_CALCPRIXHT')<>'') then
      begin
       ModifPrix (TOBD,'GA_PVHT',TYPEMAJPRIXHT.Value,ARRONDIHT.Value,VALEURPRIXHT.Value,-1);
       recalculCoefMarg (TOBD);
      end;
      if bParamTTC and ((TOBD.GetValue('GA_CALCPRIXTTC')='AUC') or (TOBD.GetValue('GA_CALCPRIXTTC')='')) then
         ModifPrix (TOBD,'GA_PVTTC',TYPEMAJPRIXTTC.Value,ARRONDITTC.Value,VALEURPRIXTTC.Value,-1);
    //
    //??? A quoi ça peut bien servir ce truc !!!!!!!!!!!
    //
    //mcd 03/05/01 on profite de cette option, pour mettre les "CON" correct
    if (CALCAUTOPV.Checked = True) or (CALCPASAUTOPV.Checked = True) then
    begin
      if HT.Checked = true then
        stTypePrix := 'HT'
      else
        stTypePrix := 'TTC';
      //
      stSavAutoTTC := TOBD.GetValue ('GA_CALCAUTOTTC');
      stSavAutoHT := TOBD.GetValue ('GA_CALCAUTOHT');
      //
      if (CALCPASAUTOPV.Checked = true) and (StrToFloat (ECART.Text) > 0.0) and
         (TOBD.GetValue ('GA_CALCPRIX' + stTypePrix) <> 'AUC') and
         (TOBD.GetVAlue ('GA_CALCPRIX' + stTypePrix) <> '') and
         (TOBD.GetValue ('GA_COEFCALC' + stTypePrix) <> 0.0) and
         (TOBD.GetValue ('GA_CALCAUTO' + stTypePrix) = '-') then
      begin
        TOBD.PutValue ('GA_CALCAUTO' + stTypePrix, 'X');
        dPrix := TOBD.GetValue ('GA_PV' + stTypePrix);
        CalculePrixArticle (TOBD);
        if TOBD.GetValue ('GA_PV' + stTypePrix) <> 0.0 then
        begin
          if (TOBD.GetValue ('GA_PV' + stTypePrix) - dPrix)*100.0/TOBD.GetValue ('GA_PV' + stTypePrix) < StrToFloat (ECART.Text) then
          begin
            TOBD.PutValue ('GA_PV' + stTypePrix, dPrix);
          end;
        end
        else
          TOBD.PutValue ('GA_PV' + stTypePrix, dPrix);
      end
      else if CALCAUTOPV.Checked = true then
      begin
        if stTypePrix = 'HT' then
          TOBD.PutValue ('GA_CALCAUTOTTC', '-')
        else
          TOBD.PutValue ('GA_CALCAUTOHT', '-');
        //
        if (TOBD.GetValue ('GA_CALCAUTO' + stTypePrix) = 'X') then
        begin
          dPrix := TOBD.GetValue ('GA_PV' + stTypePrix);
          CalculePrixArticle (TOBD);
        end;
      end;
      TOBD.PutValue ('GA_CALCAUTOTTC', stSavAutoTTC);
      TOBD.PutValue ('GA_CALCAUTOHT', stSavAutoHT);
    end;
  *}
  end;
    
  TOBArt.UpdateDB(False) ;

end;

procedure TFAssistTarifArt.ModifPrix (TOBD : TOB ; NomChamp,TypeMAJ,CodeArr : string ; Valeur,PrixRef : double);
Var Prix : Double;
begin

  Prix := 0;

  if PrixRef=-1 then PrixRef := TOBD.GetValue(NomChamp);

  if TypeMAJ = 'P01' then
    Prix:=Valeur
  else if TypeMAJ = 'P02' then
    Prix := PrixRef * Valeur
  else if TypeMAJ = 'P03' then
    Prix := (PrixRef * (1 + Valeur));

  Prix := ArrondirPrix(CodeArr, Prix);
  TOBD.PutValue (NomChamp, Prix) ;

end;

procedure TFAssistTarifArt.RecherchePxAffecte(TOBD : TOB ; CodeBase : string ; PrixRef : double);
begin

  if (TOBD.GetValue('GA_CALCPRIXHT')=CodeBase) and (TOBD.GetValue('GA_CALCAUTOHT')='X')then
  begin
    ModifPrix(TOBD,'GA_PVHT','P02',TOBD.GetValue('GA_ARRONDIPRIX'),TOBD.GetValue('GA_COEFCALCHT'),PrixRef);
    RecherchePxAffecte(TOBD,'HT',TOBD.GetValue('GA_PVHT'));
  end;

  if (TOBD.GetValue('GA_CALCPRIXTTC')=CodeBase) and (TOBD.GetValue('GA_CALCAUTOTTC')='X') then
    ModifPrix(TOBD,'GA_PVTTC','P02',TOBD.GetValue('GA_ARRONDIPRIX'),TOBD.GetValue('GA_COEFCALCTTC'),PrixRef);

end;

procedure TFAssistTarifArt.TraiterTarif ;
var ioerr : TIOErr ;
    TOBJnal : TOB ;
    QQ : TQuery ;
    NumEvt,NbEnreg : integer ;
begin
NbEnreg:=TOBArt.detail.count ;
if NbEnreg <= 0 then exit;

NumEvt:=0 ;
TOBJnal:=TOB.Create('JNALEVENT', Nil, -1) ;
TOBJnal.PutValue('GEV_TYPEEVENT', 'TAR');
TOBJnal.PutValue('GEV_LIBELLE', 'Bases tarifaires articles');
TOBJnal.PutValue('GEV_DATEEVENT', Date);
TOBJnal.PutValue('GEV_UTILISATEUR', V_PGI.User);
VerifParam;
ioerr := Transactions (MajTarif, 2);
QQ:=OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT',True,-1,'',true) ;
if Not QQ.EOF then NumEvt:=QQ.Fields[0].AsInteger+1 ;
Ferme(QQ) ;
Inc(NumEvt) ;
TOBJnal.PutValue('GEV_NUMEVENT', NumEvt);
Case ioerr of
        oeOk  : Begin
                TOBJnal.PutValue('GEV_ETATEVENT', 'OK');
                ListRecap.Items.Add('');
                ListRecap.Items.Add(IntToStr (NbEnreg) + ' ' + Msg.Mess[4]);
                TOBJnal.PutValue('GEV_BLOCNOTE', ListRecap.Items.Text);
                end;
    oeUnknown : Begin
                MessageAlerte(Msg.Mess[2]) ;
                ListRecap.Items.Add('');
                ListRecap.Items.Add(Msg.Mess[2]);
                TOBJnal.PutValue('GEV_ETATEVENT', 'ERR');
                end ;
    oeSaisie  : Begin
                MessageAlerte(Msg.Mess[3]) ;
                ListRecap.Items.Add('');
                ListRecap.Items.Add(Msg.Mess[3]);
                TOBJnal.PutValue('GEV_ETATEVENT', 'ERR');
                end ;
   end ;
TOBJnal.InsertDB(Nil) ;
TOBJnal.Free ;
end;

procedure TFAssistTarifArt.VerifParam;
begin
bParamHT:=(TYPEMAJPRIXHT.Value <> '') and (Valeur(VALEURPRIXHT.Text)<>0);
bParamTTC:=(TYPEMAJPRIXTTC.Value <> '') and (Valeur(VALEURPRIXTTC.Text)<>0);
bParamPxBase:=(TYPEMAJPRIXBASE.Value <> '') and (Valeur(VALEURPRIXBASE.Text)<>0);
end;

end.
