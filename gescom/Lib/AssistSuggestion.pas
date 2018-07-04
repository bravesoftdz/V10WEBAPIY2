unit AssistSuggestion;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, StdCtrls, ComCtrls, HRichEdt, HRichOLE, Hctrls, Spin, Mask,
  ExtCtrls, HSysMenu, hmsgbox,
{$IFDEF EAGLCLIENT}
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  HTB97,
{$IFDEF EAGLCLIENT}
	mainEagl,
{$ELSE}
  fe_main,
{$ENDIF}
  Hent1,AglInitGC, UTOB, TarifUtil, CalculSuggestion, ResulSuggestion,
  Grids, M3FP, HPanel, TntComCtrls, TntStdCtrls, TntExtCtrls ;


type
  TFAssistReappro = class(TFAssist)
    TInitiale: TTabSheet;
    PENTETE: TPanel;
    RadioGroup1: TRadioGroup;
    TcouvSimple: TRadioButton;
    TbesoinP1: TTabSheet;
    TbesoinP2: TTabSheet;
    GB_Histo: TGroupBox;
    RB_Long: TRadioButton;
    RB_Court: TRadioButton;
    SE_Histo: TSpinEdit;
    Label2: TLabel;
    TbesoinP3: TTabSheet;
    TText3: THLabel;
    TGA_FAMILLENIV1: THLabel;
    HLabel1: THLabel;
    GA_FAMILLENIV1: THValComboBox;
    HLabel4: THLabel;
    GA_FAMILLENIV1_: THValComboBox;
    TGA_FAMILLENIV2: THLabel;
    HLabel2: THLabel;
    GA_FAMILLENIV2: THValComboBox;
    HLabel5: THLabel;
    GA_FAMILLENIV2_: THValComboBox;
    TGA_FAMILLENIV3: THLabel;
    HLabel3: THLabel;
    GA_FAMILLENIV3: THValComboBox;
    HLabel6: THLabel;
    GA_FAMILLENIV3_: THValComboBox;
    TbesoinP4: TTabSheet;
    TText2: THLabel;
    TSR_ARTICLE: THLabel;
    GZZ_ARTICLE: THCritMaskEdit;
    TSR_ARTICLE_: THLabel;
    GZZ_ARTICLE_: THCritMaskEdit;
    TSR_DEPOT: THLabel;
    GZZ_DEPOT: THValComboBox;
    TSR_DEPOT_: THLabel;
    GZZ_DEPOT_: THValComboBox;
    TbesoinP5: TTabSheet;
    TText6: THLabel;
    PRIORITE: THValComboBox;
    TPRIORITE1: THLabel;
    PRIORITE1: THValComboBox;
    TPRIORITE2: THLabel;
    PRIORITE2: THValComboBox;
    TPRIORITE3: THLabel;
    PRIORITE3: THValComboBox;
    TPRIORITE4: THLabel;
    PRIORITE4: THValComboBox;
    TRecap: TTabSheet;
    PanelFin: TPanel;
    TTextFin1: THLabel;
    TTextFin2: THLabel;
    HLabel7: THLabel;
    Recap: THRichEditOLE;
    TText4: THLabel;
    TDATEDEB: THLabel;
    DATEDEB: THCritMaskEdit;
    TDATEFIN: THLabel;
    DATEDEB_: THCritMaskEdit;
    CSelDocs: TCheckBox;
    GestHisto: TCheckBox;
    TBesoinStock: TRadioButton;
    Panel1: TPanel;
    TbesoinP0: TTabSheet;
    RGBase: TRadioGroup;
    HLabel8: THLabel;
    RGValeur: TRadioGroup;
    CBFermes: TCheckBox;
    CWithStocks: TCheckBox;
    procedure bSuivantClick(Sender: TObject);
    procedure RB_CourtClick(Sender: TObject);
    procedure RB_LongClick(Sender: TObject);
    procedure DATEDEBElipsisClick(Sender: TObject);
    procedure DATEDEB_ElipsisClick(Sender: TObject);
    procedure GZZ_ARTICLE_ElipsisClick(Sender: TObject);
    procedure GZZ_ARTICLEElipsisClick(Sender: TObject);
    procedure PRIORITE1Change(Sender: TObject);
    procedure PRIORITE2Change(Sender: TObject);
    procedure PRIORITE3Change(Sender: TObject);
    procedure PRIORITE4Change(Sender: TObject);
    procedure GA_FAMILLENIV1Change(Sender: TObject);
    procedure GA_FAMILLENIV2Change(Sender: TObject);
    procedure GA_FAMILLENIV3Change(Sender: TObject);
    procedure GZZ_DEPOTChange(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure TcouvSimpleClick(Sender: TObject);
    procedure TBesoinStockClick(Sender: TObject);
    procedure GestHistoClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure RGValeurClick(Sender: TObject);
  private

    ART : THCritMaskEdit;
    NbrPage : integer;
    procedure InitZones (Sender : Tobject);
    function SelectionArticle(ART: THCritMaskEdit): string;
    function SelectionDate(Parent: TComponent): string;
    procedure AjustePageControl;
    procedure DisplayStep ;

    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

Procedure Assist_Suggestion ;
procedure LanceAssistantSuggestion;

var
  FAssistReappro: TFAssistReappro;

implementation

{$R *.DFM}

procedure LanceAssistantSuggestion;
var Fassist: TFAssistReappro;
begin
	AGLLanceFiche('BTP','BTVIDEINSIDE','','','') ;
  //
	Fassist := TFAssistReappro.create (Application);
  TRY
  	Fassist.ShowModal;
  FINALLY
  	Fassist.free;
  END;
end;


Procedure Assist_Suggestion ;
var Fassist : TFAssistReappro;
    ErrorMsg : string;
Begin
  Fassist := TFAssistReappro.Create (Application);
  if not (Presence ('PARPIECE', 'GPP_NATUREPIECEG', 'REA')) then
    begin
    ErrorMsg := 'Opération impossible, nature de pièce inexistante';
    Fassist.Msg.Execute(0, ErrorMsg, '');
    Fassist.Free;
    Exit;
    end;
  Try
     Fassist.ShowModal;
  Finally
     Fassist.free;
  End;
end;

procedure TFAssistReappro.InitZones (Sender : Tobject);
var Q : Tquery;
begin
  GZZ_DEPOT.ReLoad;
  GZZ_DEPOT_.ReLoad;
  GZZ_DEPOT.ItemIndex := 0;
//  GZZ_DEPOT_.ItemIndex := GZZ_DEPOT_.Items.Count-1;
  GZZ_DEPOT_.ItemIndex := 0;
  //
  PRIORITE.Values.Add('');
  PRIORITE.Values.Add('APP');
  PRIORITE.Values.Add('FPR');
{$IFNDEF BTP}
  PRIORITE.Values.Add('COT');
{$ENDIF}
  PRIORITE.Values.Add('MPT');
  PRIORITE.Items.Text := '<<Aucun>>';
  PRIORITE.Items.Add('Délai d''approvisionnement');
  PRIORITE.Items.Add('Fournisseur principal');
{$IFNDEF BTP}
  PRIORITE.Items.Add('Côte du fournisseur');
{$ELSE}
  PRIORITE4.Visible := false;
  TPRIORITE4.Visible := false;
{$ENDIF}
  PRIORITE.Items.Add('Meilleur prix tarif');
  //
  PRIORITE1.Values.Text := PRIORITE.Values.Text;
  PRIORITE2.Values.Text := PRIORITE.Values.Text;
  PRIORITE3.Values.Text := PRIORITE.Values.Text;
  PRIORITE4.Values.Text := PRIORITE.Values.Text;
  PRIORITE1.Items.Text := PRIORITE.Items.Text;
  PRIORITE2.Items.Text := PRIORITE.Items.Text;
  PRIORITE3.Items.Text := PRIORITE.Items.Text;
  PRIORITE4.Items.Text := PRIORITE.Items.Text;
  PRIORITE1.Items.Delete(0);
  PRIORITE1.Values.Delete(0);
  PRIORITE1.ItemIndex := 0;
  PRIORITE1Change(Sender);
  //
  P.ActivePage := Tinitiale;
  //
  Q := OpenSQL('Select CC_LIBELLE from ChoixCod where CC_TYPE="GLF" and CC_CODE="LF1"', True,-1,'',true);
  if Not Q.EOF then TGA_FAMILLENIV1.Caption := Q.Fields[0].AsString;
  Ferme(Q);
  Q := OpenSQL('Select CC_LIBELLE from ChoixCod where CC_TYPE="GLF" and CC_CODE="LF2"', True,-1,'',true);
  if Not Q.EOF then TGA_FAMILLENIV2.Caption := Q.Fields[0].AsString;
  Ferme(Q);
  Q := OpenSQL('Select CC_LIBELLE from ChoixCod where CC_TYPE="GLF" and CC_CODE="LF3"', True,-1,'',true);
  if Not Q.EOF then TGA_FAMILLENIV3.Caption := Q.Fields[0].AsString;
  Ferme(Q);
  //
  PRIORITE3.Enabled := False;
  PRIORITE4.Enabled := False;
  //
  GA_FAMILLENIV1_.Enabled := False;
  GA_FAMILLENIV2_.Enabled := False;
  GA_FAMILLENIV3_.Enabled := False;
  Se_histo.enabled := false;
  CSelDocs.checked := false;
  gestHisto.enabled := false;
end;

procedure TFAssistReappro.FormShow(Sender: TObject);
begin
inherited;
Recap.Lines.Clear;
InitZones(sender);
RGBase.ItemIndex := 0;
RGBase.Visible := False;
if P.activePageindex = NbrPage-1 then bSuivant.Enabled := false
                                 else bSuivant.Enabled := true;

if (bSuivant.Enabled) then bFin.Enabled := False
                      else bFin.Enabled := True;
DisplayStep;
AjustePageControl;
end;

procedure TFAssistReappro.bSuivantClick(Sender: TObject);
begin
  inherited;
  if P.activePageindex = NbrPage-1 then bSuivant.Enabled := false
                                   else bSuivant.Enabled := true;

  if (bSuivant.Enabled) then bFin.Enabled := False
                        else bFin.Enabled := True;

  if P.activePageindex = NbrPage-1 then
  begin

    if TcouvSimple.Checked  then
      begin
      Recap.lines.text := 'Couverture des besoins de livraison';
      if CSelDocs.checked then Recap.lines.add('Avec sélection de documents')
                          else Recap.lines.add('sur tous les documents')
      end else
      begin
      Recap.lines.text := 'Calcul des besoins de stock';
      end;

    if GestHisto.checked then
       begin
       Recap.Lines.add ('avec prise en compte de l''historique des ventes');
       if RB_Long.Checked  then Recap.lines.add('sur les 2 dernières années')
                           else Recap.lines.add ( 'sur les '+IntToStr(Se_histo.value)+' derniers mois précédent la date de fin de période');
       end;

    if RGValeur.ItemIndex = 0 then
       Recap.Lines.add ('Quantités réapprovisionnées au besoin réel calculé')
    else
       Recap.Lines.add ('Quantités réapprovisionnées au stock maximum');

    if RGBase.ItemIndex = 0 then
       Recap.Lines.add ('Calcul basé sur le stock physique')
    else
       Recap.Lines.add ('Calcul basé sur le stock net');

    if (GestHisto.checked) or (TcouvSimple.checked) then
       begin
       if DATEDEB.Text = '' then Recap.Lines.Add('Début de période : ' {+ V_PGI.})
                            else Recap.Lines.Add('Début de période : ' + DATEDEB.Text);

       if DATEDEB_.Text = '' then Recap.Lines.Add('Fin de période : ' {+ V_PGI.})
                             else Recap.Lines.Add('Fin de période : ' + DATEDEB_.Text);
       end;

    if (GZZ_ARTICLE.Text = '') and (GZZ_ARTICLE_.Text = '') then Recap.Lines.add ('Pas de selection sur les articles')
    else if (GZZ_ARTICLE.Text <> '') and (GZZ_ARTICLE_.Text <> '') then Recap.Lines.add ('Articles compris entre   ' + GZZ_ARTICLE.Text + '   et   ' + GZZ_ARTICLE_.Text)
    else if GZZ_ARTICLE.Text <> '' then Recap.Lines.add ('Articles à partir du   ' + GZZ_ARTICLE.Text + '   jusqu''à la fin')
    else Recap.Lines.add ('Articles à partir du début jusqu''à   ' + GZZ_ARTICLE_.Text);

    if CBFermes.State = cbChecked then Recap.Lines.add ('Uniquement les articles fermés')
    else if CBFermes.State = cbUnChecked then Recap.Lines.add ('Uniquement les articles non fermés')
    else if CBFermes.State = cbGrayed then Recap.Lines.add ('Tous les articles non supprimés au ' + DATEDEB.Text);

    if (GZZ_DEPOT.Value = '') and (GZZ_DEPOT_.Value = '') then Recap.Lines.Add('Pas de selection sur les dépôts')
    else if (GZZ_DEPOT.Value <> '') and (GZZ_DEPOT_.Value <> '') then Recap.Lines.Add('Dépôts compris entre   ' + GZZ_DEPOT.Text + '   et   ' + GZZ_DEPOT_.Text)
    else if GZZ_DEPOT.Value <> '' then Recap.Lines.Add('Dépôts à partir de   ' + GZZ_DEPOT.Text + '   jusqu''à la fin')
    else Recap.Lines.Add('Dépôts à partir du début jusqu''à   ' + GZZ_DEPOT_.Text);

    if GA_FAMILLENIV1.Value = '' then
    Recap.Lines.Add('Pas de selection sur la famille niveau 1')
    else if (GA_FAMILLENIV1.Value <> '') and (GA_FAMILLENIV1_.Value <> '') then
    Recap.Lines.Add('Selection sur la famille niveau 1 : de ' + GA_FAMILLENIV1.Text + ' à ' + GA_FAMILLENIV1_.Text)
    else if GA_FAMILLENIV1.Value <> '' then
    Recap.Lines.Add('Selection sur la famille niveau 1 : ' + GA_FAMILLENIV1.Text);

    if GA_FAMILLENIV2.Value = '' then
    Recap.Lines.Add('Pas de selection sur la famille niveau 2')
    else if (GA_FAMILLENIV2.Value <> '') and (GA_FAMILLENIV2_.Value <> '') then
    Recap.Lines.Add('Selection sur la famille niveau 2 : de ' + GA_FAMILLENIV2.Text + ' à ' + GA_FAMILLENIV2_.Text)
    else if GA_FAMILLENIV2.Value <> '' then
    Recap.Lines.Add('Selection sur la famille niveau 2 : ' + GA_FAMILLENIV2.Text);

    if GA_FAMILLENIV3.Value = '' then
    Recap.Lines.Add('Pas de selection sur la famille niveau 3')
    else if (GA_FAMILLENIV3.Value <> '') and (GA_FAMILLENIV3_.Value <> '') then
    Recap.Lines.Add('Selection sur la famille niveau 3 : de ' + GA_FAMILLENIV3.Text + ' à ' + GA_FAMILLENIV3_.Text)
    else if GA_FAMILLENIV3.Value <> '' then
    Recap.Lines.Add('Selection sur la famille niveau 3 : ' + GA_FAMILLENIV3.Text);


    Recap.Lines.Add('Choix du fournisseur dans l''ordre :');
    Recap.Lines.Add('              ' + PRIORITE1.Text);
    Recap.Lines.Add('              ' + PRIORITE2.Text);
    Recap.Lines.Add('              ' + PRIORITE3.Text);
    Recap.Lines.Add('              ' + PRIORITE4.Text);
  end;
  DisplayStep;
end;


procedure TFAssistReappro.bFinClick(Sender: TObject);
var Params : string;
begin
  inherited;
Params := '';
if strtodate(dateDEB_.text) < strtodate(datedeb.text) then
  begin
  PgiBox ('Incohérence dans les dates de sélections',caption);
  exit;
  end;

case RGBase.ItemIndex of
0 : Params := Params + 'P;';
1 : Params := Params + 'N;';
end;
case RGValeur.ItemIndex of
0 : Params := Params + 'B;';
1 : Params := Params + 'M;';
end;
case CBFermes.State of
cbUnChecked : Params := Params + 'N;';
cbChecked   : Params := Params + 'O;';
cbGrayed    : Params := Params + 'T;';
end;

if RB_Long.Checked then Params := Params + ';' + 'L;'
                   else Params := Params + ';' + 'C;' + IntToStr(SE_Histo.Value); // sur les N derniers mois

Params := Params + ';' + GZZ_ARTICLE.Text;
Params := Params + ';' + GZZ_ARTICLE_.Text;
Params := Params + ';' + GZZ_DEPOT.Value;
Params := Params + ';' + GZZ_DEPOT_.Value;
Params := Params + ';' + GA_FAMILLENIV1.Value;
Params := Params + ';' + GA_FAMILLENIV1_.Value;
Params := Params + ';' + GA_FAMILLENIV2.Value;
Params := Params + ';' + GA_FAMILLENIV2_.Value;
Params := Params + ';' + GA_FAMILLENIV3.Value;
Params := Params + ';' + GA_FAMILLENIV3_.Value;
Params := Params + ';' + DATEDEB.Text;
Params := Params + ';' + DATEDEB_.Text;
Params := Params + ';' + PRIORITE1.Value;
Params := Params + ';' + PRIORITE2.Value;
Params := Params + ';' + PRIORITE3.Value;
Params := Params + ';' + PRIORITE4.Value;
// NEWONE
// --
if TcouvSimple.checked then Params := Params + ';D'  // départ des documents
                        else Params := Params + ';A'; // Depart des Articles
if gestHisto.checked then Params := Params + ';H' else Params := Params + ';S';

if TcouvSimple.checked then
begin
	if CSelDocs.checked then Params := Params+';D'
  										else Params := Params+';S';
end else
begin
	Params := Params+';';
end;

if TcouvSimple.checked then
begin
  if CWithStocks.Checked  then Params := Params + ';O'
  							 					else Params := Params + ';N';
end else
begin
  Params := Params + ';O'
end;
// --
Entree_ResulSuggestion([Params, Recap.Text], 1);
close;
end;


procedure TFAssistReappro.RB_CourtClick(Sender: TObject);
begin
  inherited;
  se_histo.Enabled := true;
end;

procedure TFAssistReappro.RB_LongClick(Sender: TObject);
begin
  inherited;
  se_histo.Enabled := false;
end;


procedure TFAssistReappro.GZZ_ARTICLEElipsisClick(Sender: TObject);
begin
  inherited;
GZZ_ARTICLE.Text := SelectionArticle(GZZ_ARTICLE);
end;

procedure TFAssistReappro.GZZ_ARTICLE_ElipsisClick(Sender: TObject);
begin
  inherited;
GZZ_ARTICLE_.Text := SelectionArticle(GZZ_ARTICLE_);
end;

procedure TFAssistReappro.DATEDEBElipsisClick(Sender: TObject);
begin
  inherited;
DATEDEB.Text := SelectionDate(DATEDEB);
end;

procedure TFAssistReappro.DATEDEB_ElipsisClick(Sender: TObject);
begin
  inherited;
DATEDEB_.Text := SelectionDate(DATEDEB_);
end;

procedure TFAssistReappro.PRIORITE1Change(Sender: TObject);
var s_Select : string;
begin
  inherited;
if PRIORITE1.Value = '' then
    begin
    PRIORITE2.Values.Clear;
    PRIORITE2.Items.Clear;
    PRIORITE2.Enabled := False;
    PRIORITE3.Values.Clear;
    PRIORITE3.Items.Clear;
    PRIORITE3.Enabled := False;
    PRIORITE4.Values.Clear;
    PRIORITE4.Items.Clear;
    PRIORITE4.Enabled := False;
    Exit;
    end;
if PRIORITE2.Values.IndexOf(PRIORITE1.Value) <> -1 then
    begin
    s_Select := PRIORITE2.Value;
    PRIORITE2.Items := PRIORITE.Items;
    PRIORITE2.Values := PRIORITE.Values;
    PRIORITE2.Items.Delete(PRIORITE2.Values.IndexOf(PRIORITE1.Value));
    PRIORITE2.Values.Delete(PRIORITE2.Values.IndexOf(PRIORITE1.Value));
    PRIORITE2.ItemIndex := PRIORITE2.Values.IndexOf(s_Select);
    end;
if PRIORITE3.Values.IndexOf(PRIORITE1.Value) <> -1 then
    begin
    s_Select := PRIORITE3.Value;
    PRIORITE3.Items := PRIORITE.Items;
    PRIORITE3.Values := PRIORITE.Values;
    PRIORITE3.Items.Delete(PRIORITE3.Values.IndexOf(PRIORITE1.Value));
    PRIORITE3.Values.Delete(PRIORITE3.Values.IndexOf(PRIORITE1.Value));
    PRIORITE3.ItemIndex := PRIORITE3.Values.IndexOf(s_Select);
    end;
if PRIORITE4.Values.IndexOf(PRIORITE1.Value) <> -1 then
    begin
    s_Select := PRIORITE4.Value;
    PRIORITE4.Items := PRIORITE.Items;
    PRIORITE4.Values := PRIORITE.Values;
    PRIORITE4.Items.Delete(PRIORITE4.Values.IndexOf(PRIORITE1.Value));
    PRIORITE4.Values.Delete(PRIORITE4.Values.IndexOf(PRIORITE1.Value));
    PRIORITE4.ItemIndex := PRIORITE4.Values.IndexOf(s_Select);
    end;
PRIORITE2.Enabled := True;
end;

procedure TFAssistReappro.PRIORITE2Change(Sender: TObject);
var s_Select : string;
begin
  inherited;
if PRIORITE2.Value = '' then
    begin
    PRIORITE3.Values.Clear;
    PRIORITE3.Items.Clear;
    PRIORITE3.Enabled := False;
    PRIORITE4.Values.Clear;
    PRIORITE4.Items.Clear;
    PRIORITE4.Enabled := False;
    Exit;
    end;
if PRIORITE3.Values.IndexOf(PRIORITE2.Value) <> -1 then
    begin
    s_Select := PRIORITE3.Value;
    PRIORITE3.Items := PRIORITE.Items;
    PRIORITE3.Values := PRIORITE.Values;
    PRIORITE3.Items.Delete(PRIORITE3.Values.IndexOf(PRIORITE1.Value));
    PRIORITE3.Values.Delete(PRIORITE3.Values.IndexOf(PRIORITE1.Value));
    PRIORITE3.Items.Delete(PRIORITE3.Values.IndexOf(PRIORITE2.Value));
    PRIORITE3.Values.Delete(PRIORITE3.Values.IndexOf(PRIORITE2.Value));
    PRIORITE3.ItemIndex := PRIORITE3.Values.IndexOf(s_Select);
    end;
if PRIORITE4.Values.IndexOf(PRIORITE2.Value) <> -1 then
    begin
    s_Select := PRIORITE4.Value;
    PRIORITE4.Items := PRIORITE.Items;
    PRIORITE4.Values := PRIORITE.Values;
    PRIORITE4.Items.Delete(PRIORITE4.Values.IndexOf(PRIORITE1.Value));
    PRIORITE4.Values.Delete(PRIORITE4.Values.IndexOf(PRIORITE1.Value));
    PRIORITE4.Items.Delete(PRIORITE4.Values.IndexOf(PRIORITE2.Value));
    PRIORITE4.Values.Delete(PRIORITE4.Values.IndexOf(PRIORITE2.Value));
    PRIORITE4.ItemIndex := PRIORITE4.Values.IndexOf(s_Select);
    end;
PRIORITE3.Enabled := True;
end;

procedure TFAssistReappro.PRIORITE3Change(Sender: TObject);
var s_Select : string;
begin
  inherited;
if PRIORITE3.Value = '' then
    begin
    PRIORITE4.Values.Clear;
    PRIORITE4.Items.Clear;
    PRIORITE4.Enabled := False;
    Exit;
    end;
if PRIORITE4.Values.IndexOf(PRIORITE3.Value) <> -1 then
    begin
    s_Select := PRIORITE4.Value;
    PRIORITE4.Items := PRIORITE.Items;
    PRIORITE4.Values := PRIORITE.Values;
    PRIORITE4.Items.Delete(PRIORITE4.Values.IndexOf(PRIORITE1.Value));
    PRIORITE4.Values.Delete(PRIORITE4.Values.IndexOf(PRIORITE1.Value));
    PRIORITE4.Items.Delete(PRIORITE4.Values.IndexOf(PRIORITE2.Value));
    PRIORITE4.Values.Delete(PRIORITE4.Values.IndexOf(PRIORITE2.Value));
    PRIORITE4.Items.Delete(PRIORITE4.Values.IndexOf(PRIORITE3.Value));
    PRIORITE4.Values.Delete(PRIORITE4.Values.IndexOf(PRIORITE3.Value));
    PRIORITE4.ItemIndex := PRIORITE4.Values.IndexOf(s_Select);
    end;
PRIORITE4.Enabled := True;
end;

procedure TFAssistReappro.PRIORITE4Change(Sender: TObject);
begin
  inherited;
if PRIORITE4.Value = '' then Exit;
end;


function TFAssistReappro.SelectionArticle (ART : THCritMaskEdit) : string;
begin
//ART := THCritMaskEdit.Create(Self);
//ART.Visible := False;
ART.DataType := 'GCARTICLE';
Result:=ART.Text;
DispatchRecherche(ART, 1, '', ';RETOUR_CODEARTICLE=X', '');
if ART.Text <> '' then Result := Trim(Copy(ART.Text, 0, Length(ART.Text) - 1));
end;

function TFAssistReappro.SelectionDate(Parent : TComponent) : string;
begin
Result := '';
ART := THCritMaskEdit.Create (Parent);
ART.Top := THCritMaskEdit(Parent).Top + THCritMaskEdit(Parent).Height;
ART.Visible := False;
ART.OpeType:=otDate;
GetDateRecherche (TForm(ART.Owner), ART) ;
if ART.Text <> '' then Result := ART.Text;
end;

procedure TFAssistReappro.GA_FAMILLENIV1Change(Sender: TObject);
var i_ind1 : integer;
begin
  inherited;
GA_FAMILLENIV1_.Enabled := True;
if GA_FAMILLENIV1.Value = '' then
    GA_FAMILLENIV1_.Enabled := False
    else
    begin
    GA_FAMILLENIV1_.ReLoad;
    for i_ind1 := GA_FAMILLENIV1_.Values.Count - 1 downto 0 do
        if GA_FAMILLENIV1_.Values.Strings[i_ind1] < GA_FAMILLENIV1.Value then
            begin
            GA_FAMILLENIV1_.Values.Delete(i_ind1);
            GA_FAMILLENIV1_.Items.Delete(i_ind1);
            end;
    GA_FAMILLENIV1_.ItemIndex := 0;
    end;
end;

procedure TFAssistReappro.GA_FAMILLENIV2Change(Sender: TObject);
var i_ind1 : integer;
begin
  inherited;
GA_FAMILLENIV2_.Enabled := True;
if GA_FAMILLENIV2.Value = '' then
    GA_FAMILLENIV2_.Enabled := False
    else
    begin
    GA_FAMILLENIV2_.ReLoad;
    for i_ind1 := GA_FAMILLENIV2_.Values.Count - 1 downto 0 do
        if GA_FAMILLENIV2_.Values.Strings[i_ind1] < GA_FAMILLENIV2.Value then
            begin
            GA_FAMILLENIV2_.Values.Delete(i_ind1);
            GA_FAMILLENIV2_.Items.Delete(i_ind1);
            end;
    GA_FAMILLENIV2_.ItemIndex := 0;
    end;
end;

procedure TFAssistReappro.GA_FAMILLENIV3Change(Sender: TObject);
var i_ind1 : integer;
begin
  inherited;
GA_FAMILLENIV3_.Enabled := True;
if GA_FAMILLENIV3.Value = '' then
    GA_FAMILLENIV3_.Enabled := False
    else
    begin
    GA_FAMILLENIV3_.ReLoad;
    for i_ind1 := GA_FAMILLENIV3_.Values.Count - 1 downto 0 do
        if GA_FAMILLENIV3_.Values.Strings[i_ind1] < GA_FAMILLENIV3.Value then
            begin
            GA_FAMILLENIV3_.Values.Delete(i_ind1);
            GA_FAMILLENIV3_.Items.Delete(i_ind1);
            end;
    GA_FAMILLENIV3_.ItemIndex := 0;
    end;
end;

procedure TFAssistReappro.GZZ_DEPOTChange(Sender: TObject);
begin
  inherited;
  if GZZ_DEPOT.Value = '' then
      GZZ_DEPOT_.ItemIndex := 0
      else
      GZZ_DEPOT_.ItemIndex := GZZ_DEPOT_.Values.IndexOf(GZZ_DEPOT.Value);
end;


procedure AGLAssist_Suggestion ( Parms : array of variant ; nb : integer ) ;
BEGIN
Assist_Suggestion ;
END ;

procedure TFAssistReappro.TcouvSimpleClick(Sender: TObject);
begin
  inherited;
  AjustePageControl;
  CSelDocs.enabled := true;
  gestHisto.enabled :=false;
  Gesthisto.Checked := false;
end;

procedure TFAssistReappro.TBesoinStockClick(Sender: TObject);
begin
  inherited;
  CSelDocs.enabled := false;
  AjustePageControl;
  gestHisto.enabled := true;
end;

procedure TFAssistReappro.GestHistoClick(Sender: TObject);
begin
  inherited;
  AjustePageControl;
end;

procedure TFAssistReappro.AjustePageControl;
var Indice : integer;
begin
  NbrPage := 0;
  for Indice := 0 to P.PageCount -1 do
  begin
    if not (GestHisto.Checked) then
      begin
        if P.pages[indice].Name = 'TbesoinP2' then
        begin
          P.pages[Indice].tabvisible := false;
          continue;
        end;
      end;
    if (GestHisto.Checked) then
      begin
        if P.pages[indice].Name = 'TbesoinP1' then
        begin
          P.pages[Indice].tabvisible := true;
          P.pages[Indice].PageIndex := 2;
          inc(NbrPage);
          continue;
        end;
        if P.pages[indice].Name = 'TbesoinP2' then
        begin
          P.pages[Indice].tabvisible := true;
          P.pages[Indice].PageIndex := 3;
          inc(NbrPage);
          continue;
        end;
      end;
    if (TCouvSimple.checked) then
      begin
        if P.pages[indice].Name = 'TbesoinP0' then
        begin
          P.pages[Indice].tabvisible := False;
          continue;
        end;
        if P.pages[indice].Name = 'TbesoinP1' then
        begin
          P.pages[Indice].tabvisible := true;
          P.pages[Indice].PageIndex := 1;
          inc(NbrPage);
          continue;
        end;
      end;
    if (not (TCouvSimple.checked)) and (not (GestHisto.Checked)) then
      begin
        if P.pages[indice].Name = 'TbesoinP0' then
        begin
          P.pages[Indice].tabvisible := True;
          P.pages[Indice].PageIndex := 1;
          inc(NbrPage);
          continue;
        end;
        if P.pages[indice].Name = 'TbesoinP1' then
        begin
          P.pages[Indice].tabvisible := false;
          continue;
        end;
      end;
    P.pages[Indice].PageIndex := NbrPage;
    inc(NbrPage);
  end;
  DisplayStep;
end;

procedure TFAssistReappro.bPrecedentClick(Sender: TObject);
begin
  inherited;
  if P.activePageindex = NbrPage-1 then bSuivant.Enabled := false
                                   else bSuivant.Enabled := true;

  if (bSuivant.Enabled) then bFin.Enabled := False
                        else bFin.Enabled := True;
  DisplayStep;
end;

procedure TFAssistReappro.DisplayStep;
var st1, NumPage : string;
begin
{$IFDEF BTP}
lEtape.Caption := Msg.Mess[0] + ' ' + IntToStr(PageNumber) + '/' + IntToStr(NbrPage) ;
{$ELSE}
//  inherited;
st1 := P.ActivePage.Name;
if st1 = 'TInitiale' then
  NumPage := '1'
else if st1 = 'TRecap' then
  NumPage := '8'
else
  NumPage := IntToStr(StrToInt(Copy(st1, Length(st1), 1)) + 2);
lEtape.Caption := Msg.Mess[0] + ' ' + NumPage + '/' + IntToStr(P.PageCount);
{$ENDIF}
end;

procedure TFAssistReappro.RGValeurClick(Sender: TObject);
begin
  inherited;
  if RGValeur.ItemIndex = 0 then
  begin
    RGBase.ItemIndex := 0;
    RGBase.Visible := False;
  end
  else
  begin
    RGBase.ItemIndex := 0;
    RGBase.Visible := True;
  end;
end;

Initialization
RegisterAglProc('Assist_Suggestion',False,0,AGLAssist_Suggestion) ;
end.
