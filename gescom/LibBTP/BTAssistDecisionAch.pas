unit BTAssistDecisionAch;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, ComCtrls, ExtCtrls, HPanel, HTB97,
{$IFDEF EAGLCLIENT}
	mainEagl,
{$ELSE}
  fe_main,
{$ENDIF}

  Hctrls, HRichEdt, HRichOLE, Mask,BTUtilDecisionAch, TntComCtrls,
  TntStdCtrls, TntExtCtrls;

type
  TFassistDecisionAch = class(TFAssist)
    TassistDec1: TTabSheet;
    TabSheetFinal: TTabSheet;
    TText4: THLabel;
    TDATEDEB: THLabel;
    DATEDEB: THCritMaskEdit;
    TDATEFIN: THLabel;
    DATEDEB_: THCritMaskEdit;
    Recap: THRichEditOLE;
    HLabel7: THLabel;
    PanelFin: TPanel;                                         
    TTextFin1: THLabel;
    TTextFin2: THLabel;
    TassistDec3: TTabSheet;
    HLabel1: THLabel;
    CBETABLISSEMENT: THValComboBox;
    HLabel2: THLabel;
    Tassitdec2: TTabSheet;
    TText2: THLabel;
    TSR_ARTICLE: THLabel;
    GZZ_ARTICLE: THCritMaskEdit;
    TSR_ARTICLE_: THLabel;
    GZZ_ARTICLE_: THCritMaskEdit;
    TSR_DEPOT: THLabel;
    GZZ_DEPOT: THValComboBox;
    TSR_DEPOT_: THLabel;
    GZZ_DEPOT_: THValComboBox;
    CBSELDOC: TCheckBox;
    procedure bSuivantClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    function ControleArticleEtDepot : boolean;
    function ControleDate : boolean;
    function ControlePageIndex : boolean;
  private
    { Déclarations privées }
  	ConstructDecision : TConstructDecisionAch;
  	procedure InitZones;
    procedure PrepareAffichage;
  public
    { Déclarations publiques }
  end;


procedure LanceAssistantDecisionnelAch;

implementation
uses UTOF_VideInside;

{$R *.DFM}

procedure LanceAssistantDecisionnelAch;
var Fassist: TFassistDecisionAch;
begin
	AGLLanceFiche('BTP','BTVIDEINSIDE','','','') ;
  //
	Fassist := TFassistDecisionAch.create (Application);
  TRY
  	Fassist.ShowModal;
  FINALLY
  	Fassist.free;
  END;
end;

procedure TFassistDecisionAch.bSuivantClick(Sender: TObject);
begin
  inherited;
  if P.activePageindex = 3 then
  begin
  	bSuivant.Enabled := false;
    bprecedent.Enabled := true;
    PrepareAffichage;
  end else if P.activePageIndex = 0 then
  begin
  	bSuivant.Enabled := true;
    bprecedent.Enabled := false;
  end else
  begin
  	bSuivant.Enabled := true;
    bprecedent.Enabled := true;
  end;
end;

procedure TFassistDecisionAch.FormShow(Sender: TObject);
begin
  inherited;
  InitZones;
  bSuivant.Enabled := true;
  bprecedent.Enabled := false;
end;

procedure TFassistDecisionAch.InitZones;
begin
	recap.lines.Clear;
end;

procedure TFassistDecisionAch.PrepareAffichage;
begin
	recap.lines.clear;
  recap.lines.text := 'Paramètres de sélection des besoins de chantier';
  recap.lines.add ('');
  if DATEDEB.Text = '' then Recap.Lines.Add('  Début de période : ' {+ V_PGI.})
                      else Recap.Lines.Add('  Début de période : ' + DATEDEB.Text);

  if DATEDEB_.Text = '' then Recap.Lines.Add('  Fin de période : ' {+ V_PGI.})
                       else Recap.Lines.Add('  Fin de période : ' + DATEDEB_.Text);

  if (GZZ_ARTICLE.text <> '') and (GZZ_ARTICLE_.Text <> '') then
  begin
  	Recap.Lines.Add('  De l''article : ' + GZZ_ARTICLE.Text);
  	Recap.Lines.Add('  à l''article : ' + GZZ_ARTICLE_.Text);
  end else
  begin
    if (GZZ_ARTICLE.text = '') and (GZZ_ARTICLE_.text = '') then
    begin
  		Recap.Lines.Add('  Sans restriction d''articles');
    end else
    begin
    	if GZZ_ARTICLE.text <> '' then
      begin
  				Recap.Lines.Add('  De l''article : ' + GZZ_ARTICLE.Text+ ' jusqu''à la fin');
      end else
      begin
  				Recap.Lines.Add('  Jusqu''à l''article : ' + GZZ_ARTICLE_.Text);
      end;
    end;
  end;

  if (GZZ_DEPOT.Value = '') and (GZZ_DEPOT_.Value = '') then Recap.Lines.Add('  Pas de selection sur les dépôts')
  else if (GZZ_DEPOT.Value <> '') and (GZZ_DEPOT_.Value <> '') then Recap.Lines.Add('  Dépôts compris entre   ' + GZZ_DEPOT.Text + '   et   ' + GZZ_DEPOT_.Text)
  else if GZZ_DEPOT.Value <> '' then Recap.Lines.Add('  Dépôts à partir de   ' + GZZ_DEPOT.Text + '   jusqu''à la fin')
  else Recap.Lines.Add('  Dépôts à partir du début jusqu''à   ' + GZZ_DEPOT_.Text);

  if CBETABLISSEMENT.Value = '' then Recap.lines.add ('  pour l''ensemble des établissements')
  															else Recap.lines.add ('  pour l''établissement '+CBETABLISSEMENT.Text);
  if CBSELDOC.checked then recap.lines.add ('  Avec sélection de documents');

end;

procedure TFassistDecisionAch.bFinClick(Sender: TObject);
begin
  inherited;
  if not ControlePageIndex then exit;
  With ConstructDecision do
  begin
    PrepareAffichage;
    DateDebut := StrToDate (DATEDEB.text);
    DateFin := StrToDate (DATEDEB_.text);
    ArticleDebut := GZZ_ARTICLE.Text ;
    ArticleFin := GZZ_ARTICLE_.text;
    DepotDebut  := GZZ_DEPOT.Value;
    DepotFin  := GZZ_DEPOT_.Value;
    Etablissement := CBETABLISSEMENT.Value;
    WithSelection := CBSELDOC.Checked;
    MessageRecap :=  Recap;
    ecran := self;
    if Generate then
    begin
      // Appel de la fiche
      LanceFicheSaisie;
    end else
    begin
  		PGIBox ('Aucun document à traiter',self.caption);
    end;
  end;
  close;
end;

procedure TFassistDecisionAch.FormCreate(Sender: TObject);
begin
  inherited;
  ConstructDecision := TConstructDecisionAch.create;
end;

procedure TFassistDecisionAch.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  inherited;
  ConstructDecision.free;
end;

function TFassistDecisionAch.ControlePageIndex: boolean;
begin
  result := ControleDate;
  if result then result := ControleArticleEtDepot;
end;

function TFassistDecisionAch.ControleDate: boolean;
begin
	result := true;
  if strtodate(dateDEB_.text) < strtodate(datedeb.text) then
  begin
    PgiBox ('Incohérence dans les dates de sélections',caption);
    result := false;
  end;
end;

function TFassistDecisionAch.ControleArticleEtDepot: boolean;
begin
  result := true;
  if (GZZ_ARTICLE.text <> '') and (GZZ_ARTICLE_.text <> '') then
  begin
  	if GZZ_ARTICLE.text > GZZ_ARTICLE_.text then
    begin
    	PgiBox ('Incohérence dans la sélection d''article',caption);
      result := false;
      exit;
    end;
  end;
  if (GZZ_DEPOT.text <> '') and (GZZ_DEPOT_.text <> '') then
  begin
  	if GZZ_DEPOT.text > GZZ_DEPOT_.text then
    begin
    	PgiBox ('Incohérence dans la sélection des dépots',caption);
      result := false;
      exit;
    end;
  end;
end;

end.
