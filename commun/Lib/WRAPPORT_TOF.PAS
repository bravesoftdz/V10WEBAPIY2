{***********UNITE*************************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 23/09/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : WRAPPORT ()
Mots clefs ... : TOF;WRAPPORT
*****************************************************************}
Unit WRAPPORT_TOF ;

Interface

Uses
  StdCtrls,
  Controls,
  Classes,
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  uTob,
  wTof,
  HTB97,
  Menus,
  wCommuns,
  graphics,
  Grids,
  wRapport
  ;

Type
  TOF_WRAPPORT = Class (twTOF)
    procedure OnNew                  ; override;
    procedure OnDelete               ; override;
    procedure OnUpdate               ; override;
    procedure OnLoad                 ; override;
    procedure OnArgument(S : String ); override;
    procedure OnClose                ; override;
  private
   	{var}
    Rapport: TWRapport;
    TobWRPError, TobWRPDone: Tob;
    GRILLE, GRILLE2: THGrid;
    ouError, ouDone: Integer;
		{ Impression }
    procedure mnEditRapp_OnClick(Sender: tObject);
    procedure mnListRapp_OnClick(Sender: tObject);
    { événements... }
    procedure PCDel_OnChange(Sender: tObject);
    procedure BTShowMemo_OnClick(Sender: tObject);
    procedure BTLOG_OnClick(Sender: tObject);
    {... de la grille }
    procedure GRILLE_OnRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
  end ;

Implementation

uses
  Windows,
  LookUp,
  UtilPGI
  ;

procedure TOF_WRAPPORT.OnNew;
begin
  Inherited;
end;

procedure TOF_WRAPPORT.OnDelete;
begin
  Inherited;
end;

procedure TOF_WRAPPORT.OnUpdate;
begin
  Inherited;
end;

procedure TOF_WRAPPORT.OnLoad;
begin
  Inherited;
end;

procedure TOF_WRAPPORT.OnArgument(S : String);
var
  Cancel: Boolean;
	i: Integer;
begin
  Inherited ;

  Rapport := TWRapport(GetArgumentInteger(S, WRPTobFieldName));
  TobWRPDone := Rapport.GetTobWRP(TWRP_Done);
  TobWRPError := Rapport.GetTobWRP(TWRP_Error);

  { controls }
  if Assigned(GetControl('GRID')) then
  begin
  	GRILLE := THGrid(GetControl('GRID'));
    GRILLE.OnRowEnter := GRILLE_OnRowEnter;
  end;
	if Assigned(GetControl('GRID2')) then
  begin
  	GRILLE2 := THGrid(GetControl('GRID2'));
    GRILLE2.OnRowEnter := GRILLE_OnRowEnter;
  end;
  { événements }
  if Assigned(GetControl('MNEDITRAPP')) then
    tMenuItem(GetControl('MNEDITRAPP')).OnClick := mnEditRapp_OnClick;
  if Assigned(GetControl('MNLISTRAPP')) then
    tMenuItem(GetControl('MNLISTRAPP')).OnClick := mnListRapp_OnClick;
  if Assigned(GetControl('PCDEL')) then
    tPageControl(GetControl('PCDEL')).OnChange := PCDel_OnChange;
  if Assigned(GetControl('BTShowMemo')) then
    TToolBarButton97(GetControl('BTShowMemo')).OnClick := BTShowMemo_OnClick;
  if Assigned(GetControl('BTLOG')) then
    TToolBarButton97(GetControl('BTLOG')).OnClick := BTLOG_OnClick;
  if Assigned(GetControl('LBTITRE')) then
    THLABEL(GetControl('LBTITRE')).Caption := Rapport.TitreRapport;

  TobWRPError.PutGridDetail(GRILLE, False, False, 'IDWRP;IDENTIFIANTWXX;TEXTE', True);
  TobWRPDone.PutGridDetail(GRILLE2, False, False, 'IDWRP;IDENTIFIANTWXX;TEXTE', True);

  for i := 0 to TPageControl(GetControl('PCDEL')).PageCount - 1 do
    TPageControl(GetControl('PCDEL')).Pages[i].ImageIndex := -1;

  if TobWRPError.Detail.Count > 0 then
  begin
  	ChargeImageList;
  	TPageControl(GetControl('PCDEL')).Images := V_Pgi.GraphList;
    TPageControl(GetControl('PCDEL')).Pages[1].ImageIndex := 2-1;
  end;
  if TobWRPDone.Detail.Count > 0 then
  begin
  	ChargeImageList;
  	TPageControl(GetControl('PCDEL')).Images := V_Pgi.GraphList;
    TPageControl(GetControl('PCDEL')).Pages[0].ImageIndex := 12-1;
  end;

  Cancel := False;
  ouError := 1;
  ouDone:= 1;
  if TobWRPDone.Detail.Count = 0 then
  begin
    TPageControl(GetControl('PCDEL')).ActivePage := tTabSheet(GetControl('TSERRORS'));
  	GRILLE.SetFocus;
    GRILLE_OnRowEnter(GRILLE, 1, Cancel, True);
  end
  else
  begin
    TPageControl(GetControl('PCDEL')).ActivePage := tTabSheet(GetControl('TSDONE'));
  	GRILLE2.SetFocus;
    GRILLE_OnRowEnter(GRILLE2, 1, Cancel, True);
  end;
  if GetArgumentString(S, 'TITRECELL1') <> '' then
    GRILLE.Cells[1, 0] := GetArgumentString(S, 'TITRECELL1', false);
end ;

procedure TOF_WRAPPORT.OnClose;
begin
  Inherited;
end;

procedure TOF_WRAPPORT.mnEditRapp_OnClick(Sender: tObject);
begin
  Rapport.Print;
end;

procedure TOF_WRAPPORT.mnListRapp_OnClick(Sender: tObject);
begin
  Rapport.ExportList;
end;

procedure TOF_WRAPPORT.GRILLE_OnRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var
	G: THGrid;
begin
	if Assigned(Sender) and (Sender is thGrid) then
   	G := thGrid(Sender)
  else
  begin
    if tPageControl(GetControl('PCDEL')).ActivePage = tTabSheet(GetControl('TSDONE')) then
      G := GRILLE2
    else if tPageControl(GetControl('PCDEL')).ActivePage = tTabSheet(GetControl('TSERRORS')) then
      G := GRILLE
    else
      G := Nil;
  end;

  if Assigned(G) then
  begin
    if Assigned(GetControl('MEMTEXT')) then
    begin
      tMemo(GetControl('MEMTEXT')).Clear;
      tMemo(GetControl('MEMTEXT')).Lines.Text := G.Cells[2, Ou];
      if UpperCase(G.Name) = 'GRID' then
      	ouError := ou
      else
      	ouDone := ou;
    end;
  end;
end;

procedure TOF_WRAPPORT.PCDel_OnChange(Sender: tObject);
var
	Cancel: Boolean;
begin
	Cancel := False;
	if tPageControl(GetControl('PCDEL')).ActivePage = tTabSheet(GetControl('TSDONE')) then
  begin
  	GRILLE2.SetFocus;
  	GRILLE_OnRowEnter(GRILLE2, ouDone, Cancel, True);
  end
  else
  begin
  	GRILLE.SetFocus;
    GRILLE_OnRowEnter(GRILLE, ouError, Cancel, True);
  end;
end;

procedure TOF_WRAPPORT.BTShowMemo_OnClick(Sender: tObject);
var
  Cancel: Boolean;
  Ou: Integer;
begin
  SetControlVisible('PNBOTTOM', TToolBarButton97(GetControl('BTShowMemo')).Down);
  THSplitter(GetControl('SPLITBOTTOM')).Visible := TToolBarButton97(GetControl('BTShowMemo')).Down;
  Cancel := False;
  if TToolBarButton97(GetControl('BTShowMemo')).Down then
  begin
    if tPageControl(GetControl('PCDEL')).ActivePage = tTabSheet(GetControl('TSDONE')) then
      Ou := GRILLE2.Row
    else if tPageControl(GetControl('PCDEL')).ActivePage = tTabSheet(GetControl('TSERRORS')) then
      Ou := GRILLE.Row
    else
      Ou := 0;
    if Ou <> 0 then GRILLE_OnRowEnter(Self, Ou, Cancel, False);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 14/04/2003
Modifié le ... :   /  /
Description .. : Sauvegarde dans un fichier LOG
Mots clefs ... :
*****************************************************************}
procedure TOF_WRAPPORT.BTLOG_OnClick(Sender: tObject);
begin
  Rapport.SaveToFile;
end;

Initialization
  registerclasses([TOF_WRAPPORT]);
end.
