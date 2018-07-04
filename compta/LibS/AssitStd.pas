unit AssitStd;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls,
  ActnList, Mask, HPanel, UTOB, HQuickRP, dbtables, ImgList, Menus,
  HEnt1,LicUtil,MajTable, PGIExec, ParamSoc,  Contabon_TOM,
  RefAuto_TOM,  // ParamLibelle
{$IFDEF EAGLCLIENT}
  MainEagl,
  UtileAGL,
  Exercice_TOM,
{$ELSE}
  Fe_Main,
  Exercice,
{$ENDIF}
  HStatus,
  // Uses Commun
  FichComm,
  // Uses Compta
  Ent1,
  Devise_TOM,
  AssistExo,
  Spin;

function LanceAssistantStd (bInit : boolean) : boolean;


type
  TFAssistStd = class(TFAssist)
    TabSheet1: TTabSheet;
    HLabel2: THLabel;
    bAnalytique: TToolbarButton97;
    PopParam: TPopupMenu;
    Activer: TMenuItem;
    Desactiver: TMenuItem;
    ImageList: TImageList;
    bTVA: TToolbarButton97;
    bSaisieQte: TToolbarButton97;
    procedure ActiverClick(Sender: TObject);
    procedure DesactiverClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure OnBoutonActivationClick    ( Sender : TObject);
    procedure OnBoutonActivationMouseDown( Sender : TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure OnBoutonMouseEnter         ( Sender : TObject);

  private
    CurToolBarButtonName : string;
    CurToolbarButton : TToolbarButton97;
    bRetour : boolean;
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var FAssistStd: TFAssistStd;

implementation

uses uLibStdCpta,
     AGLStdCompta;

{$R *.DFM}

function LanceAssistantStd (bInit : boolean) : boolean;
begin
  FAssistStd := TFAssistStd.Create (Application);
  try
    FAssistStd.bRetour := false;
    if GetParamSocSecur('SO_CPPCLSANSANA', True) = True then
      FAssistStd.bAnalytique.ImageIndex := 0
    else
      FAssistStd.bAnalytique.ImageIndex := 1;

    if GetParamSocSecur('SO_CPPCLSAISIEQTE', False) = True then
      FAssistStd.bSaisieQte.ImageIndex := 1
    else
      FAssistStd.bSaisieQte.ImageIndex := 0 ;

    if GetParamSocSecur('SO_CPPCLSAISIETVA',False) = True then
      FAssistStd.bTva.ImageIndex := 1
    else
      FAssistStd.bTva.ImageIndex := 0 ;

    FAssistStd.ShowModal;
  finally
    FAssistStd.Free;
  end;
  Result :=  FAssistStd.bRetour;
end;


procedure TFAssistStd.ActiverClick(Sender: TObject);
begin
  inherited;
  CurToolBarButton.ImageIndex := 1;
  if CurToolBarButtonName = 'bAnalytique' then
  begin
     SetParamSoc ('SO_ZSAISIEANAL',TRUE);
     SetParamSoc ('SO_ZGEREANAL',TRUE);
     SetParamSoc ('SO_CPPCLSANSANA',False);
  end
  else
  if CurToolBarButtonName = 'bTVA' then
  begin
    if VH^.AnaCroisaxe then
      PgiInfo('Gestion de la TVA impossible. Analytiques "croise-axes" activés')
    else
    begin
      SetParamSoc('SO_CPPCLSAISIETVA', True);
      // Ajout GCO - 06/12/2005
      if GetParamSocSecur('SO_CPPCLSANSANA', False) then
      begin
        SetParamSoc('SO_ZSAISIEANAL', True);
        SetParamSoc('SO_ZGEREANAL', True);
        SetParamSoc('SO_CPPCLSANSANA', False);
      end;
      InitialisationTVA;
    end;
  end
  else
  if CurToolBarButtonName = 'bSaisieQte' then
    SetParamSoc('SO_CPPCLSAISIEQTE', true );
end;

procedure TFAssistStd.DesactiverClick(Sender: TObject);
begin
  inherited;
  CurToolBarButton.ImageIndex := 0;
  if CurToolBarButtonName = 'bAnalytique' then
  begin
   if HShowMessage('0;Suppression analytiques;'+'Confimez-vous la suppression de l''ensemble des données analytiques'+';Q;YN;N;N;','','')<>mrYes then
    begin   CurToolBarButton.ImageIndex := 1; exit ;
    end
    else
      DesactivationAnalytique;
  end
  else if CurToolBarButtonName = 'bTVA' then
  begin
    Pgiinfo('Désactivation non disponible dans cette version.', 'Gestion de la TVA');
    bTva.ImageIndex := 1;
    //SetParamSoc('SO_CPPCLSAISIETVA', false )
  end
  else
  if CurToolBarButtonName = 'bSaisieQte' then
    SetParamSoc('SO_CPPCLSAISIEQTE', false);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFAssistStd.OnBoutonActivationClick(Sender: TObject);
begin
  inherited;
  if Sender.ClassName <> 'TToolbarButton97' then exit;
  CurToolBarButtonName := TToolbarButton97(Sender).Name;
  CurToolbarButton := TToolbarButton97(Sender);
  if TToolbarButton97(Sender).ImageIndex = 0 then
  begin
    PopParam.Items[0].Visible := True;
    PopParam.Items[1].Visible := False;
    PopParam.Popup(mouse.CursorPos.x,mouse.CursorPos.y);
  end;

  if TToolbarButton97(Sender).ImageIndex = 1 then
  begin
    inherited;
    if CurToolBarButtonName = 'bAnalytique' then
    begin
      // ajout me
      if GetParamSoc ('SO_CPPCLSANSANA') = FALSE then
      begin
        SetParamSoc ('SO_ZSAISIEANAL', TRUE);
        SetParamSoc ('SO_ZGEREANAL', TRUE);
      end
      else
      begin
        SetParamSoc ('SO_ZSAISIEANAL', FALSE);
        SetParamSoc ('SO_ZGEREANAL', FALSE);
      end;
      AGLLanceFiche('CP','ASSANALYTIQUE','','','');
    end
    else
    if CurToolBarButtonName = 'bTVA' then
    begin
      if not GetParamSocSecur('SO_CPPCLSAISIETVA',False) then
      begin
        SetParamSoc('SO_CPPCLSAISIETVA', True);
        if GetParamSocSecur('SO_CPPCLSANSANA', False) then
        begin
          SetParamSoc('SO_ZSAISIEANAL', True);
          SetParamSoc('SO_ZGEREANAL', True);
          SetParamSoc('SO_CPPCLSANSANA', False);
        end;
      end;
    end
    else
    if CurToolBarButtonName = 'bSaisieQte' then
      SetParamSoc('SO_CPPCLSAISIEQTE', not GetParamSocSecur('SO_CPPCLSAISIEQTE', False) )
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFAssistStd.OnBoutonActivationMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  inherited;
  if CurToolBarButton = nil then exit;
  if ((Button = mbRight) and (CurToolBarButton.ImageIndex=1)) then
  begin
    PopParam.Items[0].Visible := False;
    PopParam.Items[1].Visible := True;
    PopParam.Popup(mouse.CursorPos.x,mouse.CursorPos.y);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFAssistStd.OnBoutonMouseEnter(Sender: TObject);
begin
  inherited;
  CurToolbarButton := TToolbarButton97(Sender);
  CurToolBarButtonName := TToolbarButton97(Sender).Name;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFAssistStd.bFinClick(Sender: TObject);
begin
  inherited;
  Close;
end;

////////////////////////////////////////////////////////////////////////////////

end.
