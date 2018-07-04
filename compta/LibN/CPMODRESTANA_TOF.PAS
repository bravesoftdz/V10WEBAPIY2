{***********UNITE*************************************************
Auteur  ...... :                 
Créé le ...... : 29/12/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPMODRESTANA ()
Mots clefs ... : TOF;CPMODRESTANA
*****************************************************************}
Unit CPMODRESTANA_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     HDB,       // THDBGrid
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_main,
{$else}
     eMul,
     uTob,
     Maineagl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     HTB97,
     Ent1,
     AGLInit,        // TheMULQ
     CPMODRESTANA_TOM ;

  procedure CPLanceFiche_CPMODRESTANA;

Type
  TOF_CPMODRESTANA = Class (TOF)
  private
  {$IFDEF EAGLCLIENT}
    FListe : THGrid;
  {$ELSE}
    FListe : THDBGrid;
  {$ENDIF EAGLCLIENT}
    FOnBChercheClick:   TNotifyEvent;

    {Critères}
    Axe:            THValComboBox;
    Modele:         THEdit;
    BtnInsert:      TToolbarButton97;
    TSComplement:   TTabSheet;

    procedure ListeOnDblClick(Sender : TObject);
    procedure AxeChange(Sender : TObject);
    procedure BInsertOnClick(Sender : TObject);
    procedure BChercheClick(sender: TObject);
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

procedure CPLanceFiche_CPMODRESTANA;
begin
  AGLLanceFiche('CP','CPMODRESTANA_MUL','','','');
end;

procedure TOF_CPMODRESTANA.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPMODRESTANA.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPMODRESTANA.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPMODRESTANA.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPMODRESTANA.OnArgument (S : String ) ;
begin
  Inherited ;
  Ecran.HelpContext := 1461000; //Lek 040706 1470000;  {FP 19/04/2006 FQ17839}

{$IFDEF EAGLCLIENT}
  FListe := THGrid(GetControl('FListe', True));
{$ELSE}
  FListe := THDBGrid(GetControl('FListe', True));
{$ENDIF EAGLCLIENT}
  TSComplement := TTabSheet(GetControl('PCOMPLEMENT', True));

  {Critères}
  Axe       := THValComboBox(Getcontrol('CRA_AXE', true));
  Modele    := THEdit(Getcontrol('CRA_CODE', true));
  BtnInsert := TToolbarButton97(Getcontrol('BInsert', true));

  FListe.OnDblClick := ListeOnDblClick;
  Axe.OnChange      := AxeChange;
  BtnInsert.OnClick := BInsertOnClick;
  FOnBChercheClick  := TToolBarButton97(GetControl('BCherche')).OnClick  ;
  TToolBarButton97(GetControl ('BCherche')).OnClick := BChercheClick ;

  TSComplement.Visible    := False;
  TSComplement.TabVisible := False;

  if VH^.AnaCroisaxe then
    begin
    SetControlVisible('TCRA_AXE', False);
    Axe.Visible := False;
    Getcontrol('TCRA_CODE', true).Top := Getcontrol('TCRA_AXE', true).Top;
    Getcontrol('CRA_CODE',  true).Top := Getcontrol('CRA_AXE',  true).Top;
    end;
  Axe.ItemIndex := 0;
  Axe.ItemIndex := GetFirstAxeAna(Axe.Value)-1;
  AxeChange(Axe);

  THValComboBox(GetControl('CRA_EXCLU')).ItemIndex := 0;
end ;

procedure TOF_CPMODRESTANA.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPMODRESTANA.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPMODRESTANA.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPMODRESTANA.BInsertOnClick(Sender: TObject);
begin
  if VH^.JalCtrlBud<>'' then
    begin
    PGIInfo('Le contrôle en saisie sur le budget est incompatible avec la création d''un modèle de restrictions analytiques.'+Chr($D)+Chr($A)+
            'Si vous voulez créer un modèle vous devez décocher le paramètre de contrôle dans les paramètres société.');
    end
  else
    begin
    CPLanceFiche_CPMODRESTANA_TOM('', '', taCreat);
    BChercheClick(nil) ;
    end;
end;

procedure TOF_CPMODRESTANA.BChercheClick(sender: TObject);
begin
  if Assigned ( FOnBChercheClick ) then
    FOnBChercheClick( Sender ) ;
end;

procedure TOF_CPMODRESTANA.AxeChange(Sender: TObject);
begin
  if Axe.Value <> '' then
    Modele.Plus := 'CRA_AXE="'+Axe.Value+'"'
  else
    Modele.Plus := '';
end;

procedure TOF_CPMODRESTANA.ListeOnDblClick(Sender: TObject);
var
  Axe:    String;
  Modele: String;
begin
  Axe    := GetField('CRA_AXE');
  Modele := GetField('CRA_CODE');
{$IFNDEF EAGLCLIENT}
  TheMulQ:=TFMul(Ecran).Q;
{$ELSE}
  TheMulQ:=TFMul(Ecran).Q.TQ;     
{$ENDIF}
  CPLanceFiche_CPMODRESTANA_TOM(Axe, Modele, taModif);
  BChercheClick(nil) ;
end;

Initialization
  registerclasses ( [ TOF_CPMODRESTANA ] ) ;
end.

