{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/09/2001
Modifié le ... : 11/02/2003
Description .. : Source TOF de la TABLE : ICCENATIONAUX ()
Suite ........ : GCO - ??\??\????
Suite ........ : Traitement DoubleClick de Grille dans le Script pour faire
Suite ........ : fonctionner les boutons de navigation
Mots clefs ... : TOF;ICCENATIONAUX
*****************************************************************}
Unit uTOFIccENationaux ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     UTOB,
     MaineAGL,  // AGLLanceFiche
     eMUL,      // TFMul( Ecran )
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main,   // AGLLanceFiche
     Mul,       // TFMul( Ecran )
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     uTof,
     Htb97,
     Windows;  // VK_

procedure CPLanceFiche_ICCELEMNATIONAUX;

Type
  TOF_ICCELTNAT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

   private
     BInsert  : TToolBarButton97;
     BValider : TToolBarButton97;
     //G        : THDBGrid;

     procedure OnKeyDownForm(Sender: TObject; var Key: Word; Shift: TShiftState);
     procedure OnClick_BInsert ( Sender : TObject );
     //procedure OnDblClick_G ( Sender : TObject );
   public

   end ;

Implementation

uses uLibExercice,
     uTOMIccENationaux; // CPLanceFiche_ICCPARAMELEMNAT

////////////////////////////////////////////////////////////////////////////////
procedure CPLanceFiche_ICCELEMNATIONAUX;
begin
  AGLLanceFiche('CP', 'ICCMULELTNAT', '', '','');
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_ICCELTNAT.OnNew ;
begin
  Inherited ;
  //THValComboBox(GetControl('ICN_CODETXMOYEN')).ItemIndex := 0;
  //THValComboBox(GetControl('ICN_ETATICC')).ItemIndex := 0;
  //THValComboBox(GetControl('ICN_PREDEFINI')).ItemIndex := 0;
end ;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_ICCELTNAT.OnDelete ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_ICCELTNAT.OnUpdate ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_ICCELTNAT.OnArgument (S : String ) ;
begin
  inherited ;
  Ecran.HelpContext := 999999110;
  SetControlVisible('BAIDE', True);
  BValider := TToolBarButton97(GetControl('BVALIDER'));
  BInsert  := TToolBarButton97(GetControl('BINSERT'));

  BInsert.OnClick := OnClick_BInsert;
  Ecran.OnKeyDown := OnKeyDownForm;

  TTabSheet(GetControl('PCOMPLEMENT', False)).TabVisible := False;

  SetControlVisible('BANNULER', True); // GCO - 06/11/2007 - FQ 21779

  // GCO - Traitement dans le script de navigation
  //G := THDBGrid(GetControl('FLISTE'));
  //G.OnDblClick := OnDblClick_G;
end ;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_ICCELTNAT.OnLoad ;
begin
  Inherited ;
  CExoRefOuvert( True );
end ;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_ICCELTNAT.OnClose ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/09/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : KeyDownForm
*****************************************************************}
procedure TOF_ICCELTNAT.OnKeyDownForm(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_F10    : BValider.Click;
    VK_INSERT : BInsert.Click;
    //VK_DELETE : if Shift = [ssCtrl] then ;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/09/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCELTNAT.OnClick_BInsert(Sender: TObject);
begin
  inherited;
  CPLanceFiche_ICCPARAMELEMNAT(DateToStr(Date), 'ACTION=CREATION');
  TFMUL(ECRAN).BCherche.Click;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/09/2001
Modifié le ... :   /  /
Description .. : Traitement dans le script de la Fiche
Mots clefs ... :
*****************************************************************}
(*
procedure TOF_ICCELTNAT.OnDblClick_G(Sender: TObject);
begin
  inherited;
  CPLanceFiche_ICCPARAMELEMNAT(TFMUL(Ecran).Q.FindField('ICN_DATEVALEUR').AsString, 'ACTION=MODIFICATION');
  TFMUL(ECRAN).BCherche.Click;
end;
*)

Initialization
  registerclasses ( [ TOF_ICCELTNAT ] ) ;
end.
