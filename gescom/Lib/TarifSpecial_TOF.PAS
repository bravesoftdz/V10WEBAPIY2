{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 13/11/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : TARIFSPECIAL ()
Mots clefs ... : TOF;TARIFSPECIAL
*****************************************************************}
Unit TarifSpecial_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     {$IFNDEF EAGLCLIENT}
        db,
        dbTables,
        Fe_Main,
        Fiche,
        Mul,
     {$ELSE}
        MainEagl,
        eFiche,
        eMul,
     {$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     UTOF,
     SaisieList,
     uTableFiltre
     ;

Type
  TOF_TARIFSPECIAL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

    procedure TreeEntete_DblClick( Sender: TObject );
    procedure DoSetNavigate( Sender: TObject );
    procedure bInsertFonctionnalite_OnClick( Sender: TObject );

  private
    TF: TTableFiltre;
    bInsertFonctionnalite: tToolbarButton97;

  end ;

Implementation

{-------------------------------------------------------------------------------
--------------------------------------------------------------------------------}
procedure TOF_TARIFSPECIAL.OnNew ;
begin
  Inherited ;
end ;

{-------------------------------------------------------------------------------
--------------------------------------------------------------------------------}
procedure TOF_TARIFSPECIAL.OnDelete ;
begin
  Inherited ;
end ;

{-------------------------------------------------------------------------------
--------------------------------------------------------------------------------}
procedure TOF_TARIFSPECIAL.OnUpdate ;
begin
  Inherited ;
end ;

{-------------------------------------------------------------------------------
--------------------------------------------------------------------------------}
procedure TOF_TARIFSPECIAL.OnLoad ;
begin
  Inherited ;
end ;

{-------------------------------------------------------------------------------
--------------------------------------------------------------------------------}
procedure TOF_TARIFSPECIAL.OnArgument (S : String ) ;
begin
  Inherited ;
  if (Ecran<>nil) and (Ecran is TFSaisieList ) then
  begin
    TF := TFSaisieList(Ecran).LeFiltre;
    TF.OnSetNavigate := DoSetNavigate;
  end;

  if (GetControl('TREEENTETE')<>nil) then THTreeView(GetControl( 'TreeEntete' )).OnDblClick     := TreeEntete_DblClick;

  //Bouton de création de l'entête : coloris
  bInsertFonctionnalite := tToolbarButton97(GetControl('BINSERTFONCTIONNALITE'));
  if (GetControl('BINSERTFONCTIONNALITE')<>nil) then TToolBarButton97(GetControl('BINSERTFONCTIONNALITE')).OnClick := bInsertFonctionnalite_OnClick ;
end ;

{-------------------------------------------------------------------------------
--------------------------------------------------------------------------------}
procedure TOF_TARIFSPECIAL.OnClose ;
begin
  Inherited ;
end ;

{-------------------------------------------------------------------------------
--------------------------------------------------------------------------------}
procedure TOF_TARIFSPECIAL.DoSetNavigate(Sender: TObject);
begin
  SetControlEnabled('BINSERTFONCTIONNALITE',TF.CanReload);
end;

{-------------------------------------------------------------------------------
   Accès en modification de la fiche d'entête au DblClick sur le TreeView
--------------------------------------------------------------------------------}
procedure TOF_TARIFSPECIAL.TreeEntete_DblClick( Sender: TObject );
var
  sFonctionnalite: String;
begin
  sFonctionnalite := AGLLanceFiche('Y','YTARIFSFONCT_FIC','',TF.TOBFiltre.GetValue('YFO_FONCTIONNALITE'),'ACTION=MODIFICATION;MONOFICHE;WCO_COLORIS='+TF.TOBFiltre.GetValue('YFO_FONCTIONNALITE'));
  if (sFonctionnalite<>'') then
    TF.RefreshEntete( sFonctionnalite );
end;

{-------------------------------------------------------------------------------
   Modifocation des paramètres d'une fonctionnalite
--------------------------------------------------------------------------------}
procedure TOF_TARIFSPECIAL.bInsertFonctionnalite_OnClick( Sender: TObject );
var
  sFonctionnalite: String;
begin
  sFonctionnalite := AGLLanceFiche('Y','YTARIFSFONCT_FIC','',TF.TOBFiltre.GetValue('YFO_FONCTIONNALITE'),'ACTION=MODIFICATION;MONOFICHE;WFO_FONCTIONNALITE='+TF.TOBFiltre.GetValue('YFO_FONCTIONNALITE'));
  if (sFonctionnalite<>'') then
    TF.RefreshEntete(sFonctionnalite);
end;

Initialization
  registerclasses ( [ TOF_TARIFSPECIAL ] ) ;
end.
