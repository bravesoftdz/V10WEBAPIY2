unit utilPGIGr;

interface

uses
  Forms,
  StdCtrls,
  Controls,
  TypInfo,
  Classes,
  ActnList,
{$IFDEF EAGLCLIENT}
  eMul,
{$ELSE}
  Mul,
{$ENDIF}
  HTB97;

function CreateActionFromButton97( aForm: TForm; aToolbarButton97: TToolbarButton97 ): TAction; overload;
function CreateActionFromButton97( aActionList: TActionList; aToolbarButton97: TToolbarButton97 ): TAction; overload;

procedure SetControlChildrenEnabled( Ctrl: TControl; Enabled: Boolean; IgnoreDisabled: Boolean = False );

// Fonctions GetControl, ... utilisables avec les TOF, les TOM et les TForm
function GetControl( pObject: TObject; const pControlName: string;
                     ControlClass: TClass; AvecErreurSiNonTrouve: Boolean = False  ): TControl;
function GetControlEnabled( pObject: TObject; const pControlName: string ): Boolean;
function GetControlVisible( pObject: TObject; const pControlName: string ): Boolean;
function GetControlText( pObject: TObject; const pControlName: string ): string;
function GetControlChecked( Ctrl : TControl ) : Boolean;
function GetControlProperty( pObject: TObject; const pControlName, aProperty: string ): Variant;
procedure SetControlEnabled( pObject: TObject; const pControlName: string; value: Boolean );
procedure SetControlVisible( pObject: TObject; const pControlName: string; value: Boolean );
procedure SetControlText( pObject: TObject; const pControlName, aText: string );
procedure SetControlProperty( pObject: TObject; const pControlName, aProperty: string; value: Variant );
function FListeFieldNameToColIndex( F : TFMul ; ColName : string ) : integer;

implementation

uses
  HCtrls,
  UTOM,
  UTOF;


function CreateActionFromButton97( aForm: TForm; aToolbarButton97: TToolbarButton97 ): TAction;
var
  pComp: TComponent;
  pActionList: TActionList;
begin
  if( Assigned(aForm) and Assigned(aToolbarButton97) ) then
  begin
    pComp := aForm.FindComponent( 'ActionList' );
    if( pComp is TActionList ) then
      pActionList := pComp as TActionList
    else if( not(Assigned(pComp)) ) then
    begin
      pActionList := TActionList.Create( aForm );
      pActionList.Name := 'ActionList';
    end
    else
      pActionList := nil;
    Result := CreateActionFromButton97( pActionList, aToolbarButton97 );
  end
  else
    Result := nil;
end;

function CreateActionFromButton97( aActionList: TActionList; aToolbarButton97: TToolbarButton97 ): TAction;
begin
  if( Assigned(aActionList) and Assigned(aToolbarButton97) ) then
  begin
    Result := TAction.Create( aActionList.Owner );
    if( not(Assigned(aActionList.Images)) ) then
      aActionList.Images := aToolbarButton97.Images;
    Result.ActionList := aActionList;
    Result.Caption := aToolbarButton97.Caption;
    Result.Checked := aToolbarButton97.Down;
    Result.Enabled := aToolbarButton97.Enabled;
    Result.HelpContext := aToolbarButton97.HelpContext;
    Result.Hint := aToolbarButton97.Hint;
    Result.ImageIndex := aToolbarButton97.ImageIndex;
    Result.Visible := aToolbarButton97.Visible;
    Result.OnExecute := aToolbarButton97.OnClick;
    aToolbarButton97.Action := Result;
  end
  else
    Result := nil;
end;

{***********A.G.L.***********************************************
Auteur  ...... : O. Grosclaude
Créé le ...... : 23/09/2004
Modifié le ... :   /  /
Description .. : Active/désactive un contrôle et ses éventuels fils
Mots clefs ... :
*****************************************************************}
procedure SetControlChildrenEnabled( Ctrl: TControl; Enabled: Boolean; IgnoreDisabled: Boolean = False );
var
  Ind: Integer;
begin
  if Ctrl is TControl then
  begin
    if( not(IgnoreDisabled) ) then
      Ctrl.Enabled := Enabled
    else
    begin
      // Valeur du tag pas encore affectée : on affecte -1 aux contrôles désactivés
      if( Ctrl.Tag = 0 ) then
      begin
        if( not(Ctrl.Enabled) ) then
          Ctrl.Tag := -1
        else
          Ctrl.Tag := 1;
      end;
      // On ne modifie que les contrôles initialement activés
      if( Ctrl.Tag > 0 ) then
        Ctrl.Enabled := Enabled;
    end;
    // On appelle la fonction pour les enfants
    if Ctrl is TWinControl then
      for Ind := 0 to TWinControl( Ctrl ).ControlCount - 1 do
        SetControlChildrenEnabled( TWinControl(Ctrl).Controls[Ind], Enabled, IgnoreDisabled );
  end;
end;

function GetControlChecked( Ctrl : TControl ): Boolean;
begin
  if Ctrl is TCheckBox then
    Result := TCheckBox( Ctrl ).Checked
  else if Ctrl is TRadioButton then
    Result := TRadioButton( Ctrl ).Checked
  else
    Result := False;
end;


{***********A.G.L.***********************************************
Auteur  ...... : O. Grosclaude
Créé le ...... : 10/06/2005
Modifié le ... :   /  /
Description .. : Equivalents des méthodes des TOM et TOF : GetControl,
Suite ........ : GetControlEnabled, GetControlVisible, ... pour des TForm
Mots clefs ... :
*****************************************************************}

function GetFormComponent( aForm: TForm; const aCtrl: string ): TComponent;
begin
  if( aForm is TForm ) then
    Result := aForm.FindComponent( aCtrl )
  else
    Result := nil;
end;

function GetFormControlEnabled( aForm: TForm; const aCtrl: string ): Boolean;
var
  pCtrl: TComponent;
begin
  pCtrl := GetFormComponent( aForm, aCtrl );
  if ( pCtrl is TControl ) then
    Result := TControl(pCtrl).Enabled
  else
    Result := Boolean( GetOrdProp(pCtrl, 'Enabled') );
end;

function GetFormControlVisible( aForm: TForm; const aCtrl: string ): Boolean;
var
  pCtrl: TComponent;
begin
  pCtrl := GetFormComponent( aForm, aCtrl );
  if ( pCtrl is TControl ) then
    Result := TControl(pCtrl).Visible
  else
    Result := Boolean( GetOrdProp(pCtrl, 'Visible') );
end;

function GetTextPropInfo( const aCtrl: TObject ): PPropInfo;
begin
  if( aCtrl is THValComboBox ) then
    Result := GetPropInfo( aCtrl, 'Value' )
  else
    if( Assigned(aCtrl) ) then
    begin
      Result := GetPropInfo( aCtrl, 'Text' );
      if( not(Assigned(Result)) ) then
        Result := GetPropInfo( aCtrl, 'Caption' );
    end
    else
      Result := nil;
end;

function GetComponentText( pCtrl: TComponent ): string;
var
  propInfo: PPropInfo;
begin
  if( pCtrl is TCheckBox ) then
  begin
    if( TCheckBox(pCtrl).Checked ) then
      Result := 'X'
    else
      Result := '-';
  end
  else
  begin
    propInfo := GetTextPropInfo( pCtrl );
    if( Assigned(propInfo) ) then
      Result := GetStrProp( pCtrl, propInfo );
  end;
end;

procedure SetFormControlEnabled( aForm: TForm; const aCtrl: string; value: Boolean );
var
  pCtrl: TComponent;
begin
  pCtrl := GetFormComponent( aForm, aCtrl );
  if ( pCtrl is TControl ) then
    TControl(pCtrl).Enabled := value
  else
    SetOrdProp( pCtrl, 'Enabled', Integer(value) );
end;

procedure SetFormControlVisible( aForm: TForm; const aCtrl: string; value: Boolean );
var
  pCtrl: TComponent;
begin
  pCtrl := GetFormComponent( aForm, aCtrl );
  if ( pCtrl is TControl ) then
    TControl(pCtrl).Visible := value
  else
    SetOrdProp( pCtrl, 'Visible', Integer(value) );
end;


procedure SetComponentText( pCtrl: TComponent; const aText: string );
var
  propInfo: PPropInfo;
begin
  if( pCtrl is TCheckBox ) then
  begin
    if( aText = 'X' ) then
      TCheckBox(pCtrl).Checked := True
    else if( aText = '-' ) then
      TCheckBox(pCtrl).Checked := False
    else if( TCheckBox(pCtrl).AllowGrayed ) then
      TCheckBox(pCtrl).State := cbGrayed;
  end
  else
  begin
    propInfo := GetTextPropInfo( pCtrl );
    if( Assigned(propInfo) ) then
      SetStrProp( pCtrl, propInfo, aText );
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : O. Grosclaude
Créé le ...... : 10/06/2005
Modifié le ... :   /  /
Description .. : Fonctions GetControl, GetControlEnabled, GetControlVisible,
Suite ........ : GetControlEnabled, GetControlText, GetControlProperty
Suite ........ : ainsi que leurs équivalents SetControlText, ...
Suite ........ : avec comme paramètre une TOM, une TOF ou une TForm
Mots clefs ... :
*****************************************************************}

function GetControl( pObject: TObject; const pControlName: string;
                     ControlClass: TClass; AvecErreurSiNonTrouve: Boolean ): TControl;
var
  pObj: TObject;
begin
   if( pObject is TForm ) then
    pObj := GetFormComponent( TForm(pObject), pControlName )
  else if( pObject is TOM ) then
    pObj := TOM( pObject ).GetControl( pControlName )
  else if( pObject is TOF ) then
    pObj := TOF( pObject ).GetControl( pControlName, AvecErreurSiNonTrouve )
  else
    pObj := nil;
  if( pObj is ControlClass ) then
    Result := TControl( pObj )
  else
    Result := nil;
end;

function GetControlEnabled( pObject: TObject; const pControlName: string ): Boolean;
begin
   if( pObject is TForm ) then
    Result := GetFormControlEnabled( TForm(pObject), pControlName )
  else if( pObject is TOM ) then
    Result := TOM( pObject ).GetControlEnabled( pControlName )
  else if( pObject is TOF ) then
    Result := TOF( pObject ).GetControlEnabled( pControlName )
  else
    Result := False;
end;

function GetControlVisible( pObject: TObject; const pControlName: string ): Boolean;
begin
  if( pObject is TForm ) then
    Result := GetFormControlVisible( TForm(pObject), pControlName )
  else if( pObject is TOM ) then
    Result := TOM( pObject ).GetControlVisible( pControlName )
  else if( pObject is TOF ) then
    Result := TOF( pObject ).GetControlVisible( pControlName )
  else
    Result := False;
end;

function GetControlText( pObject: TObject; const pControlName: string ): string;
begin
  if( pObject is TForm ) then
    Result := GetComponentText( GetFormComponent(TForm(pObject), pControlName) )
  else if( pObject is TOM ) then
    Result := TOM( pObject ).GetControlText( pControlName )
  else if( pObject is TOF ) then
    Result := TOF( pObject ).GetControlText( pControlName )
  else
    Result := '';
end;

function GetControlProperty( pObject: TObject; const pControlName, aProperty: string ): Variant;
begin
  if( pObject is TOM ) then
    pObject := Tom(pObject).GetControl( pControlName )
  else if( pObject is TOF ) then
    pObject := TOF(pObject).GetControl( pControlName );
  if( pObject is TForm ) then
    Result := GetPropValue( GetFormComponent(TForm(pObject), pControlName), aProperty )
  else if( pObject is TComponent ) then
    Result := GetPropValue( pObject, aProperty )
  else
    Result := varNull;
end;

procedure SetControlEnabled( pObject: TObject; const pControlName: string; value: Boolean );
begin
  if( pObject is TForm ) then
    SetFormControlEnabled( TForm(pObject), pControlName, value )
  else if( pObject is TOM ) then
    TOM( pObject ).SetControlEnabled( pControlName, value )
  else if( pObject is TOF ) then
    TOF( pObject ).SetControlEnabled( pControlName, value );
end;

procedure SetControlVisible( pObject: TObject; const pControlName: string; value: Boolean );
begin
  if( pObject is TForm ) then
    SetFormControlVisible( TForm(pObject), pControlName, value )
  else if( pObject is TOM ) then
    TOM( pObject ).SetControlVisible( pControlName, value )
  else if( pObject is TOF ) then
    TOF( pObject ).SetControlVisible( pControlName, value );
end;

procedure SetControlText( pObject: TObject; const pControlName, aText: string );
begin
  if( pObject is TForm ) then
    SetComponentText( GetFormComponent(TForm(pObject), pControlName), aText )
  else if( pObject is TOM ) then
    TOM( pObject ).SetControlText( pControlName, aText )
  else if( pObject is TOF ) then
    TOF( pObject ).SetControlText( pControlName, aText );
end;

procedure SetControlProperty( pObject: TObject; const pControlName, aProperty: string; value: Variant );
begin
  if( pObject is TForm ) then
    SetPropValue( GetFormComponent(TForm(pObject), pControlName), aProperty, value )
  else if( pObject is TOM ) then
    TOM( pObject ).SetControlProperty( pControlName, aProperty, value )
  else if( pObject is TOF ) then
    TOF( pObject ).SetControlProperty( pControlName, aProperty, value );
end;

{***********A.G.L.***********************************************
Auteur  ...... : D. Carret
Créé le ...... : 21/11/2005
Modifié le ... :   /  /
Description .. : Dans une fiche MUL, retourne le numéro de la colonne de
Suite ........ : FListe dont le libellé vaut ColName.
Suite ........ : Retourne -1 si non trouvé.
Mots clefs ... : MUL;FLISTE;COLNAME;COLINDEX
*****************************************************************}
function FListeFieldNameToColIndex( F : TFMul ; ColName : string ) : integer;
var
  iCol : integer;
begin
  Result := -1; (* Non trouvé *)
{$IFNDEF EAGLCLIENT}
  for iCol := 0 to F.Fliste.Columns.Count - 1 do
    if F.Fliste.Columns[ iCol ].FieldName = ColName then
    begin
      Result := iCol;
      break;
    end;
{$ELSE}
  for iCol := 0 to F.Fliste.ColCount - 1 do
    if F.Fliste.FColnames [ iCol ] = ColName then
    begin
      Result := iCol;
      break;
    end;
{$ENDIF}
end;

end.
