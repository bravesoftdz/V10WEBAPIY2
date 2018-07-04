{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/03/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : CREVPLAN (CREVPLAN)
Mots clefs ... : TOM;CREVPLAN
*****************************************************************}
Unit CREVPLAN_TOM ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
{$IFDEF EAGLCLIENT}
     eMul,
     eFiche,        // TFFiche
     MainEagl,      // AGLLanceFiche
{$ELSE}
     db,
     HDB,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fiche,         // TFFiche
     FE_main,       // AGLLanceFiche
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     HTB97,
     HQry,          // RecupWhereCritere
     UTOM,
     UTob ;

Type
  TOM_CREVPLAN = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    procedure OnLoadAlerte               ; override ;

  private
    FBoLoading : Boolean; // Chargement du record en cours

    procedure ChargeActivation;
    procedure SauveActivation;

    function  TrouveDataType( vStValue : string ) : string;

  public
    Pages2  : THPageControl2;
    BEfface : TToolBarbutton97;

    Z_C1    : THValComboBox;
    ZO1     : THValComboBox;
    ZV1     : THEdit;
    ZG1     : THValComboBox;

    Z_C2    : THValComboBox;
    ZO2     : THValComboBox;
    ZV2     : THEdit;
    ZG2     : THValComboBox;

    Z_C3    : THValComboBox;
    ZO3     : THValComboBox;
    ZV3     : THEdit;

    procedure OnClickBEfface       ( Sender : TObject );

    procedure OnChangeZC           ( Sender : TObject );
    procedure OnChangeZxx          ( Sender : TObject );

  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPVersion,
  {$ENDIF MODENT1}
  Ent1; // EstSpecif

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPLAN.OnArgument ( S: String ) ;
begin
  Inherited ;
  Pages2  := THPageControl2(GetControl('PAGES2', True));
  BEfface := TToolbarButton97(GetControl('BEFFACE', True));

  Z_C1    := THValComboBox(GetControl('Z_C1', True));
  ZO1     := THValComboBox(GetControl('ZO1', True));
  ZV1     := THEdit(GetControl('ZV1', True));
  ZG1     := THValComboBox(GetControl('ZG1', True));

  Z_C2    := THValComboBox(GetControl('Z_C2', True));
  ZO2     := THValComboBox(GetControl('ZO2', True));
  ZV2     := THEdit(GetControl('ZV2', True));
  ZG2     := THValComboBox(GetControl('ZG2', True));

  Z_C3    := THValComboBox(GetControl('Z_C3', True));
  ZO3     := THValComboBox(GetControl('ZO3', True));
  ZV3     := THEdit(GetControl('ZV3', True));

  // GCO - 02/10/2007 - FQ 21568
  Z_C2.DataType := 'CREVCDTACTIVATION';
  Z_C3.DataType := Z_C2.DataType;

  BEfface.OnClick := OnClickBEfface;
  Z_C1.OnChange   := OnChangeZC;
  ZO1.OnChange    := OnChangeZxx;
  ZV1.OnChange    := OnChangeZxx;
  ZG1.OnChange    := OnChangeZxx;

  Z_C2.OnChange   := OnChangeZC;
  ZO2.OnChange    := OnChangeZxx;
  ZV2.OnChange    := OnChangeZxx;
  ZG2.OnChange    := OnChangeZxx;

  Z_C3.OnChange   := OnChangeZC;
  ZO3.OnChange    := OnChangeZxx;
  ZV3.OnChange    := OnChangeZxx;

  if not (CtxPCl in V_Pgi.PgiContexte) then
    SetControlProperty('TABACTIVATION', 'TABVISIBLE', False);
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPLAN.OnClose ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPLAN.OnLoadRecord ;
begin
  Inherited ;
  FBoLoading := True;
  ChargeActivation;
  FBoLoading := False;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPLAN.OnClickBEfface( Sender : TObject );
begin
  Z_C1.ItemIndex := -1;
  ZO1.ItemIndex := -1;
  ZV1.Text := '';
  ZG1.ItemIndex := -1;

  Z_C2.ItemIndex := -1;
  ZO2.ItemIndex := -1;
  ZV2.Text := '';
  ZG2.ItemIndex := -1;

  Z_C3.ItemIndex := -1;
  ZO3.ItemIndex := -1;
  ZV3.Text := '';
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/05/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPLAN.OnNewRecord ;
begin
  Inherited ;
  if ctxStandard in V_PGI.PGIContexte then
  begin
    if EstSpecif('51502') then
      SetField('CPR_PREDEFINI', 'CEG')
    else
      SetField('CPR_PREDEFINI', 'STD');
  end
  else
    SetField('CPR_PREDEFINI', 'STD')
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/05/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPLAN.OnDeleteRecord ;
begin
  Inherited ;

  if (not EstSpecif('51502')) and (GetField('CPR_PREDEFINI') = 'CEG') then
  begin
    PgiInfo('Vous ne pouvez pas supprimer un plan de révision fournit par CEGID.', Ecran.Caption);
    LastError := -1;
    Exit;
  end; 

  // GCO - 07/06/2007 - FQ 20565
  if ExisteSQL('SELECT CPC_CODECYCLE FROM CREVPARAMCYCLE WHERE ' +
               'CPC_PLANASSOCIE LIKE "%' + GetField('CPR_PLANREVISION') + ';%"') then
  begin
    PgiInfo('Vous ne pouvez pas supprimer un plan de révision lorsqu’il est renseigné dans un cycle.', Ecran.Caption);  
    LastError := -1;
    Exit;
  end;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/05/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPLAN.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_CREVPLAN.OnUpdateRecord ;
begin
  Inherited ;
  SauveActivation;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPLAN.ChargeActivation;
var lSt : string;
    lStTemp : string;
begin
  lSt := GetField('CPR_ACTIVATION');

  lSt := FindEtReplace(lSt, '(', '', True);
  lSt := FindEtReplace(lSt, ')', '', True);

  lStTemp      := ReadTokenPipe(lSt, '|');
  Z_C1.Value   := lStTemp;

  lStTemp      := ReadTokenPipe(lSt, '|');
  ZO1.Value    := lStTemp;

  lStTemp      := ReadTokenPipe(lSt, '|');
  ZV1.Text     := lStTemp;
  ZV1.DataType := TrouveDataType( Z_C1.Value );

  lStTemp      := ReadTokenPipe(lSt, '|');
  ZG1.Value    := lStTemp;

  lStTemp      := ReadTokenPipe(lSt, '|');
  Z_C2.Value   := lStTemp;

  lStTemp      := ReadTokenPipe(lSt, '|');
  ZO2.Value    := lStTemp;

  lStTemp      := ReadTokenPipe(lSt, '|');
  ZV2.Text     := lStTemp;
  ZV2.DataType := TrouveDataType( Z_C2.Value );

  lStTemp      := ReadTokenPipe(lSt, '|');
  ZG2.Value    := lStTemp;

  lStTemp      := ReadTokenPipe(lSt, '|');
  Z_C3.Value   := lStTemp;

  lStTemp      := ReadTokenPipe(lSt, '|');
  ZO3.Value    := lStTemp;

  lStTemp      := ReadTokenPipe(lSt, '|');
  ZV3.Text     := lStTemp;
  ZV3.DataType := TrouveDataType( Z_C3.Value );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPLAN.SauveActivation;
var lSt : string;
begin
  lSt := '';

  if (Z_C1.Value <> '') and (ZO1.Value <> '') and (ZV1.Text <> '') then
  begin
    lSt := lSt + Z_C1.Value + '|' + ZO1.Value + '|' + ZV1.Text + '|' ;

    if (ZG1.Value <> '') and (Z_C2.Value <> '') and
       (ZO2.Value <> '') and (ZV2.Text <> '') then
    begin
      lSt := lSt + ZG1.Value + '|' + Z_C2.Value + '|' + ZO2.Value + '|' + ZV2.Text + '|';

      if (ZG2.Value <> '') and (Z_C3.Value <> '') and
         (ZO3.Value <> '') and (ZV3.Text <> '') then
        lSt := lSt + ZG2.Value + '|' + Z_C3.Value + '|' + ZO3.Value + '|' + ZV3.Text;
    end;
  end;

  SetField('CPR_ACTIVATION', lSt);

  lSt := RecupWhereCritere( Pages2 );
  SetField('CPR_ACTIVATIONSQL', Copy(lSt, 7, Length(lSt)));
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOM_CREVPLAN.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_CREVPLAN.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_CREVPLAN.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_CREVPLAN.OnLoadAlerte;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPLAN.OnChangeZC(Sender: TObject);
var lSt  : string;
begin
  if FBoLoading then Exit;
  if not (DS.State in [dsEdit,dsInsert]) then DS.Edit ;

  lSt := TrouveDataType( THValComBoBox(Sender).Value );

  if THValComBoBox(Sender).Name = 'Z_C1' then
  begin
    ZV1.Text := '';
    ZV1.DataType  := lSt;
  end;

  if THValComBoBox(Sender).Name = 'Z_C2' then
  begin
    ZV2.Text := '';
    ZV2.DataType  := lSt;
  end;

  if THValComBoBox(Sender).Name = 'Z_C3' then
  begin
    ZV3.Text := '';
    ZV3.DataType  := lSt;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOM_CREVPLAN.OnChangeZxx(Sender: TObject);
begin
  if FBoLoading then Exit;
  if not (DS.State in [dsEdit,dsInsert]) then
    DS.Edit;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_CREVPLAN.TrouveDataType(vStValue : string): string;
begin
  Result := '';
  if vStValue = '' then Exit;
  Result := GetColonneSQL('COMMUN', 'CO_LIBRE', 'CO_TYPE = "CDA" AND ' +
                          'CO_ABREGE = "' + vStValue + '"');
end;

////////////////////////////////////////////////////////////////////////////////

Initialization
  registerclasses ( [ TOM_CREVPLAN ] ) ; 
end.
