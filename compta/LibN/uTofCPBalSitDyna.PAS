{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/06/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPBALSITDYNA ()
Mots clefs ... : TOF;CPBALSITDYNA
*****************************************************************}
Unit uTofCPBalSitDyna ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     eMul,
     MainEAgl,        // AGLLanceFiche
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_Main,         // AGLLanceFiche
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     Ent1,            // VH^.
     HMsgBox,
     ParamSoc,        // GetParamSocSecur
     CritEdt,         // ttypEcr
     uTob,            // Tob
     ED_TOOLS,        // InitMoveProgressForm
     UTOF ;

Type
  TOF_CPBALSITDYNA = Class (TOF)

    CBMensuelle     : TCheckBox;
    CBTrimestrielle : TCheckBox;
    CBSemestrielle  : TCheckBox;
    CBAnnuelle      : TCheckBox;

    procedure OnArgument (S : String ) ; override ;
    procedure OnLoad                   ; override ;
    procedure OnDisplay                ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnCancel                 ; override ;
    procedure OnClose                  ; override ;
  end;

procedure CPLanceFiche_CPBALSITDYNA;

Implementation

uses BalSit; // TBalSit

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/06/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure CPLanceFiche_CPBALSITDYNA;
begin
  AGLLanceFiche('CP', 'CPBALSITDYNA', '', '', '');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/06/2006
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPBALSITDYNA.OnArgument (S : String ) ;
begin
  Inherited ;

  CBMensuelle     := TCheckBox(GetControl('CBMENSUELLE', True));
  CBTrimestrielle := TCheckBox(GetControl('CBTRIMESTRIELLE', True));
  CBSemestrielle  := TCheckBox(GetControl('CBSEMESTRIELLE', True));
  CBAnnuelle      := TCheckBox(GetControl('CBANNUELLE', True));

  CBMensuelle.Checked     := GetParamSocSecur('SO_CPBDSMENSUELLE', False);
  CBTrimestrielle.Checked := GetParamSocSecur('SO_CPBDSTRIMESTRIELLE', False);
  CBSemestrielle.Checked  := GetParamSocSecur('SO_CPBDSSEMESTRIELLE', False);
  CBAnnuelle.Checked      := GetParamSocSecur('SO_CPBDSANNUELLE', False);

  Ecran.Caption := TraduireMemoire('Balances de situation dynamiques');
  UpdateCaption(Ecran);

end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/06/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPBALSITDYNA.OnLoad ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/06/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPBALSITDYNA.OnDisplay () ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/06/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPBALSITDYNA.OnUpdate ;
begin
  Inherited ;
  SetParamSoc('SO_CPBDSMENSUELLE', CBMensuelle.Checked);
  SetParamSoc('SO_CPBDSTRIMESTRIELLE', CBTrimestrielle.Checked);
  SetParamSoc('SO_CPBDSSEMESTRIELLE', CBSemestrielle.Checked);
  SetParamSoc('SO_CPBDSANNUELLE', CBAnnuelle.Checked);

  try
    InitMoveProgressForm(nil, 'Création des balances de situation dynamiques en cours.', 'Veuillez patienter...', 4, False, True);

    MoveCurProgressForm('Balance mensuelle');
    if CBMensuelle.Checked then
     CreationBDSDynamique('M');

    MoveCurProgressForm('Balance trimestrielle');
    if CBTrimestrielle.Checked then
      CreationBDSDynamique('T');

    MoveCurProgressForm('Balance semestrielle');
    if CBSemestrielle.Checked then
      CreationBDSDynamique('S');

    MoveCurProgressForm('Balance annuelle');
    if CBAnnuelle.Checked then
      CreationBDSDynamique('A');


  finally
    FiniMoveProgressForm;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPBALSITDYNA.OnCancel () ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPBALSITDYNA.OnClose ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////


Initialization
  registerclasses ( [ TOF_CPBALSITDYNA ] ) ;
end.
