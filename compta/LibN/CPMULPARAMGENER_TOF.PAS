{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 09/12/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPMULPARAMGENER ()
Mots clefs ... : TOF;CPMULPARAMGENER
*****************************************************************}
unit CPMULPARAMGENER_TOF;

interface

uses StdCtrls,
  {$IFDEF VER150} Variants, {$ENDIF}
  Controls,
  Classes,
  Dialogs,
  {$IFDEF EAGLCLIENT}
  eMul,
  uTob,
  MaineAGL,
  {$ELSE}
  db,
  {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
  mul,
  FE_Main,
  HDB,
  dbGrids,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  HQry,
  CPARAMGENER_TOM;

procedure CPLanceFiche_MulParamGener(vStRange, vStLequel, vStArgs: string);

type
  TOF_CPMULPARAMGENER = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
  private
    // Composants
    Q: THQuery;
    {$IFDEF EAGLCLIENT}
    FListe: THGrid;
    {$ELSE}
    FListe: THDBGrid;
    {$ENDIF}
    Pages: TPageControl;
    // Variables
    Action: string; // Stocke XXX de l'argument "ACTION=XXX"

    CPG_FLUXENCADECA : THValComboBox ;
    // Evènement de la fiche
    procedure FListeDblClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BSupprimerClick(Sender: TObject);
    procedure CPG_FLUXENCADECAChange (Sender: TObject) ;

    // Initialisation de la fiche
    procedure AnalyseParametre(vStArg: string);
    procedure InitComposants;
    procedure InitEvenements;
    procedure InitEcran;
  end;

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  uLibEncaDeca, Ent1, UtilPGI ;

//==============================================================================
procedure CPLanceFiche_MulParamGener(vStRange, vStLequel, vStArgs: string);
begin
  AGLLanceFiche('CP', 'CPMULPARAMGENER', vStRange, vStLequel, vStArgs);
end;
//==============================================================================

procedure TOF_CPMULPARAMGENER.OnNew;
begin
  inherited;
end;

procedure TOF_CPMULPARAMGENER.OnDelete;
begin
  inherited;
end;

procedure TOF_CPMULPARAMGENER.OnUpdate;
begin
  inherited;
end;

procedure TOF_CPMULPARAMGENER.OnLoad;
begin
  inherited;
end;

procedure TOF_CPMULPARAMGENER.OnArgument(S: string);
begin
  inherited;
  if VH^.PaysLocalisation=CodeISOES then
     Action:='MODIF' ; //XVI 24/02/2005
  // Analyse des paramètres
  AnalyseParametre(S);

  // Initialisation de l'interface
  InitComposants;
  InitEvenements;
  InitEcran;

end;

procedure TOF_CPMULPARAMGENER.OnClose;
begin
  inherited;
end;

procedure TOF_CPMULPARAMGENER.OnDisplay();
begin
  inherited;
end;

procedure TOF_CPMULPARAMGENER.OnCancel();
begin
  inherited;
end;

procedure TOF_CPMULPARAMGENER.BInsertClick(Sender: TObject);
var LAction : String ;
begin
  LAction:='ACTION=' ;
  if VH^.PaysLocalisation=CodeISOES then
     LACtion:=LAction+Action
  else
     LACtion:=LAction+'CREATION' ;
  CPLanceFiche_ParamGener( '', '', LACtion) ; //XVI 24/02/2005
end;

procedure TOF_CPMULPARAMGENER.AnalyseParametre(vStArg: string);
var
  lStChaine: string;
begin

  while vStArg <> '' do
  begin
    lStChaine := ReadTokenSt(vStArg);

    // Stocke le XXX de "ACTION=XXX" dans la variable Action
    if Pos('ACTION=', Uppercase(lStChaine)) > 0 then
      Action := Copy(lStChaine, 8, length(lStChaine));
  end;

end;

procedure TOF_CPMULPARAMGENER.FListeDblClick(Sender: TObject);
begin
  //FG 14858 : SG6 02/11/2004
  if VarToStr(GetField('CPG_CODE')) <> '' then CPLanceFiche_ParamGener('', GetField('CPG_CODE'), 'ACTION=' + Action);
end;

procedure TOF_CPMULPARAMGENER.InitEcran;
begin
  if VH^.PaysLocalisation=CodeISOES then //XVI 24/02/2005
  begin
    SetControlText('CPG_FLUXENCADECA','') ;
    SetControlText('CPG_TYPEENCADECA','') ;
  End ;

  // Cacher l'onglet "compléments"
  TTabSheet(GetControl('PCOMPLEMENT', True)).TabVisible := False;

  // Faire apparaître le bouton de création si besoin
  SetControlVisible('BINSERT', (Action <> 'CONSULTATION'));

  SetControlVisible('BSUPPRIMER', (VH^.PaysLocalisation<>CodeISOES)) ;  //XVI 24/02/2005 suppression depuis la fiche, suppression en cascade....
end;

procedure TOF_CPMULPARAMGENER.InitComposants;
begin
  Q := THQuery(GetControl('Q', True));
  {$IFDEF EAGLCLIENT}
  FListe := THGrid(GetControl('FListe', True));
  {$ELSE}
  FListe := THDBGrid(GetControl('FListe', True));
  {$ENDIF}
  Pages := TPageControl(GetControl('Pages', True));
  CPG_FLUXENCADECA:=nil ;
  if VH^.PaysLocalisation=CodeISOES then //XVI 24/02/2005
     CPG_FLUXENCADECA := THValComboBox(      GetControl('CPG_FLUXENCADECA',          True) ) ;
end;

procedure TOF_CPMULPARAMGENER.InitEvenements;
begin
  // Double click sur la liste
  FListe.OnDblClick := FListeDblClick;

  // Bouton de création
  TButton(GetControl('BINSERT', True)).OnClick := BInsertClick;
  TButton(GetControl('BSUPPRIMER', True)).OnClick := BSupprimerClick;
  if (VH^.PaysLocalisation=CodeISOES) and (Assigned(CPG_FLUXENCADECA)) then //XVI 24/02/2005
     CPG_FLUXENCADECA.OnChange          := CPG_FLUXENCADECAChange ;
end;

procedure TOF_CPMULPARAMGENER.BSupprimerClick(Sender: TObject);
begin

  // Le résultat de la recherche est-il vide ?
  {$IFDEF EAGLCLIENT}
  if (Q.TQ.Detail.Count = 0) then
    {$ELSE}
  if (Q.Eof and Q.Bof) then
    {$ENDIF}
  begin
    PGIInfo('Le résultat de la recherche est vide. Veuillez relancer une recherche.', Ecran.Caption);
    exit;
  end;

  // Confirmation
  if PGIAsk('Confirmez-vous la suppression du scénario sélectionné : '
    + #10#13 + GetField('CPG_CODE') + ' - ' + GetField('CPG_LIBELLE'),
    Ecran.Caption) <> mrYes then Exit;

  // Supppression
  if executeSql('DELETE FROM CPARAMGENER WHERE CPG_CODE="' + GetField('CPG_CODE') + '"') <> 1
    then PGIInfo('Le scénario sélectionné n''a pas pu être supprimé.', Ecran.Caption);

  // On relnce la recherche
  TFMul(Ecran).BChercheClick(nil);

end;

procedure TOF_CPMULPARAMGENER.CPG_FLUXENCADECAChange (Sender: TObject) ;
Begin
  //on ne prend que les type d'enca/Déca que corresponent....
  SetControlProperty('CPG_TYPEENCADECA','PLUS',ConditionTypeEncaDeca(CPG_FLUXENCADECA.Value)) ;
  SetControltext('CPG_TYPEENCADECA','') ;
End ;

initialization
  registerclasses([TOF_CPMULPARAMGENER]);
end.

