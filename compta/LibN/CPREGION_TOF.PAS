{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 03/03/2003
Modifié le ... : 25/03/2003
Description .. : Source TOF de la FICHE : CPREGION ()
Mots clefs ... : TOF;CPREGION
*****************************************************************}
Unit CPREGION_TOF ;

Interface

Uses
     Sysutils,
     Classes,
{$IFDEF EAGLCLIENT}
     maineagl,
{$ELSE}
     FE_main, HDB,
{$ENDIF}
     HCtrls,
     HMsgBox,
     LookUp, // pour le Lookuplist
     UTOB,
     UTOF,
     UTOFGRILLEFILTRE ;
{$IFDEF EAGLCLIENT}
Procedure PaysRegion ( X_PAYS : THValComboBox ; X_DIVTERRIT : THEdit ; FromReg : boolean ) ;
{$ELSE}
Procedure PaysRegion ( X_PAYS : THDBValComboBox ; X_DIVTERRIT : THDBEdit ; FromReg : boolean ) ;
{$ENDIF}
Function FicheRegion ( vStPays, vStRegion : String ; Zoom : boolean ) : string ;
Function CPLanceFiche_Region( vStRange, vStLequel, vStArguments : String ) : String ;

Type
  TOF_CPREGION = Class (TOF_GRILLEFILTRE)
    // procedure TOF
    procedure OnArgument ( S : String )                        ; override ;
    procedure OnClose                                          ; override ;
    // Procedure virtuelle
    procedure OnLoadRecord ( vTOBLignes : TOB )                ; override ;
    function  OnUpdateRecord ( laTOB : TOB )       : Integer   ; override ;
    function  OnAfterUpdateRecord ( laTOB : TOB )  : Integer   ; override ;
    function  OnDeleteRecord ( laTOB : TOB )       : Integer   ; override ;
    procedure ParametrageFiche                                 ; override ;
    procedure ParametrageGrille ( vGrille : THGrid )           ; override ;
    function  GenererConditionPlus                 : String    ; override ;
    procedure OnCreateTOB ( vTOB : TOB )                       ; override ;
  end ;

Implementation

//==============================================================================
//==============================================================================

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 12/03/2003
Modifié le ... :   /  /    
Description .. : Traduction eAGL de la procédure de Region;pas
Mots clefs ... :
*****************************************************************}
{$IFDEF EAGLCLIENT}
Procedure PaysRegion ( X_PAYS : THValComboBox ; X_DIVTERRIT : THEdit ; FromReg : boolean ) ;
{$ELSE}
Procedure PaysRegion ( X_PAYS : THDBValComboBox ; X_DIVTERRIT : THDBEdit ; FromReg : boolean ) ;
{$ENDIF}
Var
  Pays,Region,StRegion,StPays : String ;
begin
  Pays   := X_PAYS.Value ;
  Region := X_DIVTERRIT.Text ;
  if Not FromReg then
    begin
    if Region='' then Exit ;
    if Pays='' then
      X_DIVTERRIT.Text := ''
    else
      if not ExisteSQL('SELECT * FROM REGION WHERE RG_PAYS="'+PAYS+'" AND RG_REGION="'+Region+'"')
        then X_DIVTERRIT.Text:='' ;
    end
  else
    begin
    if Pays<>''
      then StPays := 'WHERE RG_PAYS="' + PAYS + '"'
      else StPays := '' ;
    if Region<>''
      then StRegion := 'AND RG_REGION="' + Region + '"'
      else StRegion := '' ;
    if (Region<>'') And (Pays='') then
      begin
      StPays   := 'WHERE RG_REGION="'+Region+'"' ;
      StRegion := '' ;
      end ;
    if Not ExisteSQL('Select * from REGION '+StPays+' '+StRegion+'') then
      begin
      X_DIVTERRIT.Text:='' ;
      Exit ;
      end
    else
      begin
      LookUpList( X_DIVTERRIT,
                  'Sélection d''une région',
                  'REGION',
                  'RG_REGION',
                  'RG_LIBELLE',
                  'RG_PAYS = "' + Pays + '"',
                  'RG_REGION',
                  True,
                  -1
                  );
     end ;
    end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 28/02/2003
Modifié le ... :   /  /
Description .. : Fonction de lancement de la fiche YYRegion, conservé
Suite ........ : pour rétro-compatibilité
Mots clefs ... :
*****************************************************************}
function FicheRegion( vStPays, vStRegion : String ; Zoom : boolean ) : string ;
var lStArgs : String ;
begin
  Result    := '' ;
  // Lancement pour zoom = consultation
  if Zoom then
    lStArgs := 'ACTION=CONSULTATION'
  else
    begin
    if Blocage(['nrCloture'],False,'nrAucun') then Exit ;
    lStArgs := 'ACTION=MODIFICATION' ;
    end ;
  if Trim(vStPays) <> '' then
     lStArgs := lStArgs + ';' + vStPays ;
  if Trim(vStRegion) <> '' then
     lStArgs := lStArgs + ';' + vStRegion ;
  // ouverture fiche
  Result := CPLanceFiche_Region( '', '' , lStArgs ) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 28/02/2003
Modifié le ... :   /  /
Description .. : Fonction de lancement de la fiche YYREGION
Mots clefs ... :
*****************************************************************}
Function CPLanceFiche_Region( vStRange, vStLequel, vStArguments : String ) : String ;
begin
  Result := AGLLanceFiche( 'CP', 'CPREGION' , vStRange , vStLequel , vStArguments ) ;
end ;

// =============================================================================
// =============================================================================

procedure TOF_CPREGION.ParametrageFiche;
begin
  // Param de base
  TitreEcran      := 'Liste des régions' ;
  NomTable        := 'REGION' ;
  ListeChamps     := 'RG_REGION;RG_LIBELLE' ;
  CodeEtat        := '' ;
  NatureEtat      := '' ;
  ListeColsUniq   := '1';
  ChampTri        := 'RG_REGION' ;
end;

function TOF_CPREGION.OnUpdateRecord ( laTOB : TOB ) : Integer ;
begin
  Result := 0 ;
  if Trim(laTOB.GetValue('RG_REGION')) = '' then
    begin
    Result := 1 ;
    PGIBox('Vous devez renseigner un code.', Ecran.Caption ) ;
    end ;
  if Trim(laTOB.GetValue('RG_LIBELLE')) = '' then
    begin
    Result := 1 ;
    PGIBox('Vous devez renseigner un libellé.', Ecran.Caption ) ;
    end ;
end;

function TOF_CPREGION.OnAfterUpdateRecord ( laTOB : TOB ) : Integer ;
begin
  Result := 0 ;
  try
    if not ExisteSQL('SELECT * FROM REGION WHERE RG_PAYS="' + laTOB.GetValue('RG_PAYS') + '" '
                 + 'AND RG_REGION<>"' + laTOB.GetValue('RG_REGION') + '"' )
      then ExecuteSql('UPDATE PAYS SET PY_REGION="X" WHERE PY_PAYS="' + laTOB.GetValue('RG_PAYS') + '"') ;
  except
    Result := 1 ;
  end ;
end;

function TOF_CPREGION.OnDeleteRecord ( laTOB : TOB ) : Integer ;
begin
  Result := 0 ;
  try
    if ExisteSQL('SELECT * FROM REGION WHERE RG_PAYS="' + GetField('RG_PAYS') + '" '
                + 'AND RG_REGION<>"' + laTOB.GetValue('RG_REGION') + '"' )
      then ExecuteSql('UPDATE PAYS SET PY_REGION="-" WHERE PY_PAYS="' + laTOB.GetValue('RG_PAYS') + '"') ;
  except
    Result := 1 ;
  end ;
end;

function TOF_CPREGION.GenererConditionPlus: String;
begin
  Result := '' ;
end;

procedure TOF_CPREGION.ParametrageGrille(vGrille: THGrid);
begin
	// Paramétrage de la grille
  vGrille.ColLengths[1] := 17;          // Code 17 car maxi.
  vGrille.ColLengths[2] := 35;          // Lib 35 car maxi.
end;

procedure TOF_CPREGION.OnLoadRecord(vTOBLignes: TOB);
begin
  // Rien a faire
end;

procedure TOF_CPREGION.OnArgument(S: String);
Var lStPays : String ;
begin
  inherited;
  lStPays := '' ;
  ReadTokenSt(S); // Action=...
  if Trim(S) <> '' then
    lStPays := ReadTokenSt(S) ; // pays cible
  if lStPays <> '' then
    begin
    SetControlProperty('RG_PAYS', 'DATATYPE', 'TTPAYS') ;
    SetControlText('RG_PAYS', lStPays) ;
    SetControlEnabled('RG_PAYS', False) ;
    end
  else
    begin
    SetControlProperty('RG_PAYS', 'DATATYPE', 'TTPAYSREGIONS') ;
    THValComboBox(GetControl('RG_PAYS', True)).ItemIndex := 0 ;
    end ;

end;

procedure TOF_CPREGION.OnCreateTOB(vTOB: TOB);
begin
  vTOB.PutValue('RG_PAYS', GetControlText('RG_PAYS') ) ;
end;

procedure TOF_CPREGION.OnClose;
begin
  inherited;
end;

Initialization
  registerclasses ( [ TOF_CPREGION ] ) ;
end.
