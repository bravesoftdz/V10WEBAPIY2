{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/03/2003
Modifié le ... : 06/08/2003
Description .. : Source TOF de la FICHE : CPREGION ()
Suite ........ : JP 06/08/03 : Deux petites modifications dans VerifCodePostal pour 
                               corriger deux petits disfonctionnements en CWAS
Mots clefs ... : TOF;CPREGION
*****************************************************************}
Unit BTCODEPOSTAL_TOF ;

Interface

Uses
     Sysutils,
     Classes,
{$IFDEF EAGLCLIENT}
     maineagl,
{$ELSE}
     FE_main,db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     HCtrls,
     HMsgBox,
     HPanel,
     lookUp, // pour le lookUpList
     UTOB,
     UTOF,
     UTOFGRILLEFILTRE ;


Procedure OuvrePosTal (Inside : THPanel );
Function CPLanceFiche_CodePostal( vStRange, vStLequel, vStArguments : String ) : String ;
Procedure VerifCodePostal(T : TDataSet ; CP,V : THEDIT ; Code : boolean) ;

Type
  TOF_BTCODEPOSTAL = Class (TOF_GRILLEFILTRE)
    procedure OnLoadRecord ( vTOBLignes : TOB )                ; override ;
    function  OnUpdateRecord ( laTOB : TOB )       : Integer   ; override ;
    function  OnAfterUpdateRecord ( laTOB : TOB )  : Integer   ; override ;
    function  OnDeleteRecord ( laTOB : TOB )       : Integer   ; override ;
    procedure ParametrageFiche                                 ; override ;
    procedure ParametrageGrille ( vGrille : THGrid )           ; override ;
    procedure OnCreateTOB ( vTOB : TOB )                       ; override ;
    function 	GenererConditionPlus                 : String    ; override ;

    // Paramétrage dans la fille
  end ;

Implementation

//==============================================================================
//==============================================================================

Procedure VerifCodePostal(T : TDataSet ; CP,V : THEDIT ; Code : boolean) ;
var okok, lBoTrouve          : boolean ;
    Q                        : TQuery ;
    lStColBase, lStColRetour : String ;
    lStWhere, lStCodeTest    : String ;
    lCtrlRetour              : THEdit ;
begin
  okok      := False ;
  // Param base de recherche
  if Code then
    begin
    lCtrlRetour   := V ;
    lStColRetour  := 'O_VILLE' ;
    lStWhere      := 'O_CODEPOSTAL="' + CP.text + '"' ;
    lStColBase    := 'O_CODEPOSTAL' ;
    lStCodeTest   := CP.Text ;
    end
  else
    begin
    lCtrlRetour   := CP ;
    lStColRetour  := 'O_CODEPOSTAL' ;
    lStWhere      := 'O_VILLE="' + V.Text + '"' ;
    lStColBase    := 'O_VILLE' ;
    lStCodeTest   := V.Text ;
    end ;
  // Test codes saisis
  Q := OpenSQL('SELECT * FROM CODEPOST WHERE ' + lStWhere ,TRUE) ;
  {$IFDEF EAGLCLIENT}
  {JP 06/08/03 : RecordsCount(Q), en EAGL du moins, ne fonctionne pas}
  lBoTrouve := Q.Detail.Count >= 1;
  {$ELSE}
  lBoTrouve := RecordsCount( Q ) >= 1 ;
  {$ENDIF}
  if Not Q.EOF then
    begin
    if Q.findField(lStColRetour).AsString = lCtrlRetour.Text then
      okok := TRUE
    else
      begin
      Q.next ;

      {JP 06/08/03 : Ajout du test "not Q.EOF" pour éviter une vialation d'accès
                     lorsqu'il n'y a qu'un seul enregistrement dans le query}
      if not Q.EOF and (Q.findField(lStColBase).AsString <> lStCodeTest) then
        begin
        Q.prior ;
        okok:=TRUE ;
        if T.State=dsBrowse
          then T.edit ;
        lCtrlRetour.Text := Q.findField(lStColRetour).AsString ;
        end ;
      end ;
    end ;
  Ferme(Q) ;
  // Ouverture lookup si tests non satisfaisants
  if Not okok then
    begin
    // Recherche sur tous les enregistrements si pas de correspondant au code saisi
    if not lBoTrouve
      then lStWhere := '' ;
    // Recherche code postal
    LookUpList( CP,
                'Sélection d''un code postal',
                'CODEPOST',
                'O_CODEPOSTAL',
                'O_VILLE',
                lStWhere,
                'O_CODEPOSTAL',
                True,
                -1
                );
    // MAJ controle V(ille)
    if CP.Text = ''
      then V.text := ''
      else V.text := RechDom('ICODEPOSTAL',CP.Text,False) ;
    end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 28/02/2003
Modifié le ... :   /  /
Description .. : Fonction de lancement de la fiche CPCODEPOSTAL
Suite ........ : , conservé pour rétro-compatibilité
Mots clefs ... :
*****************************************************************}
Procedure OuvrePosTal (Inside : THPanel );
var lStArgs : String ;
begin
  if Blocage(['nrCloture'],False,'nrAucun') then Exit ;
  // Lancement pour zoom = consultation
  lStArgs := 'ACTION=MODIFICATION' ;
  // ouverture fiche
  CPLanceFiche_CodePostal( '', '' , lStArgs ) ;              
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 28/02/2003
Modifié le ... :   /  /
Description .. : Fonction de lancement de la fiche CPCODEPOSTAL
Mots clefs ... :
*****************************************************************}
Function CPLanceFiche_CodePostal( vStRange, vStLequel, vStArguments : String ) : String ;
begin
  Result := AGLLanceFiche( 'BTP', 'BTCODEPOSTAL' , vStRange , vStLequel , vStArguments ) ;
end ;

// =============================================================================
// =============================================================================

procedure TOF_BTCODEPOSTAL.ParametrageFiche;
begin
  // Param de base
  TitreEcran      := 'Liste des codes postaux' ;
  NomTable        := 'CODEPOST' ;
  ChampTri        := 'O_CODEPOSTAL' ;
  ListeChamps     := 'O_CODEPOSTAL;O_VILLE;O_PAYS' ;
  CodeEtat        := '' ;
  NatureEtat      := '' ;
  ListeColsUniq   := '1'
end;

function TOF_BTCODEPOSTAL.OnUpdateRecord ( laTOB : TOB ) : Integer ;
begin
  Result := 0 ;
  if Trim(laTOB.GetValue('O_CODEPOSTAL')) = '' then
    begin
    Result := 1 ;
    PGIBox('Vous devez renseigner un code postal.', Ecran.Caption ) ;
    end ;
  if Trim(laTOB.GetValue('O_VILLE')) = '' then
    begin
    Result := 1 ;
    PGIBox('Vous devez renseigner une ville.', Ecran.Caption ) ;
    end ;
end;

function TOF_BTCODEPOSTAL.OnAfterUpdateRecord ( laTOB : TOB ) : Integer ;
begin
  Result := 0 ;
end;

function TOF_BTCODEPOSTAL.OnDeleteRecord ( laTOB : TOB ) : Integer ;
begin
  Result := 0 ;
end;

procedure TOF_BTCODEPOSTAL.ParametrageGrille(vGrille: THGrid);
begin
  // 3 colonnes
  vGrille.ColCount := 4 ;

  // titres
  vGrille.Cells[1,0] := 'Code postal' ;
  vGrille.Cells[2,0] := 'Ville' ;
  vGrille.Cells[3,0] := 'Pays' ;

	// Paramétrage de la grille
  vGrille.ColLengths[1] := 9;         // Code 17 car maxi.
	vGrille.ColLengths[2] := 35;         // Lib 35 car maxi.
	vGrille.ColLengths[3] := 3;          // Pays 35 car maxi.

  // Colonne 3 (PAYS) non modifiable
  vGrille.ColEditables[3] := False ;

end;

procedure TOF_BTCODEPOSTAL.OnCreateTOB(vTOB: TOB);
begin
  // Init du pays
  vTOB.PutValue('O_PAYS', 'FRA' ) ;
end;

function TOF_BTCODEPOSTAL.GenererConditionPlus: String;
begin
  Result := '' ;
end;

procedure TOF_BTCODEPOSTAL.OnLoadRecord(vTOBLignes: TOB);
begin
  // rien a faire
end;

Initialization
  registerclasses ( [ TOF_BTCODEPOSTAL ] ) ;
end.
