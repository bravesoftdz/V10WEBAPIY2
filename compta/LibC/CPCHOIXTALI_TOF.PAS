{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 20/03/2002
Modifié le ... : 29/05/2002
Description .. : Source TOF de la FICHE : CPCHOIXTALI ()
Mots clefs ... : TOF;CPCHOIXTALI
*****************************************************************}
Unit CPCHOIXTALI_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     Maineagl,
{$ELSE}
     DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     Fe_main,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     Ent1,
     HMsgBox,
     UTOF,
     Utob,
     utilPGI, // EstTablePartagee / openselect...
     NATCPTE_TOM
     ;


procedure CPLanceFiche_ChoixtableLibre;
{$IFDEF EAGLCLIENT}
procedure AGLFicheNatCpte(Q : TQuery ; Code,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
{$ENDIF}
Type
  TOF_CPCHOIXTALI = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
  private
    ComboEntite, ComboTable : THValComboBox;
    FDossier : string ;
    procedure ChargeComboEntite;
    procedure SelectionEntite (Sender : TObject);
    function  GenererCondition: String;
  end ;

Implementation

// Entités sélectionnables
Const EntiteQte : Integer = 9;
Const EntiteListe : Array[0..8] of String =
                     ('G;Comptes généraux;',
                      'T;Comptes auxiliaires;',
                      'B;Comptes budgétaires;',
                      'D;Sections budgétaires;',
                      'S;Sections analytiques;',
                      'E;Ecritures comptables;',
                      'A;Ecritures analytiques;',
                      'U;Ecritures budgétaires;',
                      'I;Immobilisations;'
                      );

//==============================================================================
procedure CPLanceFiche_ChoixTableLibre;
begin
	AGLLanceFiche('CP','CPCHOIXTALI','','','')
end;

{$IFDEF EAGLCLIENT}
procedure AGLFicheNatCpte(Q : TQuery ; Code,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
Var
	Param : String;
BEGIN
	Case Comment of
		taCreat,taCreatOne	: Param := 'ACTION=CREATION';
    taModif							: Param := 'ACTION=MODIFICATION';
    else
    	Param := 'ACTION=CONSULTATION';
  	end;
  Param := Param + ';' + Code + ';' + Compte;

	CPLanceFiche_SaisieTableLibre(Param);
END;
{$ENDIF}

//==============================================================================

procedure TOF_CPCHOIXTALI.OnUpdate ;
begin
  Inherited ;
//  AGLFicheNatCpte(Nil,ComboTable.value,'',taCreat,0) ;
	CPLanceFiche_SaisieTableLibre('ACTION=CREATION;'+ComboTable.Value+';');
	ChargeTableLibre ;
end ;

procedure TOF_CPCHOIXTALI.OnLoad ;
begin
  Inherited ;
  if ComboEntite.Values.Count>0	Then
  	ComboEntite.Value := ComboEntite.Values[0];
end ;

procedure TOF_CPCHOIXTALI.OnArgument (S : String ) ;
begin
  Inherited ;

  // Gestion du multi-dossier
  if EstTablePartagee( 'NATCPTE' )
    then FDossier := TableToBase( 'NATCPTE' )
    else FDossier := '' ;

  // Récupération des contrôles importants
  ComboEntite := THValComboBox(GetControl('ChoixEntite'));
  ComboTable := THValComboBox(GetControl('ChoixTable'));
  // Réaffectation des évènements
  ComboEntite.OnChange := SelectionEntite;
  // Remplissage de la combo box
  ChargeComboEntite;
end ;

procedure TOF_CPCHOIXTALI.ChargeComboEntite;
// Modifier aussi TOF_CPTABLIBRELIB.ChargeCombos
Var i : Integer ;
    St : String ;
    it,va : String ;
Begin
  // init de la ComboBox
  ComboEntite.Values.Clear ;
  ComboEntite.Items.Clear ;
  // Remplissage de la combo
  for i:=0 to EntiteQte-1 do
    BEGIN
    St:=EntiteListe[i] ;
    va:=ReadTokenSt(St) ;
    it:=TraduireMemoire(ReadTokenSt(St)) ;
    if ((EstSerie(S5)) or (EstSerie(S3))) then
      BEGIN
      // Pas de table libres sur écritures analytiques et écritures budgétaires en S3 et S5
      if (va='A') or (va='U') then Continue;
      // Pas de tables libres sur comptes budgétaires et sections budgétaires en S3
      if EstSerie(S3) then
        begin
        if ((va='B') or (va='D')) then Continue ;
        // Pas de tables libres sur les sections analytiques si pas de séria analytiques
        //if EstComptaSansAna and (va='S') then Continue;
        end;
      END ;
    ComboEntite.Values.Add(va) ;
    ComboEntite.Items.Add(it) ;
    END ;
end;

procedure TOF_CPCHOIXTALI.SelectionEntite(Sender: TObject);
var Q : TQuery;
begin
  // Remplissage de la combo des Tables
  ComboTable.Values.Clear ;
  ComboTable.Items.Clear ;
  Q := OpenSelect('Select CC_CODE, CC_LIBELLE From CHOIXCOD WHERE '
  	       	+ GenererCondition + ' ORDER BY CC_CODE', FDossier);
  while (not Q.Eof) do
     BEGIN
     ComboTable.Values.Add(Q.FindField('CC_CODE').AsString) ;
     ComboTable.Items.Add(Q.FindField('CC_LIBELLE').AsString) ;
     Q.next;
     END ;
  Ferme(Q) ;
  // Sélection du 1er élément
  if ComboTable.Values.Count > 0 then
    ComboTable.Value := ComboTable.Values[0] ;
end;

function TOF_CPCHOIXTALI.GenererCondition: String;
begin
  // Par défaut
  Result := 'CC_TYPE="NAT" And CC_CODE Like "' + ComboEntite.Value + '%"';
  // Limitations S5, S3
  if (ComboEntite.Value = 'E') then
    BEGIN
    If EstSerie(S3) Then Result := 'CC_TYPE="NAT" And CC_CODE Like"'+ComboEntite.Value+'%" AND CC_CODE<="'+ComboEntite.Value+'00" '
      Else If EstSerie(S5) Then Result := 'CC_TYPE="NAT" And CC_CODE Like"'+ComboEntite.Value+'%" AND CC_CODE<="'+ComboEntite.Value+'01" '
      Else Result := 'CC_TYPE="NAT" And CC_CODE Like"'+ComboEntite.Value+'%" ' ;
    END
  Else
    BEGIN
    If EstSerie(S3) Then Result := 'CC_TYPE="NAT" And CC_CODE Like"'+ComboEntite.Value+'%" AND CC_CODE<="'+ComboEntite.Value+'02" '
                    Else Result := 'CC_TYPE="NAT" And CC_CODE Like"'+ComboEntite.Value+'%" ' ;
    END ;
end;

Initialization
  registerclasses ( [ TOF_CPCHOIXTALI ] ) ;
end.
