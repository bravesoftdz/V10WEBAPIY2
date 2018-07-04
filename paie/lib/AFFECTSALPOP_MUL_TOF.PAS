{***********UNITE*************************************************
Auteur  ...... : NA
Créé le ...... : 26/04/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFFECTSALPOP_MUL ()
                 Liste des salariés affectés par type de population
Mots clefs ... : TOF;AFFECTSALPOP_MUL
*****************************************************************
PT1   : 01/08/2007 JL V_80 Gestion accès depuis fiche salarié
}
Unit AFFECTSALPOP_MUL_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     Entpaie,
     UTOF ; 

Type
  TOF_AFFECTSALPOP_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
    affichetypepop : boolean;
    procedure Affichepopul(sender : TObject);
  end ;

Implementation
uses PgOutils2,pgpopuloutils; // appel à la fonction PgrendNodossier pour affectation n° de dossier.

procedure TOF_AFFECTSALPOP_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFFECTSALPOP_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFFECTSALPOP_MUL.OnUpdate ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 07/05/2007
Modifié le ... :   /  /    
Description .. : Chargement des données
Mots clefs ... : 
*****************************************************************}
procedure TOF_AFFECTSALPOP_MUL.OnLoad ;
var
listetypepop, st, typepop, listpopValides, typepopvalide: string;
begin
  Inherited ;
  ListPopValides := VH_Paie.PGPopulValides;
  typepop := THValcomboBox(getcontrol('PNA_TYPEPOP')).value;
  if typepop = '' then
  begin
    listetypepop := '';
    typepopvalide := readtokenst(listpopvalides);
    while typepopvalide <> '' do
    begin
      if typepopvalide <> 'HAB' then
      listetypepop:= listetypepop + '"' + typepopvalide + '"' + ',';
      typepopvalide := readtokenst(listpopvalides);
    end;

    if listetypepop <> '' then
    begin
      listetypepop := '(' + copy(listetypepop, 1, length(listetypepop) - 1) + ')';
      st := 'PNA_TYPEPOP in '+listetypepop+'';
    end
    else
      st := 'PNA_POPULATION = "XXX"';
    setcontroltext('XX_WHERE', st);
  end
  else
  if affichetypepop = false then setcontrolText('XX_WHERE', 'PNA_POPULATION = "XXX"')
  else setcontrolText('XX_WHERE', '');

end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 26/04/2007
Modifié le ... :   /  /    
Description .. : On argument
Mots clefs ... : 
*****************************************************************}
procedure TOF_AFFECTSALPOP_MUL.OnArgument (S : String ) ;
var
typepopul: Thvalcombobox;
stplus,Salarie : string;
begin
  Inherited ;
  stplus := '';
  if not VH_PAIE.PGSeriaFormation then stplus := ' AND CO_CODE <> "FOR"';
  if not VH_PAIE.PGSeriaPresence then  stplus := stplus + ' AND CO_CODE <> "PRE"';
  if stplus <> '' then
  begin
  stplus := stplus + ' AND CO_CODE <> "HAB"';
  SetControlProperty('PNA_TYPEPOP', 'Plus', stplus);
  end;

  setcontrolenabled('PNA_POPULATION', false);
  affichetypepop := true;
  typepopul := THValComboBox(GetControl('PNA_TYPEPOP',TRUE));

  if typepopul <> nil then
    typepopul.OnChange := affichepopul;
  //DEBUT PT1
  Salarie := S;
  If Salarie <> '' then
  begin
    SetControlText('PNA_SALARIE',Salarie);
    SetControlenabled('PNA_SALARIE',False);
  end;
  //FIN PT1
end ;

procedure TOF_AFFECTSALPOP_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AFFECTSALPOP_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AFFECTSALPOP_MUL.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 26/04/2007
Modifié le ... :   /  /    
Description .. : Affichage des populations selon le type de population saisi 
Suite ........ : et selon le prédéfini à afficher
Mots clefs ... : 
*****************************************************************}
procedure TOF_AFFECTSALPOP_MUL.Affichepopul(sender: TObject);
var
typepop, pop, pred, nodossier,st : string;
populok : boolean;

begin
  pop := THValcomboBox(getcontrol('PNA_POPULATION')).value;
  typepop := THValcomboBox(getcontrol('PNA_TYPEPOP')).value;
  nodossier := PgrendNodossier();

  // recherche des populations à afficher
  if typepop <> '' then
  begin
      setcontrolenabled('PNA_POPULATION', true);
      SetControlProperty('PNA_POPULATION','Value','');
   // recherche si le type de population est valide
      populok := canusepopulation(typepop);
      if populok then
      begin
        affichetypepop := true;
        pred := GetPredefiniPopulation(typepop);
        st := ' and PPC_PREDEFINI = "'+pred+'" ';
        if typepop = 'IDR' then SetControlProperty('PNA_POPULATION','DataType','PGPOPULATIONIDR');
        if typepop = 'PAI' then SetControlProperty('PNA_POPULATION','DataType','PGPOPULATIONSAL');
        if typepop = 'HAB' then SetControlProperty('PNA_POPULATION','DataType','PGPOPULATIONHAB');
        setcontrolproperty('PNA_POPULATION', 'plus', st);
      end 
      else
      begin
        affichetypepop := false;
        setcontrolenabled('PNA_POPULATION', false);
        PgiInfo('Le type de population "'+typepop+'" n''est pas valide.', Ecran.caption);
      end;
  end
  else // si type pop = '': on ne sélectionne pas de population
  setcontrolenabled('PNA_POPULATION', false);

end;

Initialization
  registerclasses ( [ TOF_AFFECTSALPOP_MUL ] ) ; 
end.
