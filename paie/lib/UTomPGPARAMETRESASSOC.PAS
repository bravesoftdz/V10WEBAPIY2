{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 27/03/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PGPARAMETRESASSOC (PGPARAMETRESASSOC)
Mots clefs ... : TOM;PGPARAMETRESASSOC
*****************************************************************
Pt1  07/09/07  NA   V8_00  FQ 14711 Modif affichage des paramètres et population selon le prédéfini de l'association
PT3 11/06/08 FC V_850 FQ 15534 Pb alimentation PGO_TYPEPARAMETRE et du coup pb à la récupération de la population
}
Unit UTomPGPARAMETRESASSOC;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     Fiche, 
     FichList, 
{$else}
     eFiche, 
     eFichList, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOM, 
     UTob, HDB, utobdebug;

Type
  TOM_PGPARAMETRESASSOC = Class (TOM)
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
    CEG, STD, DOS : boolean;
    TobPop, TobParam : TOB;
    IsLoading : Boolean;
{$IFNDEF EAGLCLIENT}
    Control_CODEASSOCIE, Control_PREDEFINI, Control_TYPEPARAMETRE : THDBValcombobox;
    Control_PGPARAMETRE : THDBEdit;
{$ELSE}
    Control_CODEASSOCIE, Control_PREDEFINI, Control_TYPEPARAMETRE : THValcombobox;
    Control_PGPARAMETRE : THEdit;
{$ENDIF}
    procedure MajTablettes(Sender : TObject);
    procedure TestValidite(var LastError : integer;var LastErrorMsg : WideString);
    function TobFindFirst(TheTob: TOB; NomChampInside,
      ValeurChampInside: String; NomChamp: array of String;
      ValeurChamp: array of Variant; MultiNiveau: Boolean): TOB;
    procedure OnClickBDelete1(Sender : TObject);
    procedure OnClickNewRecord1 (Sender : TObject);
  end ;

Implementation
uses
  pgoutils, pgoutils2,  ed_tools, HTB97, PGPOPULOUTILS, StrUtils;   //  TntDBCtrls,

procedure TOM_PGPARAMETRESASSOC.OnClickNewRecord1 (Sender : TObject);
begin
//  Inherited ;
  if (DS.state = dsEdit) or (DS.state = dsInsert) then
    if PGIAsk(TraduireMemoire('Les données salariés vont êtres mises à jour.#10#13Voulez-vous continuer ?'),Ecran.Caption) = mrYes then
    begin
      (GetControl('BValider') as TToolbarbutton97).OnClick(Sender);
    end else begin
      LastError := 1;
      LastErrorMsg := TraduireMemoire('L''association n''a pas encore été créée.');
    end;
end ;

procedure TOM_PGPARAMETRESASSOC.OnClickBDelete1(Sender : TObject) ;
begin
//  Inherited ;
  if PGIAsk(TraduireMemoire('Les données salariés vont êtres mises à jour.#10#13Voulez-vous continuer ?'),Ecran.Caption) = mrYes then
  begin
    (GetControl('BDelete') as TToolbarbutton97).OnClick(Sender);
  end else begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('L''association n''a pas été supprimée.');
  end;
end ;

procedure TOM_PGPARAMETRESASSOC.OnUpdateRecord ;
//var
//  LeTypePop : String;
begin
  Inherited ;
  if not IsLoading then
  begin
    //On met à jour l'enregistrement pour stocker le type pop :
    SetField('PGO_TYPEPOP', leftstr(GetField('PGO_PGPARAMETRE'),3) );
    lastError := 0;
    TestValidite(lastError, LastErrorMsg);
  end;
end ;

procedure TOM_PGPARAMETRESASSOC.OnAfterUpdateRecord ;
{$IFDEF IDEMPOP}
var
//  UpdateIdemPop : TUpdateIdemPop;
  LeTypePop : String; //, TypePop
{$ENDIF}
begin
  Inherited ;
{  //On recherche le type de population concerné
  Q := OpenSQL('Select PPO_TYPEPOP from CODEPOPULATION where PPO_CODEPOP = "'+ GetField('PGO_CODEASSOCIE')+'"',True,1);
  if not Q.Eof then
  begin
    LesTypePop := Q.Fields[0].AsString;
  end;
  Ferme(Q);
  //On met à jour les données et la table d'association
//  UpdateIdemPop := TUpdateIdemPop.Create;
//  UpdateIdemPop.MajSALARIEPOPUL(LeTypePop ,Date);
//  UpdateIdemPop.Free;
}
{$IFDEF IDEMPOP}
  LeTypePop := GetField('PGO_TYPEPOP');
  //On dé-valide le parametrage de ces populations
  SetValidePopulParamSoc(LeTypePop, False);
  Exit;
{$ENDIF}
end ;

procedure TOM_PGPARAMETRESASSOC.OnAfterDeleteRecord ;
{$IFDEF IDEMPOP}
var
//  UpdateIdemPop : TUpdateIdemPop;
  LesTypePop, LeTypePop : String;
  Q : TQuery;
{$ENDIF}
begin
  Inherited ;
{$IFDEF IDEMPOP}
  //On recherche le type de population concerné
  Q := OpenSQL('Select PPO_TYPEPOP from CODEPOPULATION where PPO_CODEPOP = "'+ GetField('PGO_CODEASSOCIE')+'"',True,1);
  if not Q.Eof then
  begin
    LesTypePop := Q.Fields[0].AsString;
  end;
  Ferme(Q);
  //On met à jour les données et la table d'association
//  UpdateIdemPop := TUpdateIdemPop.Create;
//  UpdateIdemPop.MajSALARIEPOPUL(LesTypePop ,Date);
//  UpdateIdemPop.Free;
  LeTypePop := Trim(ReadTokenPipe(LesTypePop, ';'));
  while (LesTypePop <> '') or (LeTypePop <> '') do
  begin
    SetValidePopulParamSoc(LeTypePop, False);
    LeTypePop := Trim(ReadTokenPipe(LesTypePop, ';'));
  end;
  Exit;
{$ENDIF}
end ;

procedure TOM_PGPARAMETRESASSOC.OnLoadRecord ;
begin
  Inherited ;
  IsLoading := False;
  if Control_TYPEPARAMETRE.value = '' then Control_TYPEPARAMETRE.value := 'POP';
  MajTablettes(Self);
end ;

procedure TOM_PGPARAMETRESASSOC.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_PGPARAMETRESASSOC.OnArgument ( S: String ) ;
begin
  IsLoading := True;
  Inherited ;
{$IFNDEF EAGLCLIENT}
  Control_PREDEFINI := GetControl('PGO_PREDEFINI') as THDBValcombobox;
  Control_CODEASSOCIE := GetControl('PGO_CODEASSOCIE') as THDBValcombobox;
  Control_TYPEPARAMETRE := GetControl('PGO_TYPEPARAMETRE') as THDBValcombobox;
  Control_PGPARAMETRE := GetControl('PGO_PGPARAMETRE') as THDBEdit;
{$ELSE}
  Control_PREDEFINI := GetControl('PGO_PREDEFINI') as THValcombobox;
  Control_CODEASSOCIE := GetControl('PGO_CODEASSOCIE') as THValcombobox;
  Control_TYPEPARAMETRE := GetControl('PGO_TYPEPARAMETRE') as THValcombobox;
  Control_PGPARAMETRE := GetControl('PGO_PGPARAMETRE') as THEdit;
{$ENDIF}
  TobPop := TOB.Create('Population mère', nil, -1);
  TobParam := TOB.Create('Paramètres mère', nil, -1);
  TobPop.ClearDetail;
  TobPop.LoadDetailDBFromSQL('Population mère',
         'select PPC_POPULATION, PPC_PREDEFINI, PPC_NODOSSIER, PPO_TYPEPOP '
        +'  from ORDREPOPULATION '
        +'  left outer join CODEPOPULATION on PPC_CODEPOP = PPO_CODEPOP '
        +'                                and PPC_PREDEFINI = PPO_PREDEFINI '
        +' where ##PPC_PREDEFINI##');
  TobParam.ClearDetail;
  TobParam.LoadDetailDBFromSQL('Paramètres mère',
         'select distinct PGP_PGPARAMETRE, PGP_TYPEPARAMETRE, PGP_PREDEFINI, PGP_NODOSSIER, PGP_MODULECHAMP '
        +'  from PGPARAMETRES '
        +' where ##PGP_PREDEFINI##');
  Control_CODEASSOCIE.OnChange := MajTablettes;
  Control_PGPARAMETRE.OnExit := MajTablettes;
  Control_PREDEFINI.OnChange := MajTablettes;
  (GetControl('BDelete1') as TToolbarbutton97).OnClick := OnClickBDelete1;
  (GetControl('BValider1') as TToolbarbutton97).OnClick := OnClickNewRecord1;
end ;

procedure TOM_PGPARAMETRESASSOC.OnClose ;
begin
  FreeAndNil(TobPop);
  FreeAndNil(TobParam);
  Inherited ;
end ;

procedure TOM_PGPARAMETRESASSOC.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_PGPARAMETRESASSOC.OnLoadAlerte;
begin
  Inherited ;
end ;

procedure TOM_PGPARAMETRESASSOC.TestValidite(var LastError: integer;
  var LastErrorMsg: WideString);
var
  Predef: String;
    QChamps : TQuery;
    stQChamps : String;
begin
  Predef := GetField('PGO_PREDEFINI');
  if (DS.State = dsinsert) then
  begin
    AccesPredefini('TOUS', CEG, STD, DOS);
    if (Predef = 'CEG') and (CEG = FALSE) then
    begin
      Predef := 'DOS';
      LastError := 1;
      LastErrorMsg := 'Vous ne pouvez pas créer d''association prédéfinie CEGID';
      SetControlProperty('PGO_PREDEFINI', 'Value', Predef);
      SetFocusControl('PGO_PREDEFINI');
      exit;
    end;
    if (Predef = 'STD') and (STD = FALSE) then
    begin
      Predef := 'DOS';
      LastError := 1;
      LastErrorMsg := 'Vous ne pouvez pas créer d''association prédéfinie Standard';
      SetControlProperty('PGO_PREDEFINI', 'Value', Predef);
      SetFocusControl('PGO_PREDEFINI');
      exit;
    end;
  end;

  if (DS.State = dsinsert) then
  begin
    if (GetField('PGO_PREDEFINI') <> 'DOS') then
      SetField('PGO_NODOSSIER', '000000')
    else
      SetField('PGO_NODOSSIER', PgRendNoDossier());
  end;

  if (GetField('PGO_PREDEFINI') = '') then
  begin
    LastError := 1;
    LastErrorMsg := 'Vous devez choisir le type de prédéfini';
    SetFocusControl('PGO_PREDEFINI');
    exit;
  end;

  if (GetField('PGO_PGPARAMETRE') = '') then
  begin
    LastError := 1;
    LastErrorMsg := 'Vous devez choisir le paramètre';
    SetFocusControl('PGO_PGPARAMETRE');
    exit;
  end;

  if (GetField('PGO_CODEASSOCIE') = '') then
  begin
    LastError := 1;
    LastErrorMsg := 'Vous devez choisir le code associé';
    SetFocusControl('PGO_CODEASSOCIE');
    exit;
  end;

  if (GetField('PGO_DATEVALIDITE') = iDate1900) then
  begin
    LastError := 1;
    LastErrorMsg := 'Vous devez choisir une date de validité';
    SetFocusControl('PGO_DATEVALIDITE');
    exit;
  end;

  //Faire le test de duplicate key
  stQChamps  := 'SELECT ##TOP 1## 1 '
               + ' FROM PGPARAMETRESASSOC '
               + 'WHERE ##PGO_PREDEFINI## '
               + '  AND PGO_PREDEFINI = "'+Predef+'" '
               + '  AND PGO_CODEASSOCIE = "'+Control_CODEASSOCIE.Value+'" '
               + '  AND PGO_PGPARAMETRE = "'+Control_PGPARAMETRE.Text+'" '
               + '  AND PGO_TYPEPARAMETRE = "'+Control_TYPEPARAMETRE.Value+'" '
               + '  AND PGO_DATEVALIDITE = "'+USDATETIME( AGlStrToDate(GetControlText('PGO_DATEVALIDITE')))+'" ';
  QChamps := OpenSQL(stQChamps,True,1);
  if not QChamps.Eof then
  begin
    LastError := 1;
    LastErrorMsg := 'Cette association existe déjà.';
    SetFocusControl('PGO_DATEVALIDITE');
  end;
  Ferme(QChamps);

end;

//Permet de rechercher la première fille qui correspond aux critères habituelles
//et dont la valeur du champ "NomChampInside" contient "ValeurChampInside"
function TOM_PGPARAMETRESASSOC.TobFindFirst(TheTob : TOB; NomChampInside : String;
    ValeurChampInside : String; NomChamp : Array of String ;
    ValeurChamp : Array of Variant ; MultiNiveau : Boolean ) : TOB;
begin
  result := TheTob.FindFirst(NomChamp,ValeurChamp, MultiNiveau);
  if NomChampInside = '' then exit;
  if ValeurChampInside = '' then exit;
  while result <> nil do
  begin
    if   ( pos(ValeurChampInside,result.getstring(NomChampInside)) > 0)
      or ( pos(result.getstring(NomChampInside),ValeurChampInside) > 0)
     then                 
      exit
    else
      result := TheTob.FindNext(NomChamp,ValeurChamp, MultiNiveau);
  end;
  result := nil;
end;

procedure TOM_PGPARAMETRESASSOC.MajTablettes(Sender: TObject);
var
  Predef, PPredef, CODEASSOCIE, PGPARAMETRE, TYPEPARAMETRE ,ListePGPPredef, PPCPredef: String;  //pt1
  PPO_TYPEPOP, PGP_MODULECHAMP : String;
  TempTOB : TOB;
  ListePop : String;
  index : Integer;
  nodossier : String;
begin
  PPCPredef := '';
  ListePGPPredef := '';
  Predef := Control_PREDEFINI.Value;
  CODEASSOCIE := Control_CODEASSOCIE.Value;
  PGPARAMETRE := Control_PGPARAMETRE.Text;
  TYPEPARAMETRE := Control_TYPEPARAMETRE.Value;
  //On cherche les types de population/modules associés qui sont déjà défini
  if (Predef <> 'DOS') then
    nodossier := '000000'
  else
    nodossier := PgRendNoDossier();

  TempTOB := TobParam.FindFirst(['PGP_PGPARAMETRE', 'PGP_TYPEPARAMETRE', 'PGP_PREDEFINI', 'PGP_NODOSSIER'], [PGPARAMETRE,TYPEPARAMETRE,Predef,nodossier], False);
  if TempTOB <> nil then
    PGP_MODULECHAMP := TempTOB.GetString('PGP_MODULECHAMP')
  else
    PGP_MODULECHAMP := '';

    if PPCPRedef = '' then PPCPredef := GetPredefiniPopulation(PGP_MODULECHAMP);  // pt1 recherche population valide pour ce type de paramètre
  TempTOB := TobPop.FindFirst(['PPC_POPULATION', 'PPC_PREDEFINI', 'PPC_NODOSSIER'], [CODEASSOCIE,PPCPredef,nodossier], False);
  if TempTOB <> nil then
    PPO_TYPEPOP := TempTOB.GetString('PPO_TYPEPOP')
  else
    PPO_TYPEPOP := '';
  // Test de l'existance des différents types de prédéfini et affectation du type recherché
  // Pour un prédéfini dossier, on regarde si il existe un paramétra dossier / Standard /cegid (pour le module concerné),
  // pour un prédéfini Standard on affiche les paramétras Std/CEG, pour un CEGID on affiche uniquement du CEGID
//  PPredef := '';  pt1
 // deb Pt1


   if (Predef = 'CEG')
     and Assigned(TobFindFirst(TobPop,'PPO_TYPEPOP',PPO_TYPEPOP,['PPC_PREDEFINI'], ['CEG'], False))
     and Assigned(TobFindFirst(TobParam,'PGP_MODULECHAMP',PGP_MODULECHAMP,['PGP_PREDEFINI'], ['CEG'], False)) then
     begin
       PPCPredef := 'CEG';
       ListePGPPredef := ',"CEG"';
     end;
   if (Predef = 'STD')
     //and Assigned(TobFindFirst(TobParam,'PGP_MODULECHAMP',PGP_MODULECHAMP,['PGP_PREDEFINI'], ['STD'], False))
     then ListePGPPredef := ',"CEG","STD"';
   if ( Predef = 'DOS')
    // and Assigned(TobFindFirst(TobParam,'PGP_MODULECHAMP',PGP_MODULECHAMP,['PGP_PREDEFINI'], ['DOS','STD','CEG'], False))
     then ListePGPPredef := ',"CEG","STD", "DOS"';

 { if ( Predef = 'DOS')
     and Assigned(TobFindFirst(TobParam,'PGP_MODULECHAMP',PGP_MODULECHAMP,['PGP_PREDEFINI'], ['DOS'], False))
     and Assigned(TobFindFirst(TobPop,'PPO_TYPEPOP',PPO_TYPEPOP,['PPC_PREDEFINI'], ['DOS'], False)) then
    PPredef := 'DOS';
  if (PPredef = '') and ((Predef = 'DOS') or (Predef = 'STD'))
     and Assigned(TobFindFirst(TobParam,'PGP_MODULECHAMP',PGP_MODULECHAMP,['PGP_PREDEFINI'], ['STD'], False))
     and Assigned(TobFindFirst(TobPop,'PPO_TYPEPOP',PPO_TYPEPOP,['PPC_PREDEFINI'], ['STD'], False)) then
    PPredef := 'STD';
  if (PPredef = '')
     and Assigned(TobFindFirst(TobPop,'PPO_TYPEPOP',PPO_TYPEPOP,['PPC_PREDEFINI'], ['CEG'], False))
     and Assigned(TobFindFirst(TobParam,'PGP_MODULECHAMP',PGP_MODULECHAMP,['PGP_PREDEFINI'], ['CEG'], False)) then
    PPredef := 'CEG'; } // pt1

  //Création de la liste des populations qui ont un typepop égal au modulechamp
  ListePop := '';
  for index := 0 to TobPop.Detail.Count -1 do
  begin
    TempTOB := TobPop.Detail[index];
    if pos(PGP_MODULECHAMP,TempTOB.GetString('PPO_TYPEPOP')) > 0 then
    begin
      ListePop := ListePop + ',"' + TempTOB.GetString('PPC_POPULATION')+'"';
    end;
  end;
  //  deb pt1

 // Control_CODEASSOCIE.Plus := ' AND PPC_PREDEFINI = "'+PPredef+'" AND ("'+PGP_MODULECHAMP+'"= "" or PPC_POPULATION in (""'+ListePop+'))';
 // Control_PGPARAMETRE.Plus := ' AND PGP_PREDEFINI = "'+PPredef+'" AND ("'+PPO_TYPEPOP+'"= "" or PGP_MODULECHAMP in ("'+StringReplace(PPO_TYPEPOP,';','","',[rfReplaceAll, rfIgnoreCase])+'"))';
 Control_CODEASSOCIE.Plus := ' AND PPC_PREDEFINI = "'+PPCPredef+'" AND ("'+PGP_MODULECHAMP+'"= "" or PPC_POPULATION in (""'+ListePop+'))';
 Control_PGPARAMETRE.Plus := ' AND PGP_PREDEFINI in (""'+ListePGPPredef+') AND ("'+PPO_TYPEPOP+'"= "" or PGP_MODULECHAMP in ("'+StringReplace(PPO_TYPEPOP,';','","',[rfReplaceAll, rfIgnoreCase])+'"))';
 //  fin pt1
  // On vide Control_PGPARAMETRE si le code déjà choisi n'est plus dans la liste
  if RechDom(Control_PGPARAMETRE.DataType,Control_PGPARAMETRE.Text,False,Control_PGPARAMETRE.plus) = '' then
  begin
    Control_PGPARAMETRE.Text := '';
    SetField('PGO_PGPARAMETRE','');
  end;
end;

procedure TOM_PGPARAMETRESASSOC.OnDeleteRecord;
begin
  inherited;
end;

procedure TOM_PGPARAMETRESASSOC.OnNewRecord;
begin
  inherited;
  //A priori ne prend que la valeur POP. Cette valeur se trouve dans la tablette PGO_TYPEPARAMETRE
  //Le problème c'est qu'en CWAS, le champ reste vide ??? et du coup zone non alimentée
  //En 2/3 le champ est alimenté par défaut par la 1ère valeur de la tablette
  SetField('PGO_TYPEPARAMETRE','POP');  //PT3
end;

Initialization
  registerclasses ( [ TOM_PGPARAMETRESASSOC ] ) ;
end.
