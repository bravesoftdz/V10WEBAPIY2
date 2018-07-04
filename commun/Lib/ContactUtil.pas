unit ContactUtil;

interface

uses  HEnt1, sysutils, UTOB, StdCtrls, Controls, Classes,
      forms,  ComCtrls, HCtrls, HMsgBox, Hdimension,  UTOM,
      {$IFDEF EAGLCLIENT}
        MaineAGL,
      {$ELSE EAGLCLIENT}
        db,
        {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
        DBGrids,
        {$IFNDEF ERADIO}
          Fe_Main,
        {$ENDIF !ERADIO}
      {$ENDIF EAGLCLIENT}
      variants,
      {$IFNDEF ERADIO}
        LookUp,
        AGLInit,
      {$ENDIF !ERADIO}
      M3FP,
      {$IFDEF PGISIDE}
      UtilGC,
      {$ENDIF PGISIDE}
      ParamSoc;

{ Creation de contact via une tob externe }
function CreateContactFromTob(Tiers, Contact: String; TheTob: Tob): String;
function CreateOrUpdateContactFromTob(Tiers, Contact: String; TheTob: Tob): String;

implementation

{$IFDEF GCGC}
uses  EntGC
    {$IFNDEF PGIMAJVER}
    ,UTomContact
    {$ENDIF PGIMAJVER}
    , TiersUtil
    ;
{$ENDIF GCGC}

//==============================================================================
//  fonction de creation d'un contact à partir d'une tob.
//  elle applique les regles de validation definies dans la tom contact
//  et retourne un message d'erreur ainsi qu'un champ Error eventuel.
//  les champs de donnees à renseigner sont fournis en tant que champs sup
//  dans TheTob.
//==============================================================================
function CreateContactFromTob(Tiers, Contact: String; TheTob: Tob): String;
var
  iChamp: Integer;
  TypeContact, FieldName: String;
  Exist, IsValid: Boolean;
  TobContact: Tob;
  TomContact: Tom;
//  BBI Correction de bug
  stPref : String;
//  BBI Fin Correction de bug
begin
  Result := '';
  { Contrôle existence du tiers }
  if not ExisteSQL('Select T_TIERS from TIERS where T_TIERS="' + Tiers + '"') then
  begin
    Result := TraduireMemoire('Le tiers ') + Tiers + TraduireMemoire(' n''existe pas dans la table : ') + 'TIERS';
//  BBI retour message erreur
    TheTob.AddChampSupValeur('Error', Result);
//  BBI Fin retour message erreur
    Exit;
  end;
  { Contrôle existence du contact si Contact contient quelque chose}
  if Contact <> '' then
    if ExisteSQL('Select C_TIERS from CONTACT where C_TIERS="' + Tiers + '" and C_NUMEROCONTACT=' + Contact) then
    begin
      Result := CreateOrUpdateContactFromTob(Tiers, Contact, TheTob);
//  BBI retour message erreur
      TheTob.AddChampSupValeur('Error', Result);
//  BBI Fin retour message erreur
      Exit;
    end;

  if TheTob <> nil then
  begin
    { Crée une TobContact et contrôle l'existence des champs de la tob passée en paramêtre }
    TobContact := Tob.Create('CONTACT', nil, -1);
    stPref := TableToPrefixe('CONTACT');
    try
      { Recopie la tob à mettre à jour dans la tob contact }
      iChamp := 999; Exist := True;
      while (iChamp < (1000 + (TheTob.ChampsSup.Count - 1))) and Exist do
      begin
        Inc(iChamp);
//  BBI Correction de bug
        FieldName := TheTob.GetNomChamp(iChamp);
        if Copy(FieldName, Pos('_', FieldName) + 1, 255) = 'CODEPER' then Continue; 
        { Vérifie si le champ fait partie de la table contact }
        Exist := (TobContact.FieldExists(FieldName) and (Copy(FieldName, 0, Pos('_', FieldName) - 1) = stPref));
//  BBI Fin Correction de bug
        if TheTob.GetNomChamp(iChamp) = 'C_TYPECONTACT' then
          TypeContact := TheTob.GetValue(TheTob.GetNomChamp(iChamp));
//  BBI Force numero de contact à 0
        if TheTob.GetNomChamp(iChamp) = 'C_NUMEROCONTACT' then
          TheTob.PutValue(TheTob.GetNomChamp(iChamp), 0);
//  BBI Fin Force numero de contact à 0
      end;
      if Exist then
      begin
        { Vérifie les données en passant par une TomContact }
        if TypeContact = '' then TypeContact := 'T';
        TomContact := CreateTOM('CONTACT', nil, false, true);
        TomContact.Argument('ORIGINE=PGISIDE;TYPE='+ TypeContact + ';TIERS=' + Tiers);
        TomContact.InitTOB(Tobcontact);
        { Recopie la tob à mettre à jour dans la tob Contact }
        iChamp := 999;
        while (iChamp < (1000 + (TheTob.ChampsSup.Count - 1))) do
        begin
          Inc(iChamp);
          FieldName := TheTob.GetNomChamp(iChamp);
          if Copy(FieldName, Pos('_', FieldName) + 1, 255) = 'CODEPER' then Continue;
          TobContact.PutValue(FieldName, TheTob.GetValue(FieldName));
        end;
        try
          IsValid := TomContact.VerifTOB(TobContact);
          Result := TomContact.LastErrorMsg;
        finally
          TomContact.Free;
        end;
        if IsValid then
        begin
          try
            TobContact.InsertDB(nil, False); { Enregistre la TobContact }
            TheTob.Dupliquer(TobContact, False, True);
          except
            on E: Exception do
            begin
              {$IFDEF PGISIDE}
              MAJJnalEvent('PSE', 'ERR', TraduireMemoire('Mise à jour de la table CONTACT'), E.Message);
              {$ENDIF PGISIDE}
              Result := E.Message;
            end;
          end;
        end;
//        else
//          Result := TraduireMemoire('Les données du contact : ') + Contact + TraduireMemoire(' ne sont pas valides');
      end
      else
        Result := TraduireMemoire('Le champ : ') + FieldName + TraduireMemoire(' n''existe pas dans la table : ') + 'CONTACT';
    finally
      TobContact.Free;
    end;
  end
  else
    Result := TraduireMemoire('Paramètres d''appel incorrect (Tob non renseignée)');
  if Result <> '' then
    TheTob.AddChampSupValeur('Error', Result);
end;

//==============================================================================
//  fonction de creation et/ou modif d'un contact à partir d'une tob.
//  elle applique les regles de validation definies dans la tom contact
//  et retourne un message d'erreur ainsi qu'un champ Error eventuel.
//  les champs de donnees à renseigner sont fournis en tant que champs sup
//  dans TheTob.
//==============================================================================
function CreateOrUpdateContactFromTob(Tiers, Contact: String; TheTob: Tob): String;
var
  iChamp: Integer;
  TypeContact, CodeAuxi, FieldName: String;
  ContactExist, Exist, IsValid: Boolean;
  TobContact, TobLoc: Tob;
  TomContact: Tom;
//  BBI Correction de bug
  stPref : String;
//  BBI Fin Correction de bug
begin
  Result := '';
  CodeAuxi := TiersAuxiliaire(Tiers);
  if TheTob.FieldExists('C_TYPECONTACT') then
    TypeContact := TheTob.GetValue('C_TYPECONTACT');
  { Contrôle existence du contact }
  ContactExist := ExisteSQL('Select C_TIERS from CONTACT where C_TIERS="' + Tiers + '" and C_NUMEROCONTACT=' + Contact);

  if TheTob <> nil then
  begin
    { Crée une TobContact et contrôle l'existence des champs de la tob passée en paramêtre }
    TobContact := Tob.Create('CONTACT', nil, -1);
    if ContactExist then
    begin
      TobLoc := TOB.Create('', nil, -1);
      TobLoc.LoadDetailDBFromSQL('CONTACT', 'Select * from CONTACT where C_TIERS="' + Tiers + '" and C_NUMEROCONTACT=' + Contact);
      if TobLoc.Detail.Count > 0 then
        TobContact.Dupliquer(TobLoc.Detail[0], True, True);
      TobLoc.Free;
    end;
//  BBI Correction de bug
    stPref := TableToPrefixe('CONTACT');
//  BBI Fin Correction de bug
    try
      { Recopie la tob à mettre à jour dans la tob contact }
      iChamp := 999; Exist := True;
      while (iChamp < (1000 + (TheTob.ChampsSup.Count - 1))) and Exist do
      begin
        Inc(iChamp);
//  BBI Correction de bug
        FieldName := TheTob.GetNomChamp(iChamp);
        { Vérifie si le champ fait partie de la table contact }
        Exist := (TobContact.FieldExists(FieldName) and (Copy(FieldName, 0, Pos('_', FieldName) - 1) = stPref));
//  BBI Fin Correction de bug
        if TheTob.GetNomChamp(iChamp) = 'C_TYPECONTACT' then
          TypeContact := TheTob.GetValue(TheTob.GetNomChamp(iChamp));
      end;
      if Exist then
      begin
        { Vérifie les données en passant par une TomContact }
        TomContact := CreateTOM('CONTACT', nil, false, true);
        TomContact.Argument('ORIGINE=PGISIDE;TYPE='+ TypeContact + ';TIERS=' + Tiers);
//        TomContact.InitTOB(Tobcontact);
        { Recopie la tob à mettre à jour dans la tob Contact }
        iChamp := 999;
        while (iChamp < (1000 + (TheTob.ChampsSup.Count - 1))) do
        begin
          Inc(iChamp);
          FieldName := TheTob.GetNomChamp(iChamp);
          TobContact.PutValue(FieldName, TheTob.GetValue(FieldName));
        end;
        try
          IsValid := TomContact.VerifTOB(TobContact);
          Result := TomContact.LastErrorMsg;
        finally
          TomContact.Free;
        end;
        if IsValid then
        begin
          TobContact.InsertOrUpdateDB; { Enregistre la TobContact }
          TheTob.Dupliquer(TobContact, False, True);
        end;
//        else
//          Result := TraduireMemoire('Les données du contact : ') + Contact + TraduireMemoire(' ne sont pas valides');
      end
      else
        Result := TraduireMemoire('Le champ : ') + FieldName + TraduireMemoire(' n''existe pas dans la table : ') + 'CONTACT';
    finally
      TobContact.Free;
    end;
  end
  else
    Result := TraduireMemoire('Paramètres d''appel incorrect (Tob non renseignée)');
  if Result <> '' then
    TheTob.AddChampSupValeur('Error', Result);
end;

end.
