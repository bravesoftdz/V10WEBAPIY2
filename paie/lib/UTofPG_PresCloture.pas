{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGPRESCLOTURE ()
Mots clefs ... : TOF;PGPRESCLOTURE
*****************************************************************}
{
PT1 07/02/2008 GGU V_81 Utilisation de 2 paramsoc pour gérer les dates de clôture
}
Unit UTofPG_PresCloture ;

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
     UTOF ; 

Type
  TOF_PGPRESCLOTURE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    ListeFinDeMois : Array of TDateTime;
  end ;

Implementation

Uses
  Paramsoc, Math, PgPresence;

Const
  iMaxSavedDate : Integer = 13;

procedure TOF_PGPRESCLOTURE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGPRESCLOTURE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGPRESCLOTURE.OnUpdate ;
var
  IndexEcran, IndexTableau : Integer;
  DateCloture, TempDate : TDateTime;
  ListeParam : String;
begin
  Inherited ;
  { Sauvegarde des dates de fin modifiées }
  { Lecture des données saisies et sauvegarde dans le tableau }
  for IndexEcran := 1 to iMaxSavedDate do
  begin
    TempDate := StrToDate(GetControlText('EDMOIS'+IntToStr(IndexEcran)));
    if TempDate < 10 then
    begin
      SetLength(ListeFinDeMois, IndexEcran -1 );
      Break;
    end;
    if Length(ListeFinDeMois) < IndexEcran then
      SetLength(ListeFinDeMois, IndexEcran);
    ListeFinDeMois[IndexEcran-1] := TempDate;
  end;  
  { Clôture à la date saisie }
  DateCloture := StrToDate(GetControlText('EDMOISCLOTURE'));
  if DateCloture > 10 then
  begin
    { On parcours les dates de fin de mois }
    for IndexTableau := 0 to Length(ListeFinDeMois)-1 do
    begin
      TempDate := ListeFinDeMois[IndexTableau];
      { On supprime celles qui sont superieur à la date de cloture }
      if (TempDate < 10) or (TempDate >= DateCloture) then
      begin
        SetLength(ListeFinDeMois, IndexTableau );
        Break;
      end;
    end;
    SetLength(ListeFinDeMois, Length(ListeFinDeMois) + 1 );
    ListeFinDeMois[Length(ListeFinDeMois)-1] := StrToDate(GetControlText('EDMOISCLOTURE'));
    { On ne garde que les 13 (iMaxSavedDate) dernières dates }
    ListeFinDeMois := copy(ListeFinDeMois, max(0, Length(ListeFinDeMois) - iMaxSavedDate) , iMaxSavedDate);
  end;
  { Construction de la chaine de caractère du parametre societe }
  for IndexTableau := 0 to Length(ListeFinDeMois) -1 do
  begin
    ListeParam := ListeParam + DateToStr(ListeFinDeMois[IndexTableau]) + ';';
  end;
  { Mise à jour du parametre societe }
  SetParamSocDatesCloture(MOIS, ListeParam); //PT1 SetParamSoc(PresenceRenvoieParamSoc(MOIS), ListeParam);
  {$IFDEF EAGLCLIENT}
  AvertirCacheServer('PARAMSOC');
  {$ENDIF}
  { Mise à jour de l'écran }
  OnArgument('');
end ;

procedure TOF_PGPRESCLOTURE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGPRESCLOTURE.OnArgument (S : String ) ;
var
  ListeParam, Param : String;
  IndexTableau : Integer;
  Q : TQuery;
  InitClotureDate : TDateTime;
begin
  Inherited ;  
  { Lecture du parametre societe }
  ListeParam := GetParamSocDatesCloture(MOIS); //PT1 GetParamSocSecur(PresenceRenvoieParamSoc(MOIS), '');
  Param := ReadTokenSt(ListeParam);
  { Remplissage du tableau }
  SetLength(ListeFinDeMois, 0);
  While (Param <> '') Do
  Begin
    SetLength(ListeFinDeMois, Length(ListeFinDeMois) + 1 );
    ListeFinDeMois[Length(ListeFinDeMois) - 1] := StrToDate(Param);
    Param := ReadTokenSt(ListeParam);
  End;
  { On ne garde que les 13 (iMaxSavedDate) dernières dates }
  ListeFinDeMois := copy(ListeFinDeMois, max(0, Length(ListeFinDeMois) - iMaxSavedDate) , iMaxSavedDate);

  { Lecture du tableau et écriture à l'écran }
  for IndexTableau := 0 to Length(ListeFinDeMois) - 1 do
  begin
    SetControlText('EDMOIS'+IntToStr(IndexTableau+1), DateToStr(ListeFinDeMois[IndexTableau]));
  end;
  for IndexTableau := Length(ListeFinDeMois) to (iMaxSavedDate - 1) do
  begin
    SetControlText('EDMOIS'+IntToStr(IndexTableau+1), DateToStr(iDate1900));
  end;
  { Initialisation du champs de cloture :
    On cherche la dernière date de calcul d'un compteur mensuel }
  InitClotureDate := iDate1900;
  Q := OpenSQL('select max(PYP_DATEFINPRES) from presencesalarie where pyp_periodicitepre = "'+periodiciteMensuelle+'"', True);
  if not Q.Eof then
    InitClotureDate := Q.Fields[0].AsDateTime;
  Ferme(Q);
  SetControlText('EDMOISCLOTURE', DateToStr(InitClotureDate));
//  ShowDate;
end ;

procedure TOF_PGPRESCLOTURE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGPRESCLOTURE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGPRESCLOTURE.OnCancel () ;
begin
  Inherited ;
end ;

//procedure TOF_PGPRESCLOTURE.ShowDate;
//var
//  IndexEcran : Integer;
//  TempDate : TDateTime;
//begin
//  for IndexEcran := 1 to 12 do
//  begin
//    TempDate := StrToDate(GetControlText('EDMOIS'+IntToStr(IndexEcran)));
//    SetControlVisible('EDMOIS'+IntToStr(IndexEcran), (TempDate >= 10));
//  end;
//end;

//procedure TOF_PGPRESCLOTURE.OnExitListDate(Sender: TObject);
//begin
//  ShowDate;
//end;

Initialization
  registerclasses ( [ TOF_PGPRESCLOTURE ] ) ; 
end.
