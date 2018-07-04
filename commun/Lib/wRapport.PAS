{***********UNITE*************************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 23/09/2002
Modifié le ... :   /  /
Description .. : Rapport d'actions
Mots clefs ... : WRAPPORT
*****************************************************************}
unit wRapport;

interface

uses
  Windows,
  Classes,
  SysUtils,
  HCtrls,
  Hent1,
  uTob,
  Dialogs,
  Forms,
  Controls,
  comctrls,
  stdctrls,
  {$IFNDEF EAGLCLIENT}
    {$IFNDEF EAGLSERVER}
    {$IFNDEF ERADIO}
      Fe_Main,
    {$ENDIF !ERADIO}
    {$ENDIF !EAGLSERVER}
  {$ELSE}
    MainEagl,
  {$ENDIF}
  HMsgBox,
  Graphics,
  Menus
  ;

  type TTypeWRP = (TWRP_Done, TWRP_Error, TWRP_None);

  type TWRapport = Class
    private
      IdWRP_Error, IdWRP_Done: Integer;
      TobError, TobDone, fTobEtat: Tob;
      fTitre: String;
      procedure PrintTobs(Liste: Boolean);
      procedure HappenTob(TobSource, TobDest: Tob);
    public
      constructor Create(TitreRapport: String);
      destructor Destroy; override;
      property TitreRapport: String Read fTitre Write fTitre;
      { Transport de l'objet }
      procedure AddObjectInTob(t: Tob);
      { Gestion }
      procedure Add(TypeWRP: TTypeWRP; const Clef, Texte: String); overload;
      procedure Add(TypeWRP: TTypeWRP; T: Tob; const Texte: String); overload;
      procedure RAZ(TypeWRP: TTypeWRP = TWRP_None);
      function GetTobWRP(TypeWRP: TTypeWRP): Tob;
      function Exists(TypeWRP: TTypeWRP): Boolean;
      function ErrorOn(Clef: Variant): Boolean;
      function DoneOn(Clef: Variant): Boolean;
      function CountError: Integer;
      function CountDone: Integer;
      function Count: Integer;
      procedure Display(ForceAffichage: Boolean = False; TitreCell1: String = '');
      { Get }
      function GetFirstError: String;
      function GetFirstDone: String;
      { Edition / Exportation }
      procedure Print;
      procedure ExportList;
      {$IFNDEF EAGLSERVER}
        {$IFNDEF ERADIO}
          procedure SaveToFile(WithAskFileName: Boolean = True; FileName: String = '');
        {$ENDIF !ERADIO}
      {$ENDIF !EAGLSERVER}
  end;

function wGetRapportFromTob(t: Tob): TWRapport;

const
  WRPTobFieldName = '_WRAPPORT';

implementation

uses
  wCommuns
  {$IFNDEF EAGLCLIENT}
    ,EdtREtat
  {$ELSE}
    ,UtilEagl
  {$ENDIF}
  {$IFNDEF EAGLSERVER}
    {$IFNDEF ERADIO}
      ,LookUp
    {$ENDIF !ERADIO}
  {$ENDIF !EAGLSERVER}
  ;

{ TWRapport }

constructor TWRapport.Create(TitreRapport: String);
begin
  fTitre := TitreRapport;
  IdWRP_Done := 0;
  IdWRP_Error := 0;

  TobError := Tob.Create('WRP_ERROR', nil, -1);
  TobDone := Tob.Create('WRP_DONE', nil, -1);
end;

destructor TWRapport.Destroy;
begin
  inherited;

  if Assigned(TobError) then
    TobError.Free;
  if Assigned(TobDone) then
    TobDone.Free;
end;

{ ajouter un élément du rapport }
procedure TWRapport.Add(TypeWRP: TTypeWRP; const Clef: String; const Texte: String);
var
  t: Tob;

  procedure LoadTobWRP(idWRP: Integer);
  begin
    t.AddChampSupValeur('TYPE', TypeWRP);
    t.AddChampSupValeur('IDENTIFIANTWXX', Clef);
    t.AddChampSupValeur('TEXTE', Texte);
    t.AddChampSupValeur('IDWRP', idWRP);
    t.AddChampSupValeur('TYPERAPPORT', iif(TypeWRP = TWRP_Done, 'DONE', 'ERROR'));
  end;

begin
  case TypeWRP of
    TWRP_Done : begin
                  Inc(IdWRP_Done);
                  t := Tob.Create('WRP_FILLE', TobDone, -1);
                  LoadTobWRP(IdWRP_Done);
                end;
    TWRP_Error: begin
                  Inc(IdWRP_Error);
                  t := Tob.Create('WRP_FILLE', TobError, -1);
                  LoadTobWRP(IdWRP_Error);
                end;
  end;
end;

procedure TWRapport.Add(TypeWRP: TTypeWRP; T: Tob; const Texte: String); 
begin
  Add(TypeWRP, wGetValueClef1(T), Texte)
end;

{ affichage du rapport }
procedure TWRapport.Display(ForceAffichage: Boolean = False; TitreCell1: String = '');
begin
{$IFNDEF EAGLSERVER}
  {$IFNDEF ERADIO}
    if Exists(TWRP_Error) or (ForceAffichage and Exists(TWRP_Done)) then
    begin
      if Count = 1 then
      begin
        if CountError = 1 then
          PgiError(GetFirstError, FTitre)
        else
          PgiInfo(GetFirstDone, FTitre)
      end
      else
        AglLanceFiche('W', 'WRAPPORT', '', '', 'TITRE=' + fTitre + ';' + WRPTobFieldName + '=' + IntToStr(LongInt(Self)) + ';TITRECELL1=' + TitreCell1);
    end;
  {$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
end;

{ renvoie la tob en fonction du type de rapport }
function TWRapport.GetTobWRP(TypeWRP: TTypeWRP): Tob;
begin
  case TypeWRP of
    TWRP_Done : Result := TobDone;
    TWRP_Error: Result := TobError;
  else Result := nil;
  end;
end;

{ Remise à zéro : vide les tobs du rapport }
procedure TWRapport.RAZ(TypeWRP: TTypeWRP);
begin
  Case TypeWRP of
    TWRP_Done : TobDone.ClearDetail;
    TWRP_Error: TobError.ClearDetail;
    TWRP_None : begin
                  RAZ(TWRP_Done);
                  RAZ(TWRP_Error);
                end;
  end;
end;

{ existance d'éléments de rapport }
function TWRapport.Exists(TypeWRP: TTypeWRP): Boolean;
begin
  case TypeWRP of
    TWRP_Done : Result := TobDone.Detail.Count > 0;
    TWRP_Error: Result := TobError.Detail.Count > 0;
    TWRP_None : Result := (TobError.Detail.Count > 0) or (TobDone.Detail.Count > 0);
  else Result := false;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 28/04/2003
Modifié le ... :   /  /
Description .. : Liste d'exportation
Mots clefs ... :
*****************************************************************}
procedure TWRapport.ExportList;
begin
  PrintTobs(True);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thbaut SUBLET
Créé le ...... : 28/04/2003
Modifié le ... :   /  /
Description .. : Edition
Mots clefs ... :
*****************************************************************}
procedure TWRapport.Print;
begin
  PrintTobs(False);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 28/04/2003
Modifié le ... :   /  /
Description .. : Fusionne les 2 tobs en 1 pour l'édition
Mots clefs ... :
*****************************************************************}
procedure TWRapport.HappenTob(TobSource, TobDest: Tob);
var
  t: Tob;
  i: Integer;
begin
  fTobEtat.Dupliquer(TobDone, True, True);

  t := Tob.Create('_X_', nil, -1);
  try
    t.Dupliquer(TobSource, True, True);
    for i := Pred(t.Detail.Count) downto 0 do
      t.Detail[i].ChangeParent(TobDest, 0);
  finally
    t.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 28/04/2003
Modifié le ... :   /  /
Description .. : Dispatch:
Suite ........ :   - Liste d'export.
Suite ........ :   - Edition papier
Mots clefs ... :
*****************************************************************}
procedure TWRapport.PrintTobs(Liste: Boolean);
begin
  fTobEtat := Tob.Create('TOB_ETAT', nil, -1);
  try
    HappenTob(TobError, fTobEtat);
    LanceEtatTob('E', 'WD0', iif(not Liste, 'RP0', 'RP1'), fTobEtat, True, Liste, False, nil, '', fTitre, False);
  finally
    fTobEtat.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 28/04/2003
Modifié le ... :   /  /
Description .. : Enregistrer sur le disque
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TWRapport.SaveToFile(WithAskFileName: Boolean = True; FileName: String = '');
var
  Log, LogWRP: TStringList;
  i: Integer;
begin
  if not WithAskFileName or GetFileName(tfSave, '*.Log;*.*', FileName) then
  begin
    Log := TStringList.Create;
    LogWRP := TStringList.Create;
    try
      if FileExists(FileName) then
        Log.LoadFromFile(FileName);
      LogWRP.Add('// ' + FormatDateTime('dd/mm/yyyy hh:nn:ss', Now) + ' : ' + fTitre + ' \\');
      { log des actions effectuées }
      LogWRP.Add('[' + TraduireMemoire('Traitements effectués') + ']');
      for i := 0 to Pred(TobDone.Detail.Count) do
        LogWRP.Add(TobDone.Detail[i].GetString('IDENTIFIANTWXX') + ' : ' + TobDone.Detail[i].GetString('TEXTE'));
      { log des actions éhouées }
      LogWRP.Add('[' + TraduireMemoire('Traitements erronés') + ']');
      for i := 0 to Pred(TobError.Detail.Count) do
        LogWRP.Add(TobError.Detail[i].GetString('IDENTIFIANTWXX') + ' : ' + TobError.Detail[i].GetString('TEXTE'));
      { concaténation }
      Log.AddStrings(LogWRP);
      Log.SaveToFile(FileName);
    finally
      LogWRP.Free;
      Log.Free;
    end;
  end;
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 29/04/2003
Modifié le ... :   /  /
Description .. : Permet de passer le pointeur de l'objet TWRapport dans
Suite ........ : une tob.
Mots clefs ... :
*****************************************************************}
procedure TWRapport.AddObjectInTob(t: Tob);
begin
  if Assigned(t) then
    t.AddChampSupValeur(WRPTobFieldName, LongInt(Self));
end;

function TWRapport.DoneOn(Clef: Variant): Boolean;
begin
  Result := TobDone.G('IDENTIFIANTWXX') <> #0;
end;

function TWRapport.ErrorOn(Clef: Variant): Boolean;
begin
  Result := TobError.G('IDENTIFIANTWXX') <> #0;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 29/04/2003
Modifié le ... :   /  /
Description .. : Récupère un objet TWRapport contenu dans une tob.
Mots clefs ... :
*****************************************************************}
function wGetRapportFromTob(t: Tob): TWRapport;
begin
  if t.FieldExists(WRPTobFieldName) then
    Result := TWRapport(T.GetInteger(WRPTobFieldName))
  else
    Result := nil;
end;

function TWRapport.GetFirstError: String;
begin
  if CountError > 0 then
    Result := TobError.Detail[0].GetString('TEXTE')
  else
    Result := ''
end;

function TWRapport.GetFirstDone: String;
begin
  if CountDone > 0 then
    Result := TobDone.Detail[0].GetString('TEXTE')
  else
    Result := ''
end;

function TWRapport.CountError: Integer;
begin
  if Assigned(TobError) then
    Result := TobError.Detail.Count
  else
    Result := 0
end;

function TWRapport.CountDone: Integer;
begin
  if Assigned(TobDone) then
    Result := TobDone.Detail.Count
  else
    Result := 0
end;

function TWRapport.Count: Integer;
begin
  Result := CountDone + CountError
end;

end.





