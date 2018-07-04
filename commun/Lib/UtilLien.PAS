unit UtilLien;

interface

uses
  HCtrls,
  UtilPgi,
  HEnt1,
{$IFDEF EAGLSERVER}
  eSession,
{$ENDIF !EAGLSERVER}
  uTob,
  uTom,
  SysUtils;

type
  T_Lien = class
  private
    TobLien: Tob;
    TomLien: Tom;
    TobChamp: Tob;
    TobTable1, TobTable2: Tob;
    NomTable: string;
    nbChamps: integer;
    NumeroTable: integer;
    CorrespondanceMetierLien: string;
    CorrespondanceLienMetier: string;
    function ListeChampsValeurs(var ListeChamp, ListeValeur: string; bVal: boolean = false): string;
    function RempliValeurChamp(TobCle1, TobCle2: Tob; bSelect: boolean = false): boolean;
    function GetNomChamp(iChamp: integer): string;
    function GetTypeChamp(iChamp: integer): string;
  public
    function AjouteTob(TobCle1, TobCle2: Tob; bTable1: boolean = false; bTable2: boolean = false): boolean;
    function MAJTob(TobCleMAJ, TobCle: Tob): boolean;
    function DeleteTob(TobCle1, TobCle2: Tob): boolean;
    function ChargeTob(TobCle1, TobCle2: Tob; Init: boolean = false): boolean;
    procedure DupliqueTobLien(var TobDest: TOB);
    function GetWhereFromTob(LaTobLien: TOB; LePrefixe: string; MetierVersLien: boolean = false): string;
    function Update: boolean;
    function Insert: boolean;
    function Delete: boolean;
    procedure SetString(iTable: integer; Champ: string; Valeur: string);
    procedure SetInteger(iTable: integer; Champ: string; Valeur: Integer);
    procedure SetDateTime(iTable: integer; Champ: string; Valeur: TDateTime);
    constructor Create(LaTable: string);
    destructor Destroy; override;
  end;

implementation

function T_Lien.Update: boolean;
begin
  Result := false;
  if TomLien.VerifTOB(TobLien) then
  begin
    Result := TobLien.UpdateDB;
  end;
end;

function T_Lien.Insert: boolean;
begin
  Result := false;
  if TomLien.VerifTOB(TobLien) then
  begin
    Result := TobLien.InsertDB(nil);
  end;
end;

function T_Lien.Delete: boolean;
begin
  Result := false;
  if TomLien.DeleteTOB(TobLien) then
  begin
    Result := TobLien.DeleteDB;
  end;
end;

function T_Lien.RempliValeurChamp(TobCle1, TobCle2: Tob; bSelect: boolean = false): boolean;
var
  TypeChamp, NomChamp, NomChampMetier: string;
  TobCle: Tob;
  iInd: integer;
begin
  Result := true;
  if (TobCle1 <> nil) or (TobCle2 <> nil) then
  begin
    for iInd := 1 to nbchamps do
    begin
      NomChamp := GetNomChamp(iInd);
      NomChampMetier := GetArgumentString(CorrespondanceLienMetier, NomChamp);
      TypeChamp := GetTypeChamp(iInd);

      if (TobCle1 <> nil) and TobCle1.FieldExists(NomChampMetier) then
        TobCle := TobCle1
      else if (TobCle2 <> nil) and TobCle2.FieldExists(NomChampMetier) then
        TobCle := TobCle2
      else
        TobCle := nil;

      if TobCle <> nil then
      begin
        if TypeChamp = 'INTEGER' then
          TobChamp.SetInteger(NomChamp, TobCle.GetInteger(NomChampMetier))
        else if TypeChamp = 'DATE' then
          TobChamp.SetDateTime(NomChamp, TobCle.GetInteger(NomChampMetier))
        else
          TobChamp.SetString(NomChamp, TobCle.GetString(NomChampMetier));
      end else
      begin
        if TypeChamp = 'INTEGER' then
          TobChamp.SetInteger(NomChamp, -1)
        else if TypeChamp = 'DATE' then
        begin
          if bSelect then
            TobChamp.SetDateTime(NomChamp, iDate2099)
          else
            TobChamp.SetDateTime(NomChamp, iDate1900);
        end else
        begin
          if bSelect then
            TobChamp.SetString(NomChamp, '-1')
          else
            TobChamp.SetString(NomChamp, '');
        end;
      end;
    end;
  end else Result := false;
end;

function T_Lien.AjouteTob(TobCle1, TobCle2: Tob; bTable1: boolean = false; bTable2: boolean = false): boolean;
var
  TobLienAjout: Tob;
  TypeChamp, NomChamp: string;
  iInd: integer;
  TobLoc1, TobLoc2: TOB;
begin
  TobLienAjout := Tob.Create(NomTable, nil, -1);

  try
    Result := true;
    if Nomtable <> '' then
    begin
      if bTable1 then
        TobLoc1 := TobTable1
      else
        TobLoc1 := TobCle1;
      if bTable2 then
        TobLoc2 := TobTable2
      else
        TobLoc2 := TobCle2;
      if RempliValeurChamp(TobLoc1, TobLoc2) then
      begin
        for iInd := 1 to nbChamps do
        begin
          NomChamp := GetNomChamp(iInd);
          TypeChamp := GetTypeChamp(iInd);
          if TypeChamp = 'INTEGER' then
            TobLienAjout.SetInteger(NomChamp, TobChamp.GetInteger(NomChamp))
          else if TypeChamp = 'DATE' then
            TobLienAjout.SetDateTime(NomChamp, TobChamp.GetInteger(NomChamp))
          else
            TobLienAjout.SetString(NomChamp, TobChamp.GetString(NomChamp));
        end;
      end else
        Result := false;
    end else
      Result := false;

    if Result then
    begin
      Tob.Create(NomTable, TobLien, -1).Dupliquer(TobLienAjout, true, true, true);
    end;
  finally
    TobLienAjout.Free;
  end;
end;

function T_Lien.DeleteTob(TobCle1, TobCle2: Tob): boolean;
begin
  Result := true;
end;

function T_Lien.MAJTob(TobCleMAJ, TobCle: Tob): boolean;
var
  TobChampMAJ: Tob;

  procedure MajChamp;
  var
    TobLienTrouve: Tob;
    iInd: integer;
    ListeChamp, ListeValeur, TypeChamp, NomChamp: string;
  begin
    ListeChampsValeurs(ListeChamp, ListeValeur, true);
    TobLienTrouve := TobLien.FindFirst([ListeChamp], [ListeValeur], true);
    while TobLienTrouve <> nil do
    begin
      for iInd := 1 to nbChamps do
      begin
        NomChamp := GetNomChamp(iInd);
        TypeChamp := GetTypeChamp(iInd);
        if TypeChamp = 'INTEGER' then
          TobLienTrouve.SetInteger(NomChamp, TobChampMAJ.GetInteger(NomChamp))
        else if TypeChamp = 'DATE' then
          TobLienTrouve.SetDateTime(NomChamp, TobChampMAJ.GetInteger(NomChamp))
        else
          TobLienTrouve.SetString(NomChamp, TobChampMAJ.GetString(NomChamp));
      end;
      TobLienTrouve := TobLien.FindNext([ListeChamp], [ListeValeur], true);
    end;
  end;

begin
  Result := true;
  TobChampMAJ := Tob.Create(NomTable, nil, -1);
  if RempliValeurChamp(TobCleMAJ, nil) then
  begin
    TobChampMAJ.Dupliquer(TobChamp, true, true);
    if RempliValeurChamp(TobCle, nil) then
    begin
      MajChamp;
    end;
  end;
end;

function T_Lien.ChargeTob(TobCle1, TobCle2: Tob; Init: boolean = false): boolean;

  function GetSelect: string;
  var
    iInd: integer;
    NomChamp, TypeChamp, stChamp: string;
  begin
    Result := 'SELECT * FROM ' + NomTable + ' WHERE ';
    stChamp := '';
    for iInd := 1 to nbChamps do
    begin
      NomChamp := GetNomChamp(iInd);
      TypeChamp := GetTypeChamp(iInd);
      if TobChamp.GetInteger(NomChamp) <> -1 then
      begin
        if iInd > 1 then
          stChamp := stChamp + ' AND ';
        stChamp := stChamp + NomChamp;
        if TypeChamp = 'INTEGER' then
          stChamp := stChamp + '='
        else
          stChamp := stChamp + '="';
        stChamp := stChamp + TobChamp.GetString(NomChamp);
        if TypeChamp <> 'INTEGER' then
          stChamp := stChamp + '"';
      end;
    end;
    if stChamp <> '' then
      Result := Result + stChamp
    else
      Result := Result + '1=2';
  end;

begin
  Result := true;
  if Nomtable <> '' then
  begin
    if Init then TobChamp.InitValeurs;

    if RempliValeurChamp(TobCle1, TobCle2, true) then
      TobLien.LoadDetailFromSQL(GetSelect, true)
    else
      Result := false;
  end else
    Result := false;
end;

procedure T_Lien.DupliqueTobLien(var TobDest: TOB);
begin
  TobDest.ClearDetail;
  TobDest.Dupliquer(TobLien, true, true);
end;

function T_Lien.GetWhereFromTob(LaTobLien: TOB; LePrefixe: string; MetierVersLien: boolean = false): string;
var Cpt: integer;
  LaCorresp, NomChamp, NomChampInverse, TypeChamp: string;
begin
  Result := '';
  if (not assigned(LaTobLien)) or (LePrefixe = '') then exit;
  if not MetierVersLien then
    LaCorresp := CorrespondanceLienMetier
  else
    LaCorresp := CorrespondanceMetierLien;
  for Cpt := 1 to nbChamps do
  begin
    NomChamp := GetNomChamp(Cpt);
    NomChampInverse := GetArgumentString(LaCorresp, NomChamp);
    if copy(NomChampInverse, 1, length(LePrefixe)) <> LePrefixe then
      continue;
    TypeChamp := GetTypeChamp(Cpt);
    if Result <> '' then
      Result := Result + ' AND ';
    if TypeChamp = 'INTEGER' then
      Result := Result + NomChampInverse + ' = ' + IntToStr(LaTobLien.GetInteger(NomChamp)) + ' '
    else if TypeChamp = 'DATE' then
      Result := Result + NomChampInverse + ' = "' + USDateTime(LaTobLien.GetDateTime(NomChamp)) + '" '
    else
      Result := Result + NomChampInverse + ' = "' + LaTobLien.GetString(NomChamp) + '" ';
  end;
end;

function T_Lien.ListeChampsValeurs(var ListeChamp, ListeValeur: string; bVal: boolean = false): string;
var
  iInd: integer;
  NomChamp: string;
begin
  Result := '';
  for iInd := 1 to nbChamps do
  begin
    if iInd > 1 then
    begin
      ListeChamp := ListeChamp + ', ';
      ListeValeur := ListeValeur + ', ';
    end;
    NomChamp := GetNomChamp(iInd);
    ListeChamp := ListeChamp + NomChamp;
    ListeValeur := ListeValeur + TobChamp.GetString(NomChamp);
  end;
end;

function T_Lien.GetNomChamp(iChamp: integer): string;
begin
  Result := V_Pgi.DeChamps[NumeroTable, iChamp].Nom;
end;

function T_Lien.GetTypeChamp(iChamp: integer): string;
begin
  Result := V_Pgi.DeChamps[NumeroTable, iChamp].Tipe;
end;

procedure T_Lien.SetString(iTable: integer; Champ: string; Valeur: string);
var
  TobTable: Tob;
begin
  if iTable = 1 then
    TobTable := TobTable1
  else
    TobTable := TobTable2;
  if TobTable.FieldExists(Champ) then
    TobTable.SetString(Champ, valeur);
end;

procedure T_Lien.SetInteger(iTable: integer; Champ: string; Valeur: Integer);
var
  TobTable: Tob;
begin
  if iTable = 1 then TobTable := TobTable1 else TobTable := TobTable2;

  if TobTable.FieldExists(Champ) then
    TobTable.SetInteger(Champ, valeur);
end;

procedure T_Lien.SetDateTime(iTable: integer; Champ: string; Valeur: TDateTime);
var
  TobTable: Tob;
begin
  if iTable = 1 then TobTable := TobTable1 else TobTable := TobTable2;

  if TobTable.FieldExists(Champ) then
    TobTable.SetDateTime(Champ, valeur);
end;

constructor T_Lien.Create(LaTable: string);
//  BBI Web Services
{$IFDEF EAGLSERVER}
var LaSession: TISession;
{$ENDIF !EAGLSERVER}
//  BBI Fin Web Services
begin
  NomTable := LaTable;
  TomLien := CreateTOM(NomTable, nil, false, true);
  TobLien := Tob.Create('LES LIENS', nil, -1);
  TobChamp := Tob.Create(NomTable, nil, -1);
  TobTable1 := Tob.Create('', nil, -1);
  TobTable2 := Tob.Create('', nil, -1);

//  BBI Web Services
{$IFNDEF EAGLSERVER}
  nbChamps := High(V_Pgi.DEChamps[TableToNum(NomTable)]);
{$ELSE  !EAGLSERVER}
  LaSession := LookupCurrentSession;
  nbChamps := High(V_Pgi.HDEChamps[LaSession.SocNum, TableToNum(NomTable)]);
{$ENDIF !EAGLSERVER}
//  BBI Fin Web Services
  NumeroTable := TableToNum(NomTable);
  if NomTable = 'YLIGECR' then
  begin
    TobTable1.AddChampSup('GL_NATUREPIECEG', false);
    TobTable1.AddChampSup('GL_SOUCHE', false);
    TobTable1.AddChampSup('GL_INDICEG', false);
    TobTable1.AddChampSup('GL_NUMERO', false);
    TobTable1.AddChampSup('GL_NUMORDRE', false);
    TobTable2.AddChampSup('E_JOURNAL', false);
    TobTable2.AddChampSup('E_EXERCICE', false);
    TobTable2.AddChampSup('E_DATECOMPTABLE', false);
    TobTable2.AddChampSup('E_NUMEROPIECE', false);
    TobTable2.AddChampSup('E_NUMLIGNE', false);
    TobTable2.AddChampSup('E_NUMECHE', false);
    TobTable2.AddChampSup('E_QUALIFPIECE', false);
    CorrespondanceMetierLien :=
      'GL_NATUREPIECEG=YLE_NAL;GL_SOUCHE=YLE_SL;GL_NUMERO=YLE_NL;GL_INDICEG=YLE_IL;GL_NUMORDRE=YLE_NOL;' +
      'E_JOURNAL=YLE_J;E_EXERCICE=YLE_E;E_DATECOMPTABLE=YLE_DC;E_NUMEROPIECE=YLE_NP;E_NUMLIGNE=YLE_NUL;E_NUMECHE=YLE_NE;E_QUALIFPIECE=YLE_Q';
    CorrespondanceLienMetier :=
      'YLE_NAL=GL_NATUREPIECEG;YLE_SL=GL_SOUCHE;YLE_NL=GL_NUMERO;YLE_IL=GL_INDICEG;YLE_NOL=GL_NUMORDRE;' +
      'YLE_J=E_JOURNAL;YLE_E=E_EXERCICE;YLE_DC=E_DATECOMPTABLE;YLE_NP=E_NUMEROPIECE;YLE_NUL=E_NUMLIGNE;YLE_NE=E_NUMECHE;YLE_Q=E_QUALIFPIECE';
  end;
end;

destructor T_Lien.Destroy;
begin
  TomLien.Free;
  TobLien.Free;
  TobChamp.Free;
  TobTable1.Free;
  TobTable2.Free;
end;

end.
