{***********UNITE*************************************************
Auteur  ...... : JT       
Créé le ...... : 10/07/2003
Modifié le ... :   /  /
Description .. : Utilitaires pour les TOB
Mots clefs ... : TOB;UTILITAIRES
*****************************************************************}
unit UtilUtilitaires;

interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} EdtREtat,
{$ELSE}
     UtileAGL,  
{$ENDIF}
     Utob,EntGc,forms,sysutils,ComCtrls,HCtrls, CBPPath,
     HEnt1,HMsgBox,SaisieTexteValeur,ED_Tools;

function UtilTobCreat(TobMere : TOB; Tri,TitreColonnes,DataPrint,DataExport : string) : TOB;
procedure UtilTobSort(TobMere : TOB);
procedure UtilTobPrint(TobToPrint : TOB; TitreEtat : string; FormatEtat : integer = 0);
procedure UtilTobExportTxt(TobToExport : TOB; NomFichier : string);
function UtilTobFindMaxValueStr(TobMere : TOB; FindChp, NomChps, ValeurChps : string) : string;
function UtilTobFindMaxValueInt(TobMere : TOB; FindChp, NomChps, ValeurChps : string) : Integer;

implementation

function UtilTobCreat(TobMere : TOB; Tri,TitreColonnes,DataPrint,DataExport : string) : TOB;
var TobTmp : TOB;
    QteTri : integer;
begin
  //TobMere = Tob dans laquelle il faut créer les filles
  //Tri = Valeur sur lesquelles il faut trier (no limite) sous forme ReadTokenSt
  //TitreColonnes = Titre des colonnes
  //DataPrint = string formaté pour l'impression
  //DataExport = sous forme ReadTokenSt (exemple GA_CODEARTICLE=10300;GA_LIBELLE=Sachet ...)
  Result := nil;
  if TobMere = nil then exit;
  QteTri := 0;
  TobTmp := TOB.Create('UtilEnreg',TobMere,-1);
  While Tri <> '' do
  begin
    QteTri := QteTri + 1;
    TobTmp.AddChampSupValeur('U_TRI'+IntToStr(QteTri),ReadTokenSt(Tri),False);
  end;
  TobTmp.AddChampSupValeur('U_TITRE',TitreColonnes,False);
  TobTmp.AddChampSupValeur('U_ENREG', DataPrint, False);
  TobTmp.AddChampSupValeur('U_EXPORT', DataExport, False);
  if not TobMere.FieldExists('QTETRI') then
    TobMere.AddChampSupValeur('QTETRI',QteTri,False);
  Result := TobMere;
end;

procedure UtilTobSort(TobMere : TOB);
var QteTri : integer;
    ChpsTri : string;
begin
  ChpsTri := '';
  for QteTri := 1 to TobMere.GetValue('QTETRI') do
    ChpsTri := ChpsTri+';U_TRI'+IntToStr(QteTri);
  ChpsTri := copy(ChpsTri,2,length(ChpsTri));
  TobMere.Detail.Sort(ChpsTri);
end;

procedure UtilTobPrint(TobToPrint : TOB; TitreEtat : string; FormatEtat : integer = 0);
begin
  //FormatEtat : 0 = paysage, 1 = portrait
  if TobToPrint = nil then exit;
  if TobToPrint.Detail.count = 0 then
  begin
    PGIInfo(TraduireMemoire('Il n''y a pas de lignes à imprimer.'),TraduireMemoire('INFORMATIONS'))
  end else
  begin
    InitMoveProgressForm(nil,TraduireMemoire('Traitement'),TraduireMemoire('Impression en cours.'),1,False,False);
    MoveCurProgressForm('');
    LanceEtatTob('E','UTI','UT'+IntToStr(FormatEtat),TobToPrint,True,False,False,nil,'',TitreEtat,False);
    FiniMoveProgressForm;
  end;
end;

procedure UtilTobExportTxt(TobToExport : TOB; NomFichier : string);
var TobExport,TobTmp : TOB;
    Cpt : integer;
    LigneExport,ChampExtrait,ChampNom,ChampValeur : string;
    FSaisie : TFSaisieTexteValeur;
begin

  if TobToExport = nil then exit;

  if TobToExport.Detail.count = 0 then
     begin
     PGIInfo(TraduireMemoire('Il n''y a pas de lignes à exporter.'),TraduireMemoire('INFORMATIONS'))
     end
  else
     begin
     FSaisie := TFSaisieTexteValeur.Create(Application);
     if NomFichier <> '' then
        //FSaisie.Saisie.Text := 'C:\PGI00\'+ NomFichier + '.txt'
        FSaisie.Saisie.Text := IncludeTrailingBackSlash(TCBPPath.GetCegidDataDistri)+ NomFichier + '.txt'
     else
        FSaisie.Saisie.Text := '';
     FSaisie.Libelle.Caption := TraduireMemoire('Emplacement du fichier');
     FSaisie.ShowModal;
     if FSaisie.Saisie.Text <> '' then
        begin
        InitMoveProgressForm(nil,TraduireMemoire('Traitement'),TraduireMemoire('Impression en cours.'),TobToExport.Detail.count,False,False);
        TobExport := TOB.Create('',nil,-1);
        for Cpt := 0 to TobToExport.Detail.count -1 do
            begin
            MoveCurProgressForm('');
            TobTmp := TOB.Create('',TobExport,-1);
            LigneExport := TobToExport.Detail[Cpt].GetValue('U_EXPORT');
            while LigneExport <> '' do
                  begin
                  ChampExtrait := ReadTokenSt(LigneExport);
                  ChampNom := Copy(ChampExtrait,1,pos('=',ChampExtrait)-1);
                  ChampValeur := Copy(ChampExtrait,pos('=',ChampExtrait)+1,length(ChampExtrait));
                  TobTmp.AddChampSupValeur(ChampNom,ChampValeur,False);
            end;
        end;
        FiniMoveProgressForm;
        TobExport.SaveToFile(FSaisie.Saisie.Text,False, True, True);
        TobExport.Free;
     end;
     FSaisie.free;
  end;

end;

function UtilTobFindMaxValueStr(TobMere : TOB; FindChp, NomChps, ValeurChps : string) : string;
var TobTmp : TOB;
    Cpt : integer;
begin
  Result := '';
  if TobMere = nil then exit;
  TobTmp := nil;
  if NomChps <> '' then
  begin
    TobTmp := TobMere.FindFirst([NomChps], [ValeurChps], True);
    while TobTmp <> nil do
    begin
      if TobTmp.GetValue(FindChp) > Result then
        Result := TobTmp.GetValue(FindChp);
      TobTmp := TobMere.FindNext([NomChps], [ValeurChps], True);
    end;
  end else
  begin
    for Cpt := 0 to TobMere.Detail.count - 1 do
    begin
      if TobMere.Detail[Cpt].GetValue(FindChp) > Result then
        Result := TobTmp.GetValue(FindChp);
    end;
  end;
end;

function UtilTobFindMaxValueInt(TobMere : TOB; FindChp, NomChps, ValeurChps : string) : integer;
var TobTmp : TOB;
    Cpt : integer;
begin
  Result := 0;
  if TobMere = nil then exit;
  TobTmp := nil;
  if NomChps <> '' then
  begin
    TobTmp := TobMere.FindFirst([NomChps], [ValeurChps], True);
    while TobTmp <> nil do
    begin
      if TobTmp.GetValue(FindChp) > Result then
        Result := TobTmp.GetValue(FindChp);
      TobTmp := TobMere.FindNext([NomChps], [ValeurChps], True);
    end;
  end else
  begin
    for Cpt := 0 to TobMere.Detail.count - 1 do
    begin
      if TobMere.Detail[Cpt].GetValue(FindChp) > Result then
        Result := TobTmp.GetValue(FindChp);
    end;
  end;
end;

end.
