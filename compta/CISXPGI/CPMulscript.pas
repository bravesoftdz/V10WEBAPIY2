unit CPMulscript;

interface

uses
  Windows, Classes, Controls, SysUtils,
{$IFDEF EAGLCLIENT}
     MaineAGL, UtileAgl, eMul, UScriptTob,
{$ELSE}
     db, Mul, HDB,
  {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
     FE_main,
{$ENDIF}
  StdCtrls, Hctrls, Dialogs, UTob,
  hmsgbox, HEnt1, UControlParam,
  UTOF, UPDomaine, UAssistConcept, UScript,
  HSysMenu, uYFILESTD, UBob, HQry;


Type
  TOF_MULSCRIPT = Class (TOF)
       procedure FListeDblClick(Sender: TObject);
       procedure OnArgument (S : String ) ; override ;
       procedure InsertClick(Sender: TObject);
       procedure DeleteClick(Sender: TObject);
       private
       ModeSelect   : Boolean;
       {$IFDEF EAGLCLIENT}
       FListe : THGrid;
       {$ELSE}
       FListe : THDBGrid;
       {$ENDIF}
       procedure BSauveClick(Sender: TObject);
       {$IFNDEF EAGLCLIENT}
       procedure InsereDansTobTable (TobGen : TOB; Q : TQuery);
       {$endif}
       procedure BobAddTableStandard(TStd: TOB; stSQL: string);
  end ;

implementation

uses UConcept;

procedure TOF_MULSCRIPT.FListeDblClick(Sender: TObject);
var
  Table, Dm, Nat: string;
  lib,Nt,edt    : string;
  Delim         : Boolean;
  TPControle    : TSVControle;
begin
{$IFDEF EAGLCLIENT} // positionnement de Query
 TFMul(Ecran).Q.TQ.Seek(THGrid(TFMul(Ecran).FListe).Row - 1);
{$ENDIF}
  Table := TFMul(Ecran).Q.FindField('CIS_CLE').AsString;
  Dm    := TFMul(Ecran).Q.FindField('CIS_Domaine').AsString;
  if ModeSelect then
  begin
       TFMul(Ecran).ModalResult := MrOk;
       TFMul(Ecran).retour:= Table + ';'+ Dm;
       exit;
  end;
  Nat   := TFMul(Ecran).Q.FindField('CIS_CLEPAR').AsString;
  lib   := RendLibelleDomaine(Dm);
  Nt    :=  TFMul(Ecran).Q.FindField('CIS_Table2').AsString;
  Edt   := TFMul(Ecran).Q.FindField('CIS_Table1').AsString;
  Delim := (TFMul(Ecran).Q.FindField('CIS_Nature').AsString = 'X');
  if Delim then
     ModifScriptDelim(Table, Nat, Nt, Edt, Dm)
  else
  begin
      TPControle := TSVControle.create;
      TPControle.LePays := TFMul(Ecran).Q.FindField('CIS_Table3').AsString;
      TPControle.ChargeTobParam(Dm);
      PExtConcept(nil, Lib, Table, Nat, Nt, Edt, taModif, TPControle);
  end;
end;

procedure TOF_MULSCRIPT.InsertClick(Sender: TObject);
begin
  CreateScript (FALSE, tacreat, '', TRUE, THEdit(GetControl('CIS_CLEPAR')).Text, THValComboBox(GetControl('CIS_TABLE3')).value);
  TFMul(Ecran).FListe.refresh;
end;

procedure TOF_MULSCRIPT.DeleteClick(Sender: TObject);
var
{$IFDEF EAGLCLIENT}
    L : THGrid;
    wBookMark: integer;
{$ELSE}
    L : THDBGrid;
    wBookMark: tBookMark;
{$ENDIF}
    i : integer;
    Q : TQuery;
begin
    L := TFMul(Ecran).FListe;
    if (L = Nil) then exit;
{$IFDEF EAGLCLIENT}
    Q := TFMul(Ecran).Q.TQ;
    Q.Seek(L.Row-1);
    wBookMark := FListe.Row;
{$ELSE}
    Q := TFMul(Longint( Ecran )).Q;
    wBookMark := FListe.DataSource.DataSet.GetBookmark;
{$ENDIF}
    // si rien de selectionné !
    if ((L.NbSelected = 0) and (not L.AllSelected)) then
    begin
        MessageAlerte('Aucun élément sélectionné.');
        exit;
    end;
    if (L.AllSelected) then // si tout ??
    begin
        Q.First;

        while (Not Q.EOF) do
        begin
            executeSql('DELETE FROM CPGZIMPREQ WHERE CIS_CLE="' + Q.FindField('CIS_CLE').asstring + '"');
            Q.Next;
        end;

        L.AllSelected:=False;
    end
    else // si pas tout ....
    begin
        for i := 0 to L.NbSelected-1 do
        begin
            L.GotoLeBookMark(i);
            {$IFDEF EAGLCLIENT}
            TFMul(Ecran).Q.TQ.Seek (L.Row-1);
            {$ENDIF}
            executeSql('DELETE FROM CPGZIMPREQ WHERE CIS_CLE="' + Q.FindField('CIS_CLE').asstring + '"');
        end;

        L.ClearSelected;
    end;

{$IFDEF EAGLCLIENT}
      Q := TFMul(Ecran).Q.TQ;
      FListe.Row := wBookMark;
      if FListe.Row >0 then
        Q.Seek( FListe.Row )
      else
        Q.First;
{$ELSE}
     Q.GotoBookmark(wBookMark);
{$ENDIF}

  TFMul(Ecran).BChercheClick(nil);
end;


procedure TOF_MULSCRIPT.OnArgument (S : String );
var
St : string;
begin
TFMul(Ecran).FListe.OnDblClick := FListeDblClick;
TFMul(Ecran).BInsert.OnClick := InsertClick;
TButton(GetControl('BDelete',True)).onClick  := DeleteClick;
TButton(GetControl('BSAUVE',True)).onClick  := BSauveClick;

ModeSelect := (S ='TRUE');
St := ReadTokenPipe(S, '=');
if S <> '' then
begin
         St := ReadTokenSt(S);
         if St <>'' then
         THEdit(GetControl('CIS_CLEPAR')).Text := St;
         St := ReadTokenPipe (S, '=');
         if S <> '' then
            THValComboBox(GetControl('CIS_TABLE3')).Itemindex := THValComboBox(GetControl('CIS_TABLE3')).Values.IndexOf(S);

         if not V_PGI.SAV then
         begin
              SetControlEnabled ('CIS_CLEPAR', FALSE);
              if (GetControlText('CIS_CLEPAR') = 'EXPORT') and
              (GetControlText('CIS_TABLE3') <> '' ) then SetControlEnabled ('CIS_TABLE3', FALSE);
         end;
end;
if (GetInfoVHCX.Domaine <> '') then
begin
         THValComboBox(GetControl('CIS_DOMAINE')).value := GetInfoVHCX.Domaine;
{$IFDEF COMPTA}
         SetControlEnabled ('CIS_DOMAINE', FALSE);
{$ENDIF}

{$IFDEF EAGLCLIENT}
    FListe := THGrid(GetControl('FListe'));
    TFMul(Ecran).HMTrad.ResizeGridColumns(FListe);
{$ELSE}
    FListe := THDBGrid (GetControl('FListe'));
    CentreDBGrid(FListe);
    THSystemMenu(GetControl('HMTrad')).ResizeDBGridColumns(FListe);
{$ENDIF}


end;

end;

procedure TOF_MULSCRIPT.BobAddTableStandard(TStd: TOB; stSQL: string);
begin
  with TOB.Create('',TStd,-1) do
  begin
    AddChampSupValeur('_OBJECTTYPE', toSQL);
    AddChampSupValeur('_OBJECTNAME', stSQL);
    AddChampSupValeur('_OBJECTWITHDATA', False );
    AddChampSupValeur('_OBJECTDOMAINE','C');
  end;
end;

procedure TOF_MULSCRIPT.BSauveClick(Sender: TObject);
var
{$IFDEF EAGLCLIENT}
    L          : THGrid;
{$ELSE}
    L          : THDBGrid;
    s          : TMemoryStream;
{$ENDIF}
    i          : integer;
    FName, Fna : string;
    SD         : TSaveDialog;
    TobSave    : TOB;
    Q          : TQuery;
    F          : TFMul;
    Operation  : string;
    Extend     : string;
    procedure TraiteLaSauvegarde (FF : string);
    var
    Res        : integer;
    FilePath   : string;
    begin
        Case Operation[1] of
        'P' : // extraction de pdf
            begin
                Res :=  AGL_YFILESTD_EXTRACT (FilePath, 'COMPTA', FF, 'ECH', 'MOP', '', '', '', FALSE, V_PGI.LanguePrinc, 'CEG' );
                if Res <> -1 then
                  AGL_YFILESTD_EXTRACT (FilePath, 'COMPTA', FF, 'ECH', 'MOP', '', '', '', FALSE, V_PGI.LanguePrinc, 'DOS' );
                DeleteFile( Pchar(FName));
                Copyfile (PChar( FilePath), Pchar(FName),TRUE);
            end;
        {$IFNDEF EAGLCLIENT}
        'S' :  InsereDansTobTable (TobSave, Q);  // sauvegarde cix
        {$endif}
        'B' :  // sauvegarde en bob
            begin
              BobAddTableStandard ( TobSave, 'SELECT * FROM CPGZIMPREQ Where CIS_CLE="'+Q.FindField ('CIS_CLE').asstring+'"');
              BobAddTableStandard ( TobSave, 'SELECT * FROM YFILESTD WHERE YFS_CODEPRODUIT="COMPTA" AND YFS_CRIT1="ECH" AND YFS_NOM= "'+Q.FindField ('CIS_CLE').asstring+'.PDF"');
              BobAddTableStandard ( TobSave, 'SELECT * FROM NFILES WHERE NFI_FILEGUID IN (SELECT YFS_FILEGUID FROM YFILESTD WHERE YFS_CODEPRODUIT="COMPTA" AND YFS_CRIT1="ECH" AND YFS_NOM= "'+Q.FindField ('CIS_CLE').asstring+'")');
              BobAddTableStandard ( TobSave, 'SELECT * FROM NFILEPARTS WHERE NFS_FILEGUID  IN (SELECT YFS_FILEGUID FROM YFILESTD WHERE YFS_CODEPRODUIT="COMPTA" AND YFS_CRIT1="ECH" AND YFS_NOM= "'+Q.FindField ('CIS_CLE').asstring+'")');
            end;
        end;
    end;
begin
    L := TFMul(Ecran).FListe;
    F := TFMul(Longint( Ecran ));
    if (F = Nil) then exit;

    if (L = Nil) then exit;
{$IFNDEF EAGLCLIENT}
    Q := F.Q;
{$endif}
    // si rien de selectionné !
    if ((L.NbSelected = 0) and (not L.AllSelected)) then
    begin
        MessageAlerte('Aucun élément sélectionné.');
        exit;
    end;

    // message d'avertisement !
    SD := TSaveDialog.Create(nil);
    SD.DefaultExt := '*.xls';
    SD.Title := 'Extraction des rapports';
    SD.Filter := 'Fichier Texte (*.cix)|*.cix|Tous les fichiers (*.pdf)|*.pdf|Tous les fichiers (*.bob)|*.bob';
    if SD.execute then
      Fna := SD.FileName;
    SD.free;
    if pos ('.cix', Fna) <> 0 then
    begin
      Extend := '.cix';
      FName := ReadTokenPipe (Fna, '.')+ Extend;
      Operation := 'S';
    end
    else
    if pos ('.pdf', Fna) <> 0 then
    begin
      Extend := '.pdf';
      FName := ReadTokenPipe (Fna, '.')+ Extend;
      Operation := 'P';
    end;
    if pos ('.bob', Fna) <> 0 then
    begin
      Extend := '.bob';
      FName := ReadTokenPipe (Fna, '.')+ Extend;
      Operation := 'B';
    end;
    if Operation = '' then exit;

    TobSave := TOB.Create('Scripts', nil, -1);
    if Operation = 'B' then
    begin
      TobSave.AddChampSupValeur('BOBNAME',FName);
      TobSave.AddChampSupValeur('BOBVERSION', '01');
      TobSave.AddChampSupValeur('BOBNUMSOCREF',V_PGI.NumVersionBase);
      TobSave.AddChampSupValeur('BOBDATEGEN',date);
      TobSave.AddChampSupValeur('BOBSOCREFDEPART', 750);
      TobSave.AddChampSupValeur('BOBSOCREFFINALE', 999);
      TobSave.AddChampSupValeur('BOBTYPE', tbNormalControl);
    end;

    // destruction
    if (L.AllSelected) then // si tout ??
    begin
        Q.First;
        while (Not Q.EOF) do
        begin
            TraiteLaSauvegarde(Q.FindField ('CIS_CLE').asstring+Extend);
            Q.Next;
        end;
        L.AllSelected:=False;
    end
    else // si pas tout ....
    begin
        for i := 0 to L.NbSelected-1 do
        begin
            L.GotoLeBookMark(i);
            {$IFDEF EAGLCLIENT}
            TFMul(Ecran).Q.TQ.Seek (L.Row-1);
            {$ENDIF}
            TraiteLaSauvegarde (Q.FindField ('CIS_CLE').asstring+Extend);
        end;

        L.ClearSelected;
    end;
    if Operation = 'S' then
        TobSave.SaveToXMLFile(FName, True, True);
    if Operation = 'B' then
        AglExportBob (FName, False, False, TobSave, True);
    TobSave.free;

end;
{$IFNDEF EAGLCLIENT}
procedure TOF_MULSCRIPT.InsereDansTobTable (TobGen : TOB; Q : TQuery);
var ind1                                   : integer;
    TobL, TobChamp                         : TOB;
    FNumero, FLocDomaine, FLocTable, FNom, FNomChamp : TStringField;
    stBlob                                 : string;
    TDelim                                 : TTable;
    NomField,Tmp                           : string;
    TBase                                  : TQuery;
begin
TobL := nil;
    TobChamp := nil;
    TBase := OpenSQl ('SELECT * FROM CPGZIMPREQ Where CIS_CLE="'+Q.FindField ('CIS_CLE').asstring+'"', TRUE);
    if not Q.EOF then
    for ind1 := 0 to TBase.FieldCount - 1 do
    begin
        NomField := TBase.Fields[ind1].FieldName;
        Tmp := ReadTokenPipe(NomField,'CIS_');
        if ind1 = 0 then
            begin
            TobL := TOB.Create(TBase.Fields[ind1].AsString, TobGen, -1);
            TobL.AddChampSupValeur('NomTable', TBase.Fields[ind1].AsString, False);
            end;
        if TBase.Fields[ind1].FieldName = 'CIS_DOMAINE' then
            TobL.AddChampSupValeur('LeDomaine' , TBase.Fields[ind1].AsString, False);
        if TBase.Fields[ind1].FieldName = 'CIS_CLE' then
            TobL.AddChampSupValeur('LaTable'   , TBase.Fields[ind1].AsString, False);
        if (TBase.Fields[ind1].DataType = ftBlob) or (TBase.Fields[ind1].DataType =  ftMemo) then
            begin
            stBlob := BlobToString(TBlobField(TBase.Fields[ind1]));
            TobL.AddChampSupValeur(NomField, stBlob, False);
            end
            else
        if TBase.Fields[ind1].FieldName = 'CIS_DATEMODIF' then
            TobL.AddChampSupValeur('DATEDEMODIF', Copy(TBase.Fields[ind1].AsString,0,10), False)
            else
            TobL.AddChampSupValeur(NomField, TBase.Fields[ind1].AsString, False);
        end;
        if TobChamp <> nil then
        begin
            for ind1 := 0 to TDelim.FieldCount - 1 do
                TobL.AddChampSupValeur('Delim-'+TDelim.Fields[ind1].FieldName,
                                       TobChamp.GetValue(TDelim.Fields[ind1].FieldName), False);
    end;
    ferme(TBase)
end;
{$endif}

initialization
  registerclasses([TOF_MULSCRIPT]);

end.
