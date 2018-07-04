{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFREVIMPORTINDICE ()
Mots clefs ... : TOF;AFREVIMPORTINDICE
*****************************************************************}
unit utofafRevImportIndice;

interface

uses StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fe_Main,
  {$ELSE}
  MainEagl,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF, UTob, HTB97, paramsoc, dicobtp;

type
  TOF_AFREVIMPORTINDICE = class(TOF)

    BTNValider: TToolbarButton97;
    LaGrille: THGrid;
    fTobIndices : Tob;
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
    procedure LaGrilleClick(sender: Tobject);
    procedure BTNValiderClick(sender: Tobject);
    function LoadTob(T: TOB): integer;
//    procedure SIMULEIMPORT;
    procedure integre_indice(lecode: string; indice: integer);
    procedure integre_valeur_indice(lecode, pubcode, inddateval: string; indice: integer);

  public
    CHEMIN: string;
    procedure RempliGrilleparTob;
  end;

procedure AFLanceFiche_AFREVIMPORTINDICE(chemin: string);

implementation

procedure TOF_AFREVIMPORTINDICE.RempliGrilleparTob;
var j, i: integer;
  old_AIN_INDCODE: string;
begin
  i := 0;
  j := 0;
  old_AIN_INDCODE := '';
  while (i < fTobIndices.Detail.count) do
  begin
    if fTobIndices.Detail[i].getvalue('AIN_INDCODE') <> old_AIN_INDCODE then
    begin
      LaGrille.Cells[0, j + 1] := fTobIndices.Detail[i].getvalue('AIN_INDCODE');
      LaGrille.Cells[1, j + 1] := 'Oui';
      inc(j);                          
    end;
    old_AIN_INDCODE := fTobIndices.Detail[i].getvalue('AIN_INDCODE');
    inc(i);
  end;
 // LaGrille.rowcount := LaGrille.rowcount - 1; // j'neleve la derniere ligne
  LaGrille.rowcount := j + 1; // j'neleve la derniere ligne 
end;

procedure TOF_AFREVIMPORTINDICE.LaGrilleClick(sender: Tobject);
begin
end;

function Decodeimport(St: string): boolean;
begin
  if (St = 'Oui') or (St = 'X') then Result := true else Result := false;
end;

procedure TOF_AFREVIMPORTINDICE.integre_indice(lecode: string; indice: integer);
var st : string;

begin
  // je vais lire la tob est inserer ou update l'indice
  if ExisteSQL('select AIN_INDCODE from afindice where ain_indcode="' + lecode + '"') then
  begin // update
    st := 'update afindice set ';
    st := St + 'AIN_INDLIBELLE="' + fTobIndices.Detail[indice].getvalue('AIN_INDLIBELLE') + '",';
    st := St + 'AIN_INDMAJ = "IMP",';
    st := St + 'AIN_INDDATECREA="' + usDateTime(strtodate(fTobIndices.Detail[indice].getvalue('AIN_INDDATECREA_STRING'))) + '",';
    st := St + 'AIN_INDDATEFIN="' + usDateTime(strtodate(fTobIndices.Detail[indice].getvalue('AIN_INDDATEFIN_STRING'))) + '",';
    st := St + 'AIN_DATEMAJ="' + usDateTime(strtodate(fTobIndices.Detail[indice].getvalue('AIN_DATEMAJ_STRING'))) + '"';
    st := St + 'where AIN_INDCODE="' + lecode + '"';
  end                         
  else
  begin // insert
    st := 'insert into afindice (';
    st := st + 'AIN_INDCODE,AIN_INDLIBELLE,AIN_INDDATECREA,AIN_INDDATEFIN,AIN_INDMAJ,AIN_DATEMAJ)  values (';
    st := St + '"' + fTobIndices.Detail[indice].getvalue('AIN_INDCODE') + '","' + fTobIndices.Detail[indice].getvalue('AIN_INDLIBELLE') + '","';
    st := St + usDateTime(strtodate(fTobIndices.Detail[indice].getvalue('AIN_INDDATECREA_STRING'))) + '","';
    st := St + usDateTime(strtodate(fTobIndices.Detail[indice].getvalue('AIN_INDDATEFIN_STRING'))) + '",';
    st := St + '"IMP","';
    st := St + usDateTime(strtodate(fTobIndices.Detail[indice].getvalue('AIN_DATEMAJ_STRING'))) + '")';
  end;
  ExecuteSQL(st);
end;

procedure TOF_AFREVIMPORTINDICE.integre_valeur_indice(lecode, pubcode, inddateval: string; indice: integer);
var st: string;
begin
  if ExisteSQL('select AFV_INDCODE from afvalindice where AFV_INDCODE="' + lecode + '" and AFV_pubcode="' + pubcode + '" and AFV_INDDATEVAL="' + usDateTime(strtodate(inddateval)) +
    '"') then
  begin // update
    //  update afvalindice set afv_pubcode='MOD',AFV_indcode='GC',AFV_inddateval='30/12/2001' where
    // afv_pubcode='MOD' and AFV_indcode='GC' and AFV_inddateval='30/12/2001'
    st := 'update afvalindice set ';
    st := St + 'AFV_DEFINITIF="' + fTobIndices.Detail[indice].getvalue('AFV_DEFINITIF') + '",';
    st := St + 'AFV_CREATEUR="' + fTobIndices.Detail[indice].getvalue('AFV_CREATEUR') + '",';
    st := St + 'AFV_INDCOMMENT="' + fTobIndices.Detail[indice].getvalue('AFV_INDCOMMENT') + '",';
    st := St + 'AFV_INDNUMPUB="' + fTobIndices.Detail[indice].getvalue('AFV_INDNUMPUB') + '",';
    st := St + 'AFV_INDVALEUR=' + variantToSql(fTobIndices.Detail[indice].getvalue('AFV_INDVALEUR')) + ',';
    st := St + 'AFV_INDCODESUIV="' + fTobIndices.Detail[indice].getvalue('AFV_INDCODESUIV') + '",';
    st := St + 'AFV_PUBCODESUIV="' + fTobIndices.Detail[indice].getvalue('AFV_PUBCODESUIV') + '",';
    st := St + 'AFV_INDMAJ = "IMP",';
    st := St + 'AFV_COEFRACCORD=' + variantToSql(fTobIndices.Detail[indice].getvalue('AFV_COEFRACCORD')) + ',';
    st := St + 'AFV_COEFPASSAGE=' + variantToSql(fTobIndices.Detail[indice].getvalue('AFV_COEFPASSAGE')) + ',';
    st := St + 'AFV_DATECREATION="' + usDateTime(strtodate(fTobIndices.Detail[indice].getvalue('AFV_DATECREATION_STRING'))) + '",';
    st := St + 'AFV_DATEMAJ="' + usDateTime(strtodate(fTobIndices.Detail[indice].getvalue('AFV_DATEMAJ_STRING'))) + '",';
    st := St + 'AFV_INDDATEFIN="' + usDateTime(strtodate(fTobIndices.Detail[indice].getvalue('AFV_INDDATEFIN_STRING'))) + '",';
    st := St + 'AFV_INDDATEPUB="' + usDateTime(strtodate(fTobIndices.Detail[indice].getvalue('AFV_INDDATEPUB_STRING'))) + '",';
    st := St + 'AFV_DATEMODIF="' + usDateTime(strtodate(fTobIndices.Detail[indice].getvalue('AFV_DATEMODIF_STRING'))) + '"';
    st := St + 'where AFV_INDCODE="' + lecode + '" and AFV_pubcode="' + pubcode + '" and AFV_INDDATEVAL="' + inddateval + '"';
  end
  else // insert
  begin                                                                 
    st := 'insert into afvalindice (';
    st := st + 'AFV_PUBCODE,AFV_INDCODE,AFV_INDDATEVAL,AFV_DEFINITIF,AFV_CREATEUR,AFV_DATECREATION,AFV_INDCOMMENT,AFV_INDVALEUR,AFV_DATEMODIF,';
    st := St + 'AFV_INDNUMPUB,AFV_INDCODESUIV,AFV_INDDATEPUB,AFV_INDMAJ,AFV_PUBCODESUIV,AFV_INDDATEFIN,AFV_COEFPASSAGE,AFV_COEFRACCORD,AFV_DATEMAJ) values (';
    st := St + '"' + fTobIndices.Detail[indice].getvalue('AFV_PUBCODE') + '","' + fTobIndices.Detail[indice].getvalue('AFV_INDCODE') + '","' +
      usDateTime(strtodate(fTobIndices.Detail[indice].getvalue('AFV_INDDATEVAL_STRING'))) + '",';
    st := St + '"' + fTobIndices.Detail[indice].getvalue('AFV_DEFINITIF') + '","' + fTobIndices.Detail[indice].getvalue('AFV_CREATEUR') + '","' +
      usDateTime(strtodate(fTobIndices.Detail[indice].getvalue('AFV_DATECREATION_STRING'))) + '",';
    st := St + '"' + fTobIndices.Detail[indice].getvalue('AFV_INDCOMMENT') + '",' + variantToSql(fTobIndices.Detail[indice].getvalue('AFV_INDVALEUR')) + ',"' +
      usDateTime(strtodate(fTobIndices.Detail[indice].getvalue('AFV_DATEMODIF_STRING'))) + '",';
    st := St + '"' + fTobIndices.Detail[indice].getvalue('AFV_INDNUMPUB') + '","' + fTobIndices.Detail[indice].getvalue('AFV_INDCODESUIV') + '","' +
      usDateTime(strtodate(fTobIndices.Detail[indice].getvalue('AFV_INDDATEPUB_STRING'))) + '",';
    st := St + '"IMP","' + fTobIndices.Detail[indice].getvalue('AFV_PUBCODESUIV') + '","' +
      usDateTime(strtodate(fTobIndices.Detail[indice].getvalue('AFV_INDDATEFIN_STRING'))) + '",';
    st := St + variantToSql(fTobIndices.Detail[indice].getvalue('AFV_COEFPASSAGE')) + ',' + variantToSql(fTobIndices.Detail[indice].getvalue('AFV_COEFRACCORD')) + ',"' +
      usDateTime(strtodate(fTobIndices.Detail[indice].getvalue('AFV_DATEMAJ_STRING'))) + '")';
  end;
  ExecuteSQL(st);
end;

procedure TOF_AFREVIMPORTINDICE.BTNValiderClick(sender: Tobject);
var j, i: integer;
  old_AIN_INDCODE: string;
  importe: boolean;
  AFV_PUBCODE, AFV_INDDATEVAL, lecode: string;

begin

  importe := false;
  i := 0;
  old_AIN_INDCODE := '';
  while (i < fTobIndices.Detail.count) do // je balaye ma tob
  begin
    lecode := fTobIndices.Detail[i].getvalue('AIN_INDCODE');
    if old_AIN_INDCODE <> lecode then //on change d'indice
    begin
      importe := false;
      // je cherche dans la grille l'indice lu dans la tob et je regarde si importe = oui
      for j := 1 to LaGrille.RowCount - 1 do
        if (lecode = laGrille.cells[0, j]) and Decodeimport(laGrille.cells[1, j]) then
          importe := true;

      if importe then // on change d'indice j'integer donc que la valeur et l'indice !
      begin
        integre_indice(lecode, i);
        AFV_PUBCODE := fTobIndices.Detail[i].getvalue('AFV_PUBCODE');
        AFV_INDDATEVAL := fTobIndices.Detail[i].getvalue('AFV_INDDATEVAL_STRING');
        integre_Valeur_indice(lecode, AFV_PUBCODE, AFV_INDDATEVAL, i);
      end;
    end
    else // on est sur le meme indice
    begin // j'integer donc que la valeur !
      if importe then
      begin
        AFV_PUBCODE := fTobIndices.Detail[i].getvalue('AFV_PUBCODE');
        AFV_INDDATEVAL := fTobIndices.Detail[i].getvalue('AFV_INDDATEVAL_STRING');
        integre_Valeur_indice(lecode, AFV_PUBCODE, AFV_INDDATEVAL, i);
      end;
    end;
    old_AIN_INDCODE := fTobIndices.Detail[i].getvalue('AIN_INDCODE');
    inc(i);
  end;
 
  PGIInfoAF('Valeurs des indices importées.', '');
end;

procedure TOF_AFREVIMPORTINDICE.OnNew;
begin
  inherited;
end;

procedure TOF_AFREVIMPORTINDICE.OnDelete;
begin
  inherited;
end;

procedure TOF_AFREVIMPORTINDICE.OnUpdate;
begin
  inherited;
end;

procedure TOF_AFREVIMPORTINDICE.OnLoad;
begin
  inherited;
end;

function TOF_AFREVIMPORTINDICE.LoadTob(T: TOB): integer;
begin
  fTobIndices := T;                    
  result := 0;
end;
  
{procedure TOF_AFREVIMPORTINDICE.SIMULEIMPORT;
var
  Q: Tquery;
  St: string;
  DateLimite: TdateTime;
begin

  DateLimite := strtodate('01/01/2002');
  try
    st := 'select *  from afvalindice left join afindice on AFV_INDCODE=AIN_INDCODE ';
    st := st + ' where AFV_INDDATEVAL >"' + UsDateTime(DateLimite) + '"';
    st := st + 'order by AIN_INDCODE';
    Q := nil;
    fTobIndices := TOB.Create('LesValeursIndices', nil, -1);
    try
      Q := OpenSQL(st, TRUE);
      fTobIndices.LoadDetailDB('', '', '', Q, false);
    finally
      Ferme(Q);
    end;
  finally
  end;
end;
}

procedure TOF_AFREVIMPORTINDICE.OnArgument(S: string);
var Critere, Champ, valeur: string;
  X: integer;

begin
  inherited;
  Critere := (Trim(ReadTokenSt(S)));
  while (Critere <> '') do
  begin
    if Critere <> '' then
    begin
      X := pos('=', Critere);
      if x <> 0 then
      begin
        Champ := copy(Critere, 1, X - 1);
        Valeur := Copy(Critere, X + 1, length(Critere) - X);
      end;
      if Champ = 'CHEMIN' then CHEMIN := Valeur;
    end;
    Critere := (Trim(ReadTokenSt(S)));
  end;
  LaGrille := THGrid(Getcontrol('LAGRILLE'));
  LaGrille.OnClick := LaGrilleClick;
  BTNValider := TToolbarButton97(Getcontrol('BVALIDER'));
  BTNValider.OnClick := BTNValiderClick;

                                             
  fTobIndices:=TOB.Create('LesValeursIndices',nil,-1);
  TOBLoadFromXMLFile(CHEMIN,LoadTob) ;

  //SIMULEIMPORT;
  rempliGrillepartob;

end;

procedure TOF_AFREVIMPORTINDICE.OnClose;
begin
  inherited;
  fTobIndices.Free;
end;

procedure TOF_AFREVIMPORTINDICE.OnDisplay();
begin
  inherited;
end;

procedure TOF_AFREVIMPORTINDICE.OnCancel();
begin
  inherited;
end;

procedure AFLanceFiche_AFREVIMPORTINDICE(chemin: string);
begin
  AglLanceFiche('AFF', 'AFREVIMPORTINDICE', '', '', 'CHEMIN=' + chemin);
end;

initialization
  registerclasses([TOF_AFREVIMPORTINDICE]);
end.
