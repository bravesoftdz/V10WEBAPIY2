unit UTOFReprisestand;

interface

uses
  Windows, Messages, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Mask, ExtCtrls, ComCtrls,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  classes, sysutils,UTOF, HMsgBox, HTB97, ParamSoc, Grids, HCtrls,
  UTob, AGLInitPlus,AssistPL, HEnt1,
  LicUtil,MajTable, PGIExec, fe_main, HStatus,UlibStdCpta;


type
  TOF_REPRISESTAND = Class (TOF)
     procedure OnArgument(stArgument: String); override ;
     private
     NumStd,NumRup,NumCorre,NumLibre    : string;
     Tbg, ATob                          : TOB;
     procedure Clickcpta(Sender: TObject);
     procedure MiseAjourgrid;
     procedure ClickRup(Sender: TObject);
     procedure ClickCorre(Sender: TObject);
     procedure ClickLibre(Sender: TObject);
     procedure SupprimeListeEnreg(L : THGrid);
     procedure DeleteListe(Sender: TObject);
     procedure ValideClick(Sender: TObject);
     procedure SaveInfoListe (NumStd : integer; LIST : string);
end;

implementation

procedure TOF_REPRISESTAND.OnArgument(stArgument: String);
begin
  TToolbarButton97(GetControl('BTSTD')).Onclick := Clickcpta;
  TToolbarButton97(GetControl('RUPEDIT')).Onclick := ClickRup;
  TToolbarButton97(GetControl('CPTECORRE')).Onclick := ClickCorre;
  TToolbarButton97(GetControl('BTLIBRE')).Onclick := ClickLibre;
  TToolbarButton97(GetControl('BDELETE')).Onclick := DeleteListe;
  TToolbarButton97(GetControl('BVALIDER')).Onclick := ValideClick;
end;

procedure TOF_REPRISESTAND.Clickcpta(Sender: TObject);
begin
     NumStd := AglLanceFiche('CP','STDPLCPTE','','','');
     MiseAjourgrid;
end;
procedure TOF_REPRISESTAND.ClickRup(Sender: TObject);
begin
     NumRup := AglLanceFiche('CP','STDPLCPTE','','','');
     MiseAjourgrid;
end;
procedure TOF_REPRISESTAND.ClickCorre(Sender: TObject);
begin
     NumCorre := AglLanceFiche('CP','STDPLCPTE','','','');
     MiseAjourgrid;
end;
procedure TOF_REPRISESTAND.ClickLibre(Sender: TObject);
begin
     NumLibre := AglLanceFiche('CP','STDPLCPTE','','','');
     MiseAjourgrid;
end;

procedure TOF_REPRISESTAND.DeleteListe(Sender: TObject);
begin
SupprimeListeEnreg(THGrid(GetControl('LISTESTD')));
end;

procedure TOF_REPRISESTAND.SupprimeListeEnreg(L : THGrid);
var i : integer;
Texte : string;
begin
  if (L.NbSelected=0) and (not L.AllSelected) then
  begin
    MessageAlerte('Aucun élément sélectionné');
    exit;
  end;
  Texte:='Vous allez supprimer définitivement les informations.#13#10Confirmez vous l''opération ?';
  if HShowMessage('0;Suppression historique;'+Texte+';Q;YN;N;N;','','')<>mrYes then exit ;
  if L.AllSelected then
  BEGIN
      MoveCur(False);
      L.AllSelected:=False;
  END
  ELSE
  BEGIN
    InitMove(L.NbSelected,'');
    for i:=0 to L.NbSelected-1 do
    BEGIN
      MoveCur(False);
      if L.Cells[1,L.Row] ='Standard d''édition' then NumStd := '';
      if L.Cells[1,L.Row] ='Plan de Rupture' then NumRup := '';
      if L.Cells[1,L.Row] ='Comptes de correspondance' then NumCorre := '';
      if L.Cells[1,L.Row] ='Tables libres' then NumLibre := '';
      L.DeleteRow(L.Row);
      L.GotoLeBookmark(i);
      L.ClearSelected;
    END;
  END;
  FiniMove;
end;


procedure TOF_REPRISESTAND.MiseAjourgrid;
begin
  Tbg := TOB.Create('', nil, -1);
  if (NumStd <> '') then
  begin
     ATob := TOB.Create('STANDARD',Tbg,-1);
     ATob.AddChampSup('No',True);
     ATob.PutValue('No', NumStd);
     ATob.AddChampSup('Libelle',True);
     ATob.PutValue('Libelle','Standard d''édition');
  end;

  if (NumRup <> '') then
  begin
     ATob := TOB.Create('RUPTURE',Tbg,-1);
     ATob.AddChampSup('No',True);
     ATob.PutValue('No', NUMRUP);
     ATob.AddChampSup('Libelle',True);
     ATob.PutValue('Libelle','Plan de Rupture');
  end;


  if (NumCorre <> '') then
  begin
     ATob := TOB.Create('CORRES',Tbg,-1);
     ATob.AddChampSup('No',True);
     ATob.PutValue('No', NumCorre);
     ATob.AddChampSup('Libelle',True);
     ATob.PutValue('Libelle','Comptes de correspondance');
  end;


  if (NumLibre <> '') then
  begin
     ATob := TOB.Create('LIBRE',Tbg,-1);
     ATob.AddChampSup('No',True);
     ATob.PutValue('No', NumLibre);
     ATob.AddChampSup('Libelle',True);
     ATob.PutValue('Libelle','Tables libres');
  end;
  Tbg.PutGridDetail(THGrid(GetControl('LISTESTD')),False,False,'No;Libelle');
  Tbg.Free;
end;

procedure TOF_REPRISESTAND.ValideClick(Sender: TObject);
begin
  if NumStd <> '' then
  begin
     LoadStandardMaj(StrToInt(NumStd),'FILTRES','FILTRESREF','TABLE;LIBELLE',' ');
     SaveInfoListe(StrToInt(NumStd), 'MULMMVTS');
     SaveInfoListe(StrToInt(NumStd), 'MULMANAL');
     SaveInfoListe(StrToInt(NumStd), 'MULVMVTS');

  end;
  if (NumRup <> '') then
  begin
     LoadStandardMaj(StrToInt(NumRup),'CHOIXCOD','CHOIXCODREF','TYPE;CODE', ' ');
     LoadStandardMaj(StrToInt(NumRup),'RUPTURE','RUPTUREREF','NATURERUPT;PLANRUPT;CLASSE', ' ');
  end;
  if (NumCorre <> '') then
     LoadStandardMaj(StrToInt(NumCorre),'CORRESP','CORRESPREF','TYPE;CORRESP', ' ');
  if (NumLibre <> '') then
     LoadStandardMaj(StrToInt(NumLibre),'NATCPTE','NATCPTEREF','TYPECPTE;NATURE', ' ');

end;

procedure TOF_REPRISESTAND.SaveInfoListe (NumStd : integer; LIST : string);
var  Q1, Q2 : TQuery;
begin
       Q1 := OpenSQL('SELECT * FROM LISTEREF WHERE LIR_NUMPLAN='+IntToStr (NumStd)+ ' AND LIR_LISTE="'+LIST+'" '
       + 'AND LIR_UTILISATEUR="---"', TRUE);
  if not Q1.Eof then
  begin
       Q2 := OpenSQL('SELECT * FROM LISTE WHERE LI_LISTE="'+LIST+'" '
       + 'AND LI_UTILISATEUR="---"', FALSE);
       if Q2.Eof then
       begin
            Q2.Insert;
            InitNew(Q2);
            Q2.FindField('LI_NUMPLAN').AsInteger := NumStd;
            Q2.FindField('LI_LISTE').AsString := LIST;
            Q2.FindField('LI_UTILISATEUR').AsString := '---';
            Q2.FindField('LI_LIBELLE').AsString := Q1.findField ('LIR_LIBELLE').asstring;
            Q2.FindField('LI_SOCIETE').AsString := Q1.findField ('LIR_SOCIETE').asstring;
            Q2.FindField('LI_NUMOK').AsVariant := Q1.findField ('LIR_NUMOK').AsVariant;
            Q2.FindField('LI_TRIOK').AsVariant := Q1.findField ('LIR_TRIOK').AsVariant;
            Q2.FindField('LI_LANGUE').AsString := Q1.findField ('LIR_LANGUE').asstring;
            Q2.FindField('LI_DATA').AsVariant := Q1.findField ('LIR_DATA').AsVariant;
            Q2.Post;
       end
       else
           Q2.Edit;
           Q2.FindField('LI_DATA').Asstring := Q1.findField ('LIR_DATA').Asstring;
           Q2.Post;
       begin
       end;
       Ferme (Q2);
  end;
  Ferme (Q1);
end;


Initialization
RegisterClasses([TOF_REPRISESTAND]);
end.
