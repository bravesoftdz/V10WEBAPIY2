unit UTOFCREATMAQ;

interface
uses

  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, HTB97, Grids, ExtCtrls,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Math,
  HEnt1,HPanel,UiUtil, hmsgbox, UTOB,FE_Main, ComCtrls, Menus, HPop97,
  UTOF,HDB,Mul,
  FichList,
  Spin,
  Ent1;
type
  TOF_CREATMAQ = Class (TOF)
     procedure OnLoad ; override ;
     procedure OnArgument(stArgument: String); override ;
     private
     Typemaquette   : string;
     NumMaquette   : integer;
     function ImportFromListeFichier (stPath, stType : string) : integer;
     procedure bClickValide(Sender: TObject);
     procedure BtSuppClick(Sender: TObject);
     procedure PChange(Sender: TObject);
     procedure PCreat(Sender: TObject);
     function RendNummaq : string;
     procedure LSTMAQDblClick(Sender: TObject);
     procedure BTNUMSTDClick(Sender: TObject);
     procedure NUMSTDChange(Sender: TObject);
     procedure BimprimeClick(Sender: TObject);
  end;

implementation
uses
Gridsynth,Critedt,ImprimeMaquette,galOutil;


procedure TOF_CREATMAQ.OnArgument(stArgument: String);
begin
     Typemaquette := ReadTokenSt (stArgument);
     NumMaquette  := 0;
     SetControlvisible ('FE__TABSHEET2', FALSE);
     TToolbarButton97(GetControl('BValider')).Onclick := bClickValide;
     TToolbarButton97(GetControl('BDelete')).Onclick := BtSuppClick;
     TPageControl(GetControl('P')).OnChange := PChange;
     TListBox(GetControl ('LISTE')).OnDblClick := LSTMAQDblClick;
     TSpinEdit(GetControl('NUMSTD')).OnDblClick := BTNUMSTDClick;
     TSpinEdit(GetControl('NUMSTD')).OnChange := NUMSTDChange;
     TToolbarButton97(GetControl('BInsert')).Onclick := PCreat;
     TToolbarButton97(GetControl('BImprimer')).Onclick := BimprimeClick;
end;

procedure TOF_CREATMAQ.OnLoad ;
begin
  ImportFromListeFichier(ChangeStdDatPath('$STD'), Typemaquette);

  if EstSpecif('51502') then
  TSpinEdit(GetControl('NUMSTD')).MinValue := 1;
end;


function TOF_CREATMAQ.RendNummaq : string;
var
Index    : integer;
Text     : string;
begin
     Index := TListBox(GetControl ('LISTE')).ItemIndex;
     if Index >= 0 then
     Text := TListBox(GetControl ('LISTE')).Items[Index];
     Text := copy (Text, 1, 2);
     Result := Text;
end;


function TOF_CREATMAQ.ImportFromListeFichier (stPath, stType : string) : integer;
var SearchRec   : TSearchRec;
    r_search    : integer;
    NumMaq      : integer;
    St          : string;
    Q1          : TQuery;
    Libelle,Pred: string;
begin
  stPath := UpperCase (stPath);
  r_search := FindFirst(stPath+'\'+stType+'*.txt', faAnyFile, SearchRec);
  while (r_search = 0) do
  begin
    if SearchRec.Attr <> faDirectory then
    begin
       if stType= 'CR' then Libelle := 'COMPTE DE RESULTAT'
       else if stType= 'SIG' then Libelle := 'SOLDE INTERMEDIAIRE DE GESTION'
       else if stType= 'BIL' then Libelle := 'BILAN';

       // SearchRec.Name;
       if stType= 'CR' then
       St := copy (SearchRec.Name, 3, 2)
       else
       St := copy (SearchRec.Name, 4, 2);

       NumMaq := StrToInt (St);
       Q1 := OpenSql ('SELECT * from STD0MAQ Where STM_NUMPLAN='+St+' AND STM_TYPEMAQ="'+stType+'"',False);
       if not Q1.eof then
           Libelle := Q1.findfield ('STM_LIBELLE').asstring
       else
       begin
       if NumMaq <= 20 then Pred := 'CEG' else Pred := 'STD';
       ExecuteSql('Insert into STDMAQ (STM_NUMPLAN,STM_LIBELLE,STM_PREDEFINI,STM_TYPEMAQ)'+
                     'Values ('+St+',"'+Libelle+'","'+Pred+'","'+stType+'")') ;
       end;
       St := St + ' ' + Libelle;
       TListBox(GetControl ('LISTE')).Items.Add(St);
       Ferme (Q1);
    end;
   r_search := FindNext(SearchRec);
   end;
  FindClose(SearchRec);
end;

procedure TOF_CREATMAQ.bClickValide(Sender: TObject);
var
Text,Chemin,Chemin2      : string;
Index                    : integer;
Crit                     : TCritEdtPCL;
Pred                     : string;
Q1                       : TQuery;
NoMaq1, NoMaq2           : integer;
CritEdtChaine : TCritEdtChaine;
begin
  Fillchar(CritEdtChaine,SizeOf(CritEdtChaine),#0) ;
  CritEdtChaine.Utiliser := False;

     if Typemaquette = 'CR' then Crit.Te := esCR;
     if Typemaquette = 'SIG' then Crit.Te := esSIG;
     if Typemaquette = 'BIL' then Crit.Te := esBIL;
     Index := TListBox(GetControl ('LISTE')).ItemIndex;
     if Index >= 0 then
     Text := TListBox(GetControl ('LISTE')).Items[Index]
     else begin
           PGIInfo('Selectionnez une maquette','') ;
           exit;
      end;
     Text := copy (Text, 1, 2);


  // cas modification
  if NumMaquette = 0 then
  begin
    if not EstSpecif('51502') then
     if StrToInt(Text) <= 20 then
           PGIInfo('Maquette CEGID, uniquement en consultation','') ;
     // CA - 28/11/2001
     if StrToInt(Text) <= 20 then
       Chemin := ChangeStdDatPath('$STD') + '\'+Typemaquette+Text+'.txt'
     else
       Chemin := ChangeStdDatPath('$DAT') + '\'+Typemaquette+Text+'.txt';

     LanceLiasse(Chemin,FALSE,True,['N'],['#,##0.00'],['#,##0.00'],Crit,TRUE, CritEdtChaine) ;
  end
  else
  begin
     if THEDIT(GetControl ('LIBSTD')).Text = '' then
     begin
           PGIInfo('Libelle de la maquette est obligatoire','') ;
           SetFocusControl('LIBSTD');
           exit;
     end;
     // CA - 27/11/2001 - Gestion des maquettes dans le répertoire DAT
     NoMaq1 := StrToInt (Text);
     NoMaq2 := TSpinEdit(GetControl ('NUMSTD')).value;
     
     if NoMaq1 <= 20 then
       Chemin := ChangeStdDatPath('$STD') + '\' + Typemaquette+Text+'.txt'
     else
       Chemin := ChangeStdDatPath('$DAT') + '\' + Typemaquette+Text+'.txt';

     if NoMaq2 <= 20 then
       Chemin2 := ChangeStdDatPath('$STD') + '\'+Typemaquette+Format('%.02d',[TSpinEdit(GetControl ('NUMSTD')).value])+'.txt'
     else
       Chemin2 := ChangeStdDatPath('$DAT') + '\'+Typemaquette+Format('%.02d',[TSpinEdit(GetControl ('NUMSTD')).value])+'.txt';

     if not(FileExists(Chemin2)) then
     CopyFile (Pchar(Chemin), Pchar(Chemin2), FALSE)
     else
     begin
     Text := 'La maquette ' + IntToStr(TSpinEdit(GetControl('NUMSTD')).value) + ' existe déjà'+
     ' Vous pouvez la modifier uniquement';
          PGIBox(Text,'Modification de la maquette');
     end;

     LanceLiasse(Chemin2,FALSE,True,['N'],['#,##0.00'],['#,##0.00'],Crit,TRUE, CritEdtChaine) ;
     if TSpinEdit(GetControl('NUMSTD')).value <= 20 then Pred := 'CEG' else Pred := 'STD';
     Q1 := OpenSql ('SELECT * from STDMAQ Where STM_NUMPLAN='+IntToStr(TSpinEdit(GetControl('NUMSTD')).value)+' AND STM_TYPEMAQ="'+Typemaquette+'"',False);
     if Q1.eof then
     ExecuteSql('Insert into STDMAQ (STM_NUMPLAN,STM_LIBELLE,STM_PREDEFINI,STM_TYPEMAQ)'+
                     'Values ('+IntToStr(TSpinEdit(GetControl('NUMSTD')).value)+',"'+THEdit(GetControl('LIBSTD')).Text+'","'+Pred+'","'+Typemaquette+'")')
     else
     ExecuteSql('Update STDMAQ set STM_LIBELLE="'+THEdit(GetControl('LIBSTD')).Text+
     '" Where STM_NUMPLAN='+ IntToStr(TSpinEdit(GetControl('NUMSTD')).value)+' AND STM_TYPEMAQ="'+Typemaquette+'"') ;
     ferme (Q1);
     Text := Format('%.02d',[TSpinEdit(GetControl ('NUMSTD')).value])
     + ' ' + THEdit(GetControl('LIBSTD')).Text;
     TListBox(GetControl ('LISTE')).Items.Add(Text);
     NumMaquette := 0;
     TPageControl(GetControl('P')).Activepage := TTabSheet(GetControl('FE__TABSHEET1'));
     SetControlvisible ('FE__TABSHEET2', FALSE);
  end;
// TFMul(Ecran).ModalResult := 1;
end;

procedure TOF_CREATMAQ.BtSuppClick(Sender: TObject);
var
Predef            : string;
Where,Fichier     : string;
NoStd             : string;
begin
  inherited;
if not IsSuperviseur(TRUE) then exit;
NoStd := RendNummaq;
if StrToint (NoStd) < 21 then Predef := 'CEG'
else Predef := 'STD';
if not EstSpecif('51502') then
if Strtoint(noStd) <= 20 then
begin PGIInfo('Maquette CEGID,impossible de supprimer','') ;  exit;end;

Where := 'Where STM_TYPEMAQ="'+ Typemaquette+'" and STM_NUMPLAN='+ NoStd+
' and STM_PREDEFINI="'+Predef+'"';
ExecuteSQL('DELETE FROM STDMAQ '+Where);
// CA - 27/11/2001 - Gestion maquettes cabinet dans répertoire DAT
if Strtoint(noStd) <= 20 then
  Fichier := ChangeStdDatPath('$STD') + '\' + Typemaquette + NoStd + '.txt'
else
  Fichier := ChangeStdDatPath('$DAT') + '\'+Typemaquette + NoStd + '.txt';

DeleteFile (Pchar(Fichier));
TListBox(GetControl ('LISTE')).Items.Delete(TListBox(GetControl ('LISTE')).ItemIndex);
end;

procedure TOF_CREATMAQ.PChange(Sender: TObject);
begin
if TPageControl(GetControl('P')).Activepage = TTabSheet(GetControl('FE__TABSHEET1'))
then SetControlvisible ('FE__TABSHEET2', FALSE);
end;

procedure TOF_CREATMAQ.PCreat(Sender: TObject);
var
Texte, tmp      : string;
Index           : integer;
Q1              : TQuery;
begin
     TPageControl(GetControl('P')).Activepage := TTabSheet(GetControl('FE__TABSHEET2'));
     SetControlvisible ('FE__TABSHEET2', TRUE);
     Index := TListBox(GetControl ('LISTE')).ItemIndex;
     if Index < 0 then
     begin
           PGIInfo('Selectionnez une maquette','') ;
           exit;
     end;
     tmp := copy (TListBox(GetControl ('LISTE')).Items[Index], 1, 2);
     Texte := 'Duplication à partir de la maquette n° :' + tmp;
     THEdit(GetControl ('LIBMAQUETTE')).text := Texte;

     Q1:=OpenSQL('SELECT max(STM_NUMPLAN) as maxno FROM STDMAQ Where STM_TYPEMAQ="'+Typemaquette+'"',True) ;
     if Not Q1.EOF then
        TSpinEdit(GetControl ('NUMSTD')).value := Q1.Fields[0].AsInteger+1;
     Ferme(Q1) ;
     NumMaquette := TSpinEdit(GetControl ('NUMSTD')).value;
end;

procedure TOF_CREATMAQ.LSTMAQDblClick(Sender: TObject);
var
Text, Lib       : string;
Index, Num      : integer;
Lib2            : string;
begin
  inherited;
     Index := TListBox(GetControl ('LISTE')).ItemIndex;
     if Index >= 0 then
     Text := TListBox(GetControl ('LISTE')).Items[Index];
     Text := copy (Text, 1, 2);
     Num := StrToInt (Text);
     Lib := AGLLanceFiche('CP','STDMAQUETTE',Typemaquette+';'+IntToStr(Num), '','');
     Lib2 := copy (TListBox(GetControl ('LISTE')).Items[Index], 4, length(TListBox(GetControl ('LISTE')).Items[Index])-3);
     if (Lib <> Lib2) and (Lib <> '') then
       TListBox(GetControl ('LISTE')).Items[Index] := Text + ' ' + Lib;
end;

procedure TOF_CREATMAQ.BTNUMSTDClick(Sender: TObject);
begin
  inherited;
  AGLLanceFiche('CP','STDMAQUETTE','', '','');

end;
procedure TOF_CREATMAQ.NUMSTDChange(Sender: TObject);
var
Q1 : TQuery;
begin
  inherited;
  NumMaquette := TSpinEdit (GetControl ('NUMSTD')).value;
  Q1 := OpenSql ('SELECT STM_LIBELLE FROM STDMAQ WHERE STM_NUMPLAN='+ IntToStr (NumMaquette)+
  ' and STM_TYPEMAQ="'+ TypeMaquette+'"', TRUE);
  if not Q1.EOF then
     THEdit(GetControl('LIBSTD')).Text := Q1.Fields[0].asstring;
  Ferme(Q1) ;
end;

procedure TOF_CREATMAQ.BimprimeClick(Sender: TObject);
var
Chemin,Numm    : string;
begin
  inherited;
  Numm := RendNummaq;
  // CA - 27/11/2001
  if StrToInt(Numm)<=20 then
    Chemin := ChangeStdDatPath('$STD') + '\' + Typemaquette+Numm+'.txt'
  else
    Chemin := ChangeStdDatPath('$DAT') + '\'+Typemaquette+Numm+'.txt';

  ControlTextToPrinter (Chemin);
end;


Initialization
RegisterClasses([TOF_CREATMAQ]);

end.
