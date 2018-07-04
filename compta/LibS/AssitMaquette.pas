unit AssitMaquette;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls,
  dbtables, UTOF, ParamSoc, Grids,FE_Main,PGIEnv, Spin,
  Hqry, HPanel, UiUtil, UtilPGI, galOutil,
  Hcompte, Buttons, Mask, HEnt1, SaisUtil , GuidUtil, Ent1;


type
  TAssitMaq = class(TFAssist)
    TabSheet1: TTabSheet;
    LSTMAQ: TListBox;
    TabSheet2: TTabSheet;
    HLabel4: THLabel;
    GroupBox3: TGroupBox;
    LabelNUMSTD: TLabel;
    NUMSTD: TSpinEdit;
    BTNUMSTD: TToolbarButton97;
    Label1: TLabel;
    LIBSTD: TEdit;
    Bimprime: TToolbarButton97;
    BtSupp: TToolbarButton97;
    procedure bFinClick(Sender: TObject);
    procedure LSTMAQDblClick(Sender: TObject);
    procedure BTNUMSTDClick(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure NUMSTDChange(Sender: TObject);
    procedure BimprimeClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure BtSuppClick(Sender: TObject);
  private
    Typemaquette  : string;
    NumMaquette   : integer;
    { Déclarations privées }
    procedure ImportFromListeFichier (stPath, stType : string);
    function RendNummaq : string;
  public
    { Déclarations publiques }
  end;

var
  AssitMaq: TAssitMaq;
  procedure ListerMaquette (TypeMaq : string);

implementation
uses
Gridsynth,Critedt,ImprimeMaquette;

{$R *.DFM}
procedure ListerMaquette (TypeMaq : string);
begin
     AssitMaq := TAssitMaq.Create(application);
     AssitMaq.Typemaquette := TypeMaq;
     AssitMaq.NumMaquette  := 0;
     AssitMaq.ImportFromListeFichier (V_PGI_Env.PathStd, typeMaq);
     try
        AssitMaq.ShowModal;
        finally
        AssitMaq.Free;
        end;
         Screen.Cursor:=SyncrDefault ;
end;

procedure TAssitMaq.ImportFromListeFichier (stPath, stType : string);
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
       Q1 := OpenSql ('SELECT * from STDMAQ Where STM_NUMPLAN='+St+' AND STM_TYPEMAQ="'+stType+'"',False);
       if not Q1.eof then
           Libelle := Q1.findfield ('STM_LIBELLE').asstring
       else
       begin
       if NumMaq <= 20 then Pred := 'CEG' else Pred := 'STD';
       ExecuteSql('Insert into STDMAQ (STM_NUMPLAN,STM_LIBELLE,STM_PREDEFINI,STM_TYPEMAQ)'+
                     'Values ('+St+',"'+Libelle+'","'+Pred+'","'+stType+'")') ;
       end;
       St := St + ' ' + Libelle;
       LSTMAQ.Items.Add(St);
       Ferme (Q1);
    end;
   r_search := FindNext(SearchRec);
   end;
  FindClose(SearchRec);
end;

function TAssitMaq.RendNummaq : string;
var
Index    : integer;
Text     : string;
begin
     Index := LSTMAQ.ItemIndex;
     if Index >= 0 then
     Text := LSTMAQ.Items[Index];
     Text := copy (Text, 1, 2);
     Result := Text;
end;
procedure TAssitMaq.bFinClick(Sender: TObject);
var
Text,Chemin,Chemin2      : string;
Index                    : integer;
Crit                     : TCritEdtPCL;
Pred                     : string;
Q1                       : TQuery;
begin
  inherited;
     if Typemaquette = 'CR' then Crit.Te := esCR;
     if Typemaquette = 'SIG' then Crit.Te := esSIG;
     if Typemaquette = 'BIL' then Crit.Te := esBIL;
     Index := LSTMAQ.ItemIndex;
     if Index >= 0 then
     Text := LSTMAQ.Items[Index]
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
     Chemin := V_PGI_Env.PathStd + '\'+Typemaquette+Text+'.txt';
     LanceLiasse(Chemin,FALSE,True,['N'],['#,##0.00'],['#,##0.00'],Crit,TRUE) ;
  end
  else
  begin
     if LIBSTD.Text = '' then
     begin
           PGIInfo('Libelle de la maquette est obligatoire','') ;
           LIBSTD.SetFocus;
           exit;
     end;
     Chemin := V_PGI_Env.PathStd + '\'+Typemaquette+Text+'.txt';
     Chemin2 := V_PGI_Env.PathStd + '\'+Typemaquette+Format('%.02d',[NUMSTD.value])+'.txt';

     if not(FileExists(Chemin2)) then
     CopyFile (Pchar(Chemin), Pchar(Chemin2), FALSE)
     else
     begin
     Text := 'La maquette ' + IntToStr(NUMSTD.value) + ' existe déjà'+
     ' Vous pouvez la modifier uniquement';
          PGIBox(Text,'Modification de la maquette');
     end;

     LanceLiasse(Chemin2,FALSE,True,['N'],['#,##0.00'],['#,##0.00'],Crit,TRUE) ;
     if NUMSTD.value <= 20 then Pred := 'CEG' else Pred := 'STD';
     Q1 := OpenSql ('SELECT * from STDMAQ Where STM_NUMPLAN='+IntToStr(NUMSTD.value)+' AND STM_TYPEMAQ="'+Typemaquette+'"',False);
     if Q1.eof then
     ExecuteSql('Insert into STDMAQ (STM_NUMPLAN,STM_LIBELLE,STM_PREDEFINI,STM_TYPEMAQ)'+
                     'Values ('+IntToStr(NUMSTD.value)+',"'+LIBSTD.Text+'","'+Pred+'","'+Typemaquette+'")')
     else
     ExecuteSql('Update STDMAQ set STM_LIBELLE="'+LIBSTD.Text+
     '" Where STM_NUMPLAN='+ IntToStr(NUMSTD.value)+' AND STM_TYPEMAQ="'+Typemaquette+'"') ;
     ferme (Q1);
  end;
     ModalResult := mrOK;
end;

procedure TAssitMaq.LSTMAQDblClick(Sender: TObject);
var
Text, Lib       : string;
Index, Num      : integer;
Lib2            : string;
begin
  inherited;
     Index := LSTMAQ.ItemIndex;
     if Index >= 0 then
     Text := LSTMAQ.Items[Index];
     Text := copy (Text, 1, 2);
     Num := StrToInt (Text);
     Lib := AGLLanceFiche('CP','STDMAQUETTE',Typemaquette+';'+IntToStr(Num), '','');
     Lib2 := copy (LSTMAQ.Items[Index], 4, length(LSTMAQ.Items[Index])-3);
     if (Lib <> Lib2) and (Lib <> '') then
       LSTMAQ.Items[Index] := Text + ' ' + Lib;
end;

procedure TAssitMaq.BTNUMSTDClick(Sender: TObject);
begin
  inherited;
  AGLLanceFiche('CP','STDMAQUETTE','', '','');

end;

procedure TAssitMaq.bSuivantClick(Sender: TObject);
var
Texte, tmp      : string;
Index           : integer;
Q1              : TQuery;
begin
  inherited;
     bFin.Caption := '&Enregistrer';
     Index := LSTMAQ.ItemIndex;
     if Index < 0 then
     begin
           PGIInfo('Selectionnez une maquette','') ;
           exit;
     end;
     tmp := copy (LSTMAQ.Items[Index], 1, 2);

     Texte := 'Voulez-vous créer la maquette à partir de la maquette ' + tmp;
     if HShowMessage('0;Création d'' une maquette;'+Texte+';Q;YN;N;N;','','')<>mrYes then
        bPrecedentClick(Sender)
     else
     begin
         Q1:=OpenSQL('SELECT max(STM_NUMPLAN) as maxno FROM STDMAQ Where STM_TYPEMAQ="'+Typemaquette+'"',True) ;
         if Not Q1.EOF then
            NUMSTD.value := Q1.Fields[0].AsInteger+1;
         Ferme(Q1) ;
         NumMaquette := NUMSTD.value;
     end;
end;

procedure TAssitMaq.NUMSTDChange(Sender: TObject);
var
Q1 : TQuery;
begin
  inherited;
  NumMaquette := NUMSTD.value;
  Q1 := OpenSql ('SELECT STM_LIBELLE FROM STDMAQ WHERE STM_NUMPLAN='+ IntToStr (NumMaquette)+
  ' and STM_TYPEMAQ="'+ TypeMaquette+'"', TRUE);
  if not Q1.EOF then
     LIBSTD.Text := Q1.Fields[0].asstring;
  Ferme(Q1) ;
end;

procedure TAssitMaq.BimprimeClick(Sender: TObject);
var
Chemin,Numm    : string;
begin
  inherited;
     Numm := RendNummaq;
  Chemin := V_PGI_Env.PathStd + '\'+Typemaquette+Numm+'.txt';
  ControlTextToPrinter (Chemin);
end;

procedure TAssitMaq.FormCreate(Sender: TObject);
begin
  inherited;
  if EstSpecif('51502') then NUMSTD.MinValue := 1;
end;

procedure TAssitMaq.bPrecedentClick(Sender: TObject);
begin
  inherited;
     bFin.Caption := '&Modification';
end;

procedure TAssitMaq.BtSuppClick(Sender: TObject);
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

// CA - 27/11/2001 - Maquettes dans Dat
if Strtoint(noStd) <= 20 then
  Fichier := V_PGI_Env.PathStd + '\'+Typemaquette + NoStd + '.txt'
else Fichier := V_PGI_Env.PathDat + '\'+Typemaquette + NoStd + '.txt';
DeleteFile (Pchar(Fichier));
LSTMAQ.Items.Delete(LSTMAQ.ItemIndex);
end;

end.
