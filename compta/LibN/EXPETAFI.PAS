unit ExpEtafi;

interface
uses  Windows,Controls,StdCtrls,ExtCtrls,Graphics,Grids,
      Classes,ComCtrls,sysutils,dbTables,Ent1,Vierge,
      HCtrls,HEnt1,LettUtil,HMsgBox,LettAuto,UTOF,UTOB;
Type
     TOF_ExpEtafi = class (TOF)
     private
       F6,UnEtab : boolean ;
       DecLocal : integer ;
       DateDebut,DateFin : TDateTime ;
       CFormat,Cexo,CEtab : THValComboBox ;
       DatDeb,DatFin,FileName,TBenef,TPerte,Tresult : THedit ;
       RMonnaie,RSeparateur : TRadioGroup ;
       CHuit : TCheckBox ;
       TheTob : TOB ;
       function CreeLaFille : TOB ;
       function ParamOK : boolean ;
       procedure ExportExo(Exo : string) ;
       procedure RecupSolde(Exo : string;N : integer) ;
       procedure ExportLaTob(Fname : string ) ;
       procedure OnChangeCExo(Sender: TObject) ;
       function CalculExo(Exo : string; N : integer) : string;
       Function FormatSpec(Dbl : double;Typ,Deci : Integer) : string ;
     public
       procedure OnLoad ; override ;
       //procedure OnCancel ; override ;
       procedure OnUpdate ; override ;
       procedure OnClose ; override ;
     end;

implementation
Function TOF_ExpEtafi.FormatSpec(Dbl : double;Typ,Deci : Integer) : string ;
var  OldDecimalSeparator : char ;  i,Expo,Taille : integer ;
begin
Expo:=1 ;for i:=1 to Deci do Expo:=Expo*10;
if F6 then Taille:=17 else Taille:=16 ;
OldDecimalSeparator:=DecimalSeparator ;
case Typ of
  0 : begin DecimalSeparator:=',' ; Result:=format('%*.2f',[Taille,Dbl]) ; end ;
  1 : begin DecimalSeparator:='.' ; Result:=format('%*.2f',[Taille,Dbl]) ; end ;
  2 : Result:=format('%16.0f',[Dbl*Expo]) ;
  else
  Result:='' ;
  end ;
DecimalSeparator:=OldDecimalSeparator ;
End ;

procedure TOF_ExpEtafi.OnLoad ;
var PControl : TPageControl ;
begin
CFormat:=THValComboBox(GetControl('CFORMAT')) ;
CExo   :=THValComboBox(GetControl('CEXO')) ;
CEtab  :=THValComboBox(GetControl('CETAB')) ;
DatFin  :=THEdit(GetControl('DATEFIN')) ;
DatDeb  :=THEdit(GetControl('DATEDEB')) ;
FileName:=THEdit(GetControl('FILENAME')) ;
TBenef  :=THEdit(GetControl('TBENEF')) ;
TPerte  :=THEdit(GetControl('TPERTE')) ;
TResult :=THEdit(GetControl('TRESULT')) ;
RMonnaie   :=TRadioGroup(GetControl('RMONNAIE')) ;
RSeparateur:=TRadioGroup(GetControl('RSEPARATEUR')) ;
CHuit:=TCheckBox(GetControl('CHUIT')) ;
PControl:=TPageControl(GetControl('PCONTROL'));
CFormat.ItemIndex:=0 ; CExo.ItemIndex:=0 ; CEtab.ItemIndex:=0 ;
if (CExo<>nil)  and (not Assigned(CExo.OnChange))  then CExo.OnChange:=OnChangeCexo;
RMonnaie.Items[0]:=VH^.LibDevisePivot ;
if VH^.LibDeviseFongible<>'' then RMonnaie.Items[1]:=VH^.LibDeviseFongible ;
DateDebut:=StrToDate(StDate1900) ;DateFin:=StrToDate(StDate2099) ;
FileName.Text:='C:\ExpEtaf5.txt' ;
if (PControl<>nil) then PControl.ActivePage:=Pcontrol.Pages[0];
end ;

procedure TOF_ExpEtafi.OnUpdate ;
var QQ : TQuery ;
begin
if not ParamOK then Exit ;
TheTob:=TOB.Create('_ETAFI',nil,-1) ;
ExportExo(Cexo.Value) ;
if TheTob<>nil then TheTob.Free ;
end ;

procedure TOF_ExpEtafi.ExportExo(Exo : string) ;
var i : integer ;
begin
for i:=0 to 3 do RecupSolde(Exo,i) ;
ExportLaTob(FileName.Text) ;
end ;

procedure TOF_ExpEtafi.OnClose ;
begin
//if TheTOB<>nil then TheTob.Free ;
end ;

function TOF_ExpEtafi.CreeLaFille : TOB ;
var T1 : TOB ;
begin
T1:=Tob.Create('_ETAFI2',TheTob,-1) ;
T1.AddChampSup('SGENERAL', False) ; T1.PutValue('SGENERAL', '') ;
T1.AddChampSup('GENERAL', False) ; T1.PutValue('GENERAL', '') ;
if F6 then T1.AddChampSup('AUXILIAIRE', False) ;
T1.AddChampSup('LIBELLE', False) ;
T1.AddChampSup('SOLDEAN', False) ; T1.PutValue('SOLDEAN', arrondi(0,DecLocal)) ;
T1.AddChampSup('MVDEBIT', False) ; T1.PutValue('MVDEBIT', arrondi(0,DecLocal));
T1.AddChampSup('MVCREDIT',False) ; T1.PutValue('MVCREDIT',arrondi(0,DecLocal));
T1.AddChampSup('SOLDE',   False) ; T1.PutValue('SOLDE',   arrondi(0,DecLocal));
T1.AddChampSup('SOLDEN1', False) ; T1.PutValue('SOLDEN1', arrondi(0,DecLocal));
T1.AddChampSup('SOLDEN2', False) ; T1.PutValue('SOLDEN2', arrondi(0,DecLocal));
T1.AddChampSup('SOLDEN3', False) ; T1.PutValue('SOLDEN3', arrondi(0,DecLocal));
Result:=T1 ;
end;

function TOF_ExpEtafi.ParamOK : boolean ;
var TMsg : array[1..10] of string ;
begin
TMsg[1]:='1;Export Etafi;La date est en dehors de l''exercice;W;O;O;O' ;
TMsg[2]:='1;Export Etafi;La date est superieur a la date inferieur;W;O;O;O' ;
ParamOK:=True ;
if StrToDate(DatDeb.Text)<DateDebut then begin HShowMessage(TMsg[1],'','') ; ParamOK:=False; DatDeb.SetFocus ; Exit ; end ;
if StrToDate(DatFin.Text)>DateFin   then begin HShowMessage(TMsg[1],'','') ; ParamOK:=False; DatFin.SetFocus ; Exit ; end ;
if StrToDate(DatDeb.Text)>StrToDate(DatFin.Text) then begin HShowMessage(TMsg[2],'','') ; ParamOK:=False; DatFin.SetFocus ; Exit ; end ;
if RMonnaie.ItemIndex=0 then DecLocal:=V_PGI.OkDecV else DecLocal:=V_PGI.OkDecE ;
if CFormat.ItemIndex=1 then F6:=True else F6:=False ;
if CEtab.ItemIndex<>0 then UnEtab:=True else UnEtab:=False ;
if FileExists(FileName.Text) then
  if HshowMessage('1;Export ETAFI;Le fichier '+FileName.Text+' existe déjà, voulez-vous l''ecraser ?;W;YN;Y;Y','','')<>mrYes then
    begin ParamOK:=False; FileName.SetFocus ; Exit ; end ;
end ;

procedure TOF_ExpEtafi.RecupSolde(Exo : string;N : integer) ;
var QQ : TQuery ; T1 : TOB ; Debit,Credit : double ; Sql,ExoLocal,General,GeneTmp : string ;
begin
if N<>0 then ExoLocal:=CalculExo(Exo,N)else ExoLocal:=Exo ; if ExoLocal='' then Exit ;
                              Sql:='SELECT E_GENERAL,G_LIBELLE' ;
if F6 then                Sql:=Sql+',T_LIBELLE,E_AUXILIAIRE' ;
if RMonnaie.ItemIndex=0
then                      Sql:=Sql+',SUM(E_DEBIT) AS SUMDEBIT,SUM(E_CREDIT) AS SUMCREDIT'
else                      Sql:=Sql+',SUM(E_DEBITEURO) AS SUMDEBIT,SUM(E_CREDITEURO) AS SUMCREDIT' ;
if N=0 then               Sql:=Sql+',E_ECRANOUVEAU' ;
                          Sql:=Sql+' FROM ECRITURE ' ;
                          Sql:=Sql+' LEFT JOIN GENERAUX ON G_GENERAL=E_GENERAL' ;
if F6 then                Sql:=Sql+' LEFT JOIN TIERS ON T_AUXILIAIRE=E_AUXILIAIRE' ;
                          Sql:=Sql+' WHERE E_QUALIFPIECE="N" AND E_EXERCICE="'+ExoLocal+'"' ;
if UnEtab then            Sql:=Sql+' AND E_ETABLISSEMENT="'+CEtab.Value+'"';
if N=0 then               Sql:=Sql+' AND E_DATECOMPTABLE>="'+UsDateTime(StrToDate(DatDeb.text))+'" AND E_DATECOMPTABLE<="'+UsDateTime(StrToDate(DatFin.text))+'"' ;
if not CHuit.Checked then Sql:=Sql+' AND E_GENERAL NOT LIKE "8%"' ;
                          Sql:=Sql+' GROUP BY E_GENERAL,G_LIBELLE' ;
if F6 then                Sql:=Sql+',E_AUXILIAIRE,T_LIBELLE' ;
if N=0 then               Sql:=Sql+',E_ECRANOUVEAU ' ;
QQ:=OpenSql(Sql,True) ;
while not QQ.Eof do
  begin
  T1:=nil ;
  General:=QQ.FindField('E_GENERAL').AsString ;
  GeneTmp:=General;
  if (General=TBenef.Text) or (General=TPerte.Text) then General:=TResult.Text ;
  if F6 then
    begin
    T1:=TheTob.FindFirst(['AUXILIAIRE'],[General],True) ;
    if T1=nil then T1:=TheTob.FindFirst(['GENERAL','AUXILIAIRE'],[General,QQ.FindField('E_AUXILIAIRE').AsString],True) ;
    end
    else
      T1:=TheTob.FindFirst(['GENERAL'],[General],True) ;
  if (T1=nil) then  T1:=CreeLaFille ;
  if T1<>nil then
    begin
    if GeneTmp<>General then
      begin
      T1.PutValue('SGENERAL',General) ;
      if not F6 then
        begin
        T1.PutValue('GENERAL',General) ;
        T1.PutValue('LIBELLE',QQ.FindField('G_LIBELLE').AsString) ;
        end
        else begin
        if QQ.FindField('E_AUXILIAIRE').AsString<>'' then
          begin
          T1.PutValue('AUXILIAIRE',QQ.FindField('E_AUXILIAIRE').AsString) ;
          T1.PutValue('GENERAL',General) ;
          T1.PutValue('LIBELLE',QQ.FindField('T_LIBELLE').AsString) ;
          end
          else begin
          T1.PutValue('AUXILIAIRE',General) ;
          T1.PutValue('LIBELLE',QQ.FindField('G_LIBELLE').AsString) ;
          end ;
        end ;
      end ;
    Debit:=arrondi(QQ.FindField('SUMDEBIT').AsFloat,DecLocal) ;
    Credit:=arrondi(QQ.FindField('SUMCREDIT').AsFloat,DecLocal) ;
    case N of
      0 :  begin
           if (QQ.FindField('E_ECRANOUVEAU').AsString='OAN') OR (QQ.FindField('E_ECRANOUVEAU').AsString='H')
           then begin T1.PutValue('SOLDEAN',arrondi(T1.GetValue('SOLDEAN')+Debit-Credit,DecLocal)) ; end
           else begin T1.PutValue('MVDEBIT',arrondi(T1.GetValue('MVDEBIT')+Debit,DecLocal)) ;T1.PutValue('MVCREDIT',arrondi(Credit,DecLocal)) ; end ;
           T1.PutValue('SOLDE',arrondi(T1.GetValue('SOLDEAN')+T1.GetValue('MVDEBIT')-T1.GetValue('MVCREDIT'),DecLocal)) ;
           end ;
      1 :  T1.PutValue('SOLDEN1',arrondi(T1.GetValue('SOLDEN1')+Debit-Credit,DecLocal)) ;
      2 :  T1.PutValue('SOLDEN2',arrondi(T1.GetValue('SOLDEN2')+Debit-Credit,DecLocal)) ;
      3 :  T1.PutValue('SOLDEN3',arrondi(T1.GetValue('SOLDEN3')+Debit-Credit,DecLocal)) ;
      end ;
    QQ.Next ;
    end ;
  end ;
ferme(QQ) ;
end ;

procedure TOF_ExpEtafi.ExportLaTob(FName : string) ;
var FDest : TextFile ;Taille,i : integer ; Ligne : string ; test : double ;
begin
AssignFile(FDest, FName ) ;
ReWrite(FDest) ;
TheTob.Detail.Sort('SGENERAL') ;
if F6 then Taille:=40 else Taille:=35;
for i:=0 to TheTob.Detail.Count-1 do
  begin
  Ligne:='';
  if F6 then Ligne:=Ligne+format('%-16.16s',[TheTob.detail[i].GetValue('AUXILIAIRE')]) ;
  Ligne:=Ligne+format('%-12.12s',[TheTob.detail[i].GetValue('GENERAL')]) ;
  if F6 then Ligne:=Ligne+'    ' ; //racine auxiliaire 4 car.
  Ligne:=Ligne+format('%-*.*s',[Taille,Taille,TheTob.detail[i].GetValue('LIBELLE')]) ;
  Ligne:=Ligne+FormatSpec(double(TheTob.detail[i].GetValue('SOLDEAN') ),RSeparateur.ItemIndex,DecLocal) ;
  Ligne:=Ligne+FormatSpec(double(TheTob.detail[i].GetValue('MVDEBIT') ),RSeparateur.ItemIndex,DecLocal) ;
  Ligne:=Ligne+FormatSpec(double(TheTob.detail[i].GetValue('MVCREDIT')),RSeparateur.ItemIndex,DecLocal) ;
  Ligne:=Ligne+FormatSpec(double(TheTob.detail[i].GetValue('SOLDE')   ),RSeparateur.ItemIndex,DecLocal) ;
  Ligne:=Ligne+FormatSpec(double(TheTob.detail[i].GetValue('SOLDEN1') ),RSeparateur.ItemIndex,DecLocal) ;
  Ligne:=Ligne+FormatSpec(double(TheTob.detail[i].GetValue('SOLDEN2') ),RSeparateur.ItemIndex,DecLocal) ;
  Ligne:=Ligne+FormatSpec(double(TheTob.detail[i].GetValue('SOLDEN3') ),RSeparateur.ItemIndex,DecLocal) ;
  WriteLn(FDest,Ligne) ;
  end ;
CloseFile(FDest) ;
end;

procedure TOF_ExpEtafi.OnChangeCExo(Sender: TObject) ;
var i : integer ;
begin
if (DatDeb=nil) or (DatFin=nil) then Exit ;
ExoToDates(CExo.value,DatDeb,DatFin) ;
end ;

function TOF_ExpEtafi.CalculExo(Exo : string; N : integer) : string;
var i : integer ; TabExo : array[1..8] of string ;
label Fin ;
begin
TabExo[1]:=VH^.Suivant.Code ;
TabExo[2]:=VH^.EnCours.Code ;
TabExo[3]:=VH^.Precedent.Code ;
for i:=1 to 5 do if VH^.Exoclo[i].Code<>VH^.Precedent.Code then TabExo[i+3]:=VH^.Exoclo[i].Code ;
for i:=1 to 8-N do if TabExo[i]=Exo then begin Result:=TabExo[i+N] ; Exit ; end ;
Result:='' ;
end ;

Initialization
registerclasses([TOF_ExpEtafi]);
end.
