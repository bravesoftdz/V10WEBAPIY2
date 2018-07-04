unit MajTPayeur ;

interface
uses Windows,Controls,StdCtrls,ExtCtrls,Graphics,Grids,
      Classes,ComCtrls,sysutils,dbTables,
      HCtrls,HEnt1,LettUtil,HMsgBox,LettAuto,UTOF,UTOB,LookUp,GetSema,Ent1,test;
type
     TOF_MajTPayeur = class (TOF)
     private
       procedure OnFacturClick(Sender : TObject) ;
       procedure OnPayeurClick(Sender : TObject) ;
       procedure OnNatureClick(Sender : TObject) ;
     public
       procedure OnLoad ; override ;
       procedure OnUpdate ; override ;
//       procedure OnCancel ; override ;
//       procedure OnClose ; override ;
     end;

implementation
procedure TOF_MajTPayeur.OnLoad;
begin
THEdit(GetControl('FACTURDEB')).OnElipsisClick:=OnFacturClick ;
THEdit(GetControl('FACTURFIN')).OnElipsisClick:=OnFacturClick ;
THEdit(GetControl('PAYEURDEB')).OnElipsisClick:=OnPayeurClick ;
THEdit(GetControl('PAYEURFIN')).OnElipsisClick:=OnPayeurClick ;
TRadioGroup(GetControl('NATURE')).OnClick:=OnNatureClick ;
end;

procedure TOF_MajTPayeur.OnFacturClick(Sender : TObject) ;
var stWhere : string ;
begin
StWhere:= 'T_NATUREAUXI="'+THRadioGroup(GetControl('NATURE')).Value+'" AND T_ISPAYEUR="-"' ;
LookupList(THEdit(Sender),TraduireMemoire('Comptes Auxiliaires'),'TIERS','T_AUXILIAIRE','T_LIBELLE',stWhere,'T_AUXILIAIRE',FALSE,1) ;
end;


procedure TOF_MajTPayeur.OnPayeurClick(Sender : TObject) ;
var stWhere : string ;
begin
StWhere:= 'T_NATUREAUXI="'+THRadioGroup(GetControl('NATURE')).Value+'" AND T_ISPAYEUR="X"' ;
LookupList(THEdit(Sender),TraduireMemoire('Comptes Auxiliaires'),'TIERS','T_AUXILIAIRE','T_LIBELLE',stWhere,'T_AUXILIAIRE',FALSE,1) ;
end;

procedure TOF_MajTPayeur.OnNatureClick(Sender : TObject) ;
var stWhere : string ;
begin
THEdit(GetControl('FACTURDEB')).Text:='' ;
THEdit(GetControl('FACTURFIN')).Text:='' ;
THEdit(GetControl('PAYEURDEB')).Text:='' ;
THEdit(GetControl('PAYEURFIN')).Text:='' ;
end;

procedure TOF_MajTPayeur.OnUpdate;
var QQ,Q1 : TQuery ; TheTob,TheTob2 : TOB ; Sql,FacturDeb,FacturFin,Payeurdeb,PayeurFin : string ; Adr,Rib : boolean ; i,j,Retour : integer ;
begin
Adr:=TCheckBox(GetControl('ADRESSE')).Checked ;
Rib:=TCheckBox(GetControl('RIB')).Checked ;
FacturDeb:=THEdit(GetControl('FACTURDEB')).Text ;
FacturFin:=THEdit(GetControl('FACTURFIN')).Text ;
PayeurDeb:=THEdit(GetControl('PAYEURDEB')).Text ;
PayeurFin:=THEdit(GetControl('PAYEURFIN')).Text ;
if (FacturDeb>FacturFin) then begin HShowMessage('1;Mise à jour des tiers Payeurs;Le client facturé de debut est superieur à celui de fin;W;O;O;O','','') ; Exit ; end ;
if (PayeurDeb>PayeurFin) then begin HShowMessage('1;Mise à jour des tiers Payeurs;Le client payeur de debut est superieur à celui de fin;W;O;O;O','','') ; Exit ; end ;
if Not Adr and Not Rib then Exit ;
Sql:='SELECT T1.T_AUXILIAIRE,T1.T_PAYEUR,T2.T_ADRESSE1,T2.T_ADRESSE2,T2.T_ADRESSE3,T2.T_TELEPHONE,'
      +'T2.T_TELEX,T2.T_FAX,T2.T_CODEPOSTAL,T2.T_VILLE,T2.T_PAYS,T2.T_DIVTERRIT,T2.T_RVA '
      +'FROM TIERS T1 LEFT JOIN TIERS T2 ON T2.T_AUXILIAIRE=T1.T_PAYEUR '
      +'WHERE T1.T_PAYEUR<>"" ' ;
if FacturDeb<>'' then  Sql:=Sql+'AND T1.T_AUXILIAIRE>="'+FacturDeb+'" ';
if FacturFin<>'' then  Sql:=Sql+'AND T1.T_AUXILIAIRE<="'+FacturFin+'" ';
if PayeurDeb<>'' then  Sql:=Sql+'AND T1.T_PAYEUR>="'+PayeurDeb+'" ';
if PayeurFin<>'' then  Sql:=Sql+'AND T1.T_PAYEUR<="'+PayeurFin+'" ';
Sql:=Sql+'AND T1.T_NATUREAUXI="'+THRadioGroup(GetControl('NATURE')).Value+'" ' ;
QQ:=OpenSql(Sql,True) ;
While not QQ.eof do
  begin
  if Adr then
    begin
    Retour:=ExecuteSql('UPDATE TIERS SET '
        +'T_ADRESSE1="'  +QQ.FindField('T_ADRESSE1').AsString  +'",'
        +'T_ADRESSE2="'  +QQ.FindField('T_ADRESSE2').AsString  +'",'
        +'T_ADRESSE3="'  +QQ.FindField('T_ADRESSE3').AsString  +'",'
        +'T_TELEPHONE="' +QQ.FindField('T_TELEPHONE').AsString +'",'
        +'T_TELEX="'     +QQ.FindField('T_TELEX').AsString     +'",'
        +'T_FAX="'       +QQ.FindField('T_FAX').AsString       +'",'
        +'T_CODEPOSTAL="'+QQ.FindField('T_CODEPOSTAL').AsString+'",'
        +'T_VILLE="'     +QQ.FindField('T_VILLE').AsString     +'",'
        +'T_PAYS="'      +QQ.FindField('T_PAYS').AsString      +'",'
        +'T_DIVTERRIT="' +QQ.FindField('T_DIVTERRIT').AsString +'",'
        +'T_RVA="'       +QQ.FindField('T_RVA').AsString       +'" '
        +'WHERE T_AUXILIAIRE="'+QQ.FindField('T_AUXILIAIRE').AsString+'" ') ;
    end ;
  if Rib then
    begin
    ExecuteSql('DELETE FROM RIB WHERE R_AUXILIAIRE="'+QQ.FindField('T_AUXILIAIRE').AsString+'" ') ;
    TheTob:=Tob.Create('§RIB',nil,-1) ;
    TheTob.LoadDetailDB('RIB','"'+QQ.FindField('T_PAYEUR').AsString+'"','',nil,True,True) ;
    for i:=0 to TheTob.Detail.Count-1 do
      begin
      TheTob.Detail[i].PutValue('R_AUXILIAIRE',QQ.FindField('T_AUXILIAIRE').AsString) ;
      TheTob.Detail[i].SetAllModifie(True) ;
      TheTob.Detail[i].InsertDB(nil,True) ;
      end ;
    TheTob.Free ;
    end;
  QQ.next ;
  end ;
ferme(QQ) ;
end ;

Initialization
registerclasses([TOF_MajTPayeur]);
end.
