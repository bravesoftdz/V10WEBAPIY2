unit UtofMailingFic;

interface


uses  Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HMsgBox,UTOF,HQry,HEnt1,StdCtrls,
{$IFDEF EAGLCLIENT}
      eMul,MaineAGL,
{$ELSE}
      mul,Fe_Main,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}DBgrids,
{$ENDIF}
      Hstatus,HDB,M3FP,UTOB, UtilSelection,Paramsoc,UtilPGI,UtilRT,UtilWord;

Type
    TOF_MailingFic = Class (TOF)
    Public
        procedure OnArgument (Arguments : String ); override ;
        procedure OnLoad ; override ;
     END;

Function GetTitreGrid (DBListe : TDbGrid ) : string ;

implementation


procedure TOF_MailingFic.OnArgument (Arguments : String );
var    F : TForm;
begin
inherited ;
F := TForm (Ecran);
MulCreerPagesCL(F);
end;

procedure TOF_MailingFic.OnLoad;
var xx_where,StWhere,StListeActions,StJoint,ListeCombos: string;
    DateDeb,DateFin : TDateTime;
    i: integer;
begin
inherited ;
StWhere := '';
if (GetControlText('SANSSELECT') <> 'X') then
   begin
   DateDeb := StrToDate(GetControlText('DATEACTION'));
   DateFin := StrToDate(GetControlText('DATEACTION_'));
   if GetControlText('SANS') = 'X' then StWhere := 'NOT ';
   StWhere:=StWhere + 'EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE (RAC_AUXILIAIRE = RTTIERS.T_AUXILIAIRE';
   if GetControlText('TYPEACTION') <> '' then
      begin
      StListeActions:=FindEtReplace(GetControlText('TYPEACTION'),';','","',True);
      StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
      StWhere:=StWhere + ' AND RAC_TYPEACTION in '+StListeActions;
      end;
   if GetControlText('RESPONSABLE') <> '' then
      begin
      ListeCombos := GetControlText('RESPONSABLE');
      if copy(ListeCombos,length(ListeCombos),1) <> ';' then ListeCombos := ListeCombos + ';';
      ListeCombos:=FindEtReplace(ListeCombos,';','","',True);
      ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
      StWhere:=StWhere + ' AND RAC_INTERVENANT in '+ListeCombos;
      end;
   StWhere := StWhere + ' AND RAC_DATEACTION >= "'+UsDateTime(DateDeb) +'" AND RAC_DATEACTION <= "'+UsDateTime(DateFin)+'"))';
   if (GetControlText('XX_WHERE') = '') then
      SetControlText('XX_WHERE',StWhere)
   else
       begin
       xx_where := GetControlText('XX_WHERE');
       xx_where := xx_where + ' and (' + StWhere + ')';
       SetControlText('XX_WHERE',xx_where) ;
       end;
   end;  

if (GetControlText('XX_WHERE') = '') then
    SetControlText('XX_WHERE',RTXXWhereConfident('CON'))
else
    begin
    xx_where := GetControlText('XX_WHERE');
    xx_where := xx_where + RTXXWhereConfident('CON');
    SetControlText('XX_WHERE',xx_where) ;
    end;
StJoint := '';
if (GetCheckBoxState ('PRINCIPAL') <> cbGrayed) then
   StJoint := StJoint + 'C_PRINCIPAL = "' + GetControlText ('PRINCIPAL') + '" and ';
if (GetCheckBoxState ('PUBLIPOSTAGE') <> cbGrayed) then
   StJoint := StJoint + 'C_PUBLIPOSTAGE = "' + GetControlText ('PUBLIPOSTAGE') + '" and ';
if (GetControlText('FONCTIONCODEE') <> '') and (GetControlText('FONCTIONCODEE') <> '<<Tous>>') then
   begin
//   ListeCombos:=FindEtReplace(GetControlText('FONCTIONCODEE'),';','","',True);
//   ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
//   StJoint := StJoint + 'C_FONCTIONCODEE in ' + ListeCombos + ' and ';
     StJoint := StJoint + 'C_FONCTIONCODEE = "' + GetControlText ('FONCTIONCODEE') + '" and ';
   end;
For i:=1 to 3 do
    begin
    if (GetControlText('LIBRECONTACT'+intToStr(i)) <> '') and (GetControlText('LIBRECONTACT'+intToStr(i)) <> '<<Tous>>') then
        begin
//        ListeCombos:=FindEtReplace(GetControlText('LIBRECONTACT'+intToStr(i)),';','","',True);
//        ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
//        StJoint := StJoint + 'C_LIBRECONTACT' +intToStr(i) +' in ' + ListeCombos + ' and ';
          StJoint := StJoint + 'C_LIBRECONTACT' +intToStr(i) +' = "' + GetControlText('LIBRECONTACT'+intToStr(i)) + '" and ';
        end;
    end;
SetControlText ('XX_JOIN',StJoint);
end;


Function GetTitreGrid (DBListe : TDbGrid ) : string ;
var C : integer;
    st : string;
begin
Result:='' ;
if DBListe=Nil then exit;
{$IFDEF EAGLCLIENT}
For C:=0 to DBListe.ColCount-1 do Result:=Result+Trim(DBListe.Cells[C,0])+';' ;
{$ELSE}
For C:=0 to DBListe.FieldCount-1 do Result:=result+DBListe.Fields[C].DisplayLabel+';' ;
{$ENDIF}
end;

Initialization
registerclasses([TOF_MailingFic]);
end.
