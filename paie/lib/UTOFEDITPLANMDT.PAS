{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 28/06/2001
Modifié le ... : 30/08/2001
Description .. : Source TOF de la TABLE : EDITPLANMDT ()
Mots clefs ... : PAIE;MEDECINETRAVAIL
*****************************************************************
PT- 1 : 10/10/2001 : JL 562: Procédures de vérification des dates + calendrier
PT- 2 : 09/01/2002 : JL 571: Exclusion des visites effectuées dans la requête + 585 : Correction PVM_INAPTE au lieu de PVM_APTE ds la requete
PT- 3 : 12/09/2003 : JL V_42: FQ 10710 Tri par heure de RDV
PT- 4 : 21/11/2003 : JL V_50 FQ 10972 ajout date de naissance et numéro SS
PT- 5 : 14/09/2005 : JL V_60 : FQ 12497 Ajout salarié présent au
---- JL 20/03/2006 modification clé annuaire ----
PT6   : 31/01/2007 V_80 FC Mise en place filtrage des habilitations/poupulations
}
Unit UTOFEDITPLANMDT ;

Interface

Uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}DBCtrls,HDB,QRS1,
{$ELSE}
     eQRS1,UTOB,
{$ENDIF}
     HCtrls,HEnt1,HMsgBox,UTOF,ParamDat,
     P5Def //MonHabilitation
      ;

Type
  TOF_EDITPLANMDT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
    procedure VerifDate1(Sender:TObject);
    procedure VerifDate2(Sender:TObject);
    procedure DateElipsisclick(Sender: TObject);
    procedure CkSortieClick (Sender : TObject);
  end ;

Implementation

procedure TOF_EDITPLANMDT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_EDITPLANMDT.OnDelete ;
begin
  Inherited ;
end ;
{
Sur le bouton Valider, on confectionne la requete
}
procedure TOF_EDITPLANMDT.OnUpdate ;
var Etab1,Etab2,requete,orderby,medecine,reqmedecine,reqtype,StWhere:String;
    DATE1,DATE2:TDateTime;
    VTypeVisite:THValCombobox;
    Habilitation:String;//PT6
begin
     Inherited ;
     //DEBUT PT- 5
     If GetCheckBoxState('CKSORTIE') = CbChecked then StWhere := ' AND (PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>"'+UsDateTime(StrToDate(GetControlText('DATESORTIE')))+'")'
     else StWhere := '';
     //FIN PT- 5
     VTypeVisite:=THValCombobox(GetControl('TYPEVISITE'));
     Etab1:=GetControlText('XX_VARETAB1');
     Etab2:=GetControlText('XX_VARETAB2');
     Date1:=StrToDateTime(GetControlText('PVM_DATEVISITE'));
     Date2:=StrToDateTime(GetControlText('PVM_DATEVISITE_'));
     Requete:='SELECT PVM_SALARIE,PVM_DATEVISITE,PVM_MEDTRAVGU,PVM_TYPEVISITMED,PVM_HEUREVISITE'+
              ' ,ANN_APNOM,ANN_APRUE1,ANN_APRUE2,ANN_APRUE3,ANN_APCPVILLE,ANN_EMAIL,ANN_TEL1,ANN_FAX'+
              ' , ANN_NOM1,ANN_NOM2,ANN_NOM3'+
              ' ,PSA_LIBELLE,PSA_PRENOM,PSA_SALARIE,PSA_ETABLISSEMENT,PSA_DATENAISSANCE,PSA_NUMEROSS '+  //PT- 4
              ' FROM VISITEMEDTRAV left join ANNUAIRE on PVM_MEDTRAVGU=ANN_GUIDPER'+
              ' left join SALARIES on PVM_SALARIE=PSA_SALARIE'+
              ' WHERE PSA_ETABLISSEMENT>="'+Etab1+'" AND PSA_ETABLISSEMENT<="'+Etab2+'"'+
              ' AND PVM_DATEVISITE>="'+UsDateTime(Date1)+'" AND PVM_DATEVISITE<="'+UsDateTime(Date2)+'"'+
              ' AND PVM_APTE<>"X" AND PVM_INAPTE<>"X" AND PVM_RECLASSE<>"X"'+StWhere;   // PT- 2
     If GetCheckBoxState('CHEURE') = CbChecked then Orderby:=' order by PSA_ETABLISSEMENT,PVM_MEDTRAVGU,PVM_DATEVISITE,PVM_HEUREVISITE,PVM_SALARIE'  //PT- 3
     else Orderby:=' order by PSA_ETABLISSEMENT,PVM_MEDTRAVGU,PVM_DATEVISITE,PVM_SALARIE' ;
     If (VTypeVisite<>NIL) AND (VTypeVisite.value<>'') Then  reqtype:='AND PVM_TYPEVISITMED="'+VTypeVisite.value+'"'
     Else reqtype:='';

    //DEB PT6
    Habilitation := '';
    if (Assigned(MonHabilitation)) and (MonHabilitation.LeSQL<>'') then
    begin
      Habilitation := ' AND ' + MonHabilitation.LeSQL;
    end;
    //FIN PT6

     If GetControlText ('PVM_MEDTRAVGU')='' then
        TFQRS1(Ecran).WhereSQL:=requete+reqtype+Habilitation+orderby
     else
         begin
         Medecine:=GetControlText('PVM_MEDTRAVGU');
         reqmedecine:= ' AND PVM_MEDTRAVGU="'+Medecine+'"' ;
         TFQRS1(Ecran).WhereSQL:=requete+reqtype+reqmedecine+Habilitation+orderby;
         end;
end ;

procedure TOF_EDITPLANMDT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_EDITPLANMDT.OnArgument (S : String ) ;
var    Q1,Q2:TQuery;
   Etab1,Etab2:String;
   Date1,Date2:THEdit;
   Check : TCheckBox;
begin
  Inherited ;
  Etab1:='';Etab2:='';
  Q1:=OpenSQL('SELECT ET_ETABLISSEMENT FROM ETABLISS ORDER BY ET_ETABLISSEMENT',True);
  If Not Q1.eof Then Etab1:=Q1.FindField('ET_ETABLISSEMENT').AsString;  // // PortageCWAS
  Ferme(Q1);
  SetControlText('XX_VARETAB1',Etab1);
  Q2:=OpenSQL('SELECT ET_ETABLISSEMENT FROM ETABLISS ORDER BY ET_ETABLISSEMENT DESC',True);
  If Not Q2.eof Then Etab2:=Q2.FindField('ET_ETABLISSEMENT').AsString;  // // PortageCWAS
  Ferme(Q2);
  SetControlText('XX_VARETAB2',Etab2);
  Date1:=THEdit(GetControl('PVM_DATEVISITE'));
  If Date1<>NIL Then Date1.OnExit:=VerifDate1;
  If Date1<>NIL Then Date1.OnElipsisClick := DateElipsisclick;
  Date2:=THEdit(GetControl('PVM_DATEVISITE_'));
  If Date2<>NIL Then Date2.OnExit:=VerifDate2;
  If Date2<>NIL Then Date2.OnElipsisClick := DateElipsisclick;
  Check := TCheckBox(GetControl('CKSORTIE'));
  If Check <> Nil Then Check.OnCLick := CkSortieClick;
end ;

//PT- 1 
procedure TOF_EDITPLANMDT.VerifDate1(Sender:TObject);
var THDate1,THDate2:THEdit;
    Date1,Date2:TDatetime;
begin
THDate1:=THEdit(GetControl('PVM_DATEVISITE'));
If Not IsValidDate(THDate1.text) Then
   begin
   PGIBox(''''+THDate1.text+''' n''est pas une date correcte','CEGID PAIE S5');
   THDate1.SetFocus ;
   exit;
   end
Else
    begin
    THDate2:=THEdit(GetControl('PVM_DATEVISITE_'));
    Date1:=StrToDateTime(THDate1.text);
    Date2:=StrToDatetime(THDate2.text);
    If Date1>Date2 Then
       begin
       PGIBox('La date saisie doit être inférieure au '''+THDate2.Text+'''','Edition médecine du travail');
       THDate1.SetFocus ;
       end;
    end;
end;

procedure TOF_EDITPLANMDT.VerifDate2(Sender:TObject);
var THDate1,THDate2:THEdit;
    Date1,Date2:TDatetime;
begin
THDate2:=THEdit(GetControl('PVM_DATEVISITE_'));
If Not IsValidDate(THDate2.text) Then
   begin
   PGIBox(''''+THDate2.text+''' n''est pas une date correcte','CEGID PAIE S5');
   THDate2.SetFocus ;
   exit;
   end
Else
    begin
    THDate1:=THEdit(GetControl('PVM_DATEVISITE'));
    Date1:=StrToDateTime(THDate1.text);
    Date2:=StrToDatetime(THDate2.text);
    If Date1>Date2 Then
       begin
       PGIBox('La date saisie doit être supérieure au '''+THDate1.Text+'''','Edition médecine du travail');
       THDate2.SetFocus ;
       end;
    end;
end;

procedure TOF_EDITPLANMDT.DateElipsisclick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;


procedure TOF_EDITPLANMDT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_EDITPLANMDT.CkSortieClick(Sender : TObject);
begin
        If Sender = Nil then Exit;
        If GetCheckBoxState('CKSORTIE') = CbChecked then
        begin
                SetControlVisible('DATESORTIE',True);
                SetControlText('DATESORTIE',DateToStr(Date));
                SetControlVisible('TDATEARRET',True);
        end
        else
        begin
                SetControlVisible('DATESORTIE',False);
                SetControlVisible('TDATEARRET',False);
        end;
end;

Initialization
  registerclasses ( [ TOF_EDITPLANMDT ] ) ;
end.
