{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 28/05/2001
Modifié le ... : 30/08/2001
Description .. : Source TOF Edition des dernières visites médicales du
Suite ........ : salarie
Mots clefs ... : PAIE;MEDECINETRAVAIL
*****************************************************************
PT- 1 : 10/10/2001 : JL 562: Procédures de vérification des dates + calendrier
PT- 2 : 26/03/2002 : JL 571: Ajout tri alphabétique et modification requête (champs de la table salariés en supp)
PT- 3 : 14/09/2005 : JL V_60 : FQ 12497 Ajout salarié présent au
PT- 4 : 10/01/2007 : FCL V_80 : FQ 13326 Ajout du champ PVM_DATEVISITMED à la requête
---- JL 20/03/2006 modification clé annuaire ----
PT-5  : 31/01/2007 V_80 FC Mise en place filtrage des habilitations/poupulations
PT6   : 05/10/2007 V_80 FC FQ 14820 Apte sous réserve non disponible
}
Unit UTOFEDITMEDTRAV ;

Interface

Uses StdCtrls, Controls,Classes,forms,sysutils,ComCtrls,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}QRS1,EdtEtat,
{$ELSE}
     eQrs1,UtileAGL,UTOB,
{$ENDIF}
     HCtrls,HEnt1,HMsgBox,UTOF,ParamDat,HTB97,
     P5Def //MonHabilitation
     ;

Type
  TOF_EDITMEDTRAV = Class (TOF)

    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    procedure DateElipsisclick(Sender: TObject);
    procedure VerifDate1(Sender:TObject);
    procedure VerifDate2(Sender:TObject);
    procedure ChangeTri(Sender:TObject);
    procedure CkSortieClick (Sender : TObject);
   end ;

Implementation

// Confection de la requete en fonction des critères de la fiche QRS1
procedure TOF_EDITMEDTRAV.OnUpdate ;
var A1,A2 : THEdit;
    T1,T2:TDateTime;
    Etab1,Etab2,Org,StWhere: String;
    SQL,orderby,result,SQLETAB,SQLOrg:String;
    CEtab:TCheckBox;
    Habilitation:String;//PT5
begin
Inherited ;
//DEBUT PT- 3
If GetCheckBoxState('CKSORTIE') = CbChecked then StWhere := 'AND (PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>"'+UsDateTime(StrToDate(GetControlText('DATESORTIE')))+'") '
else StWhere := '';
//FIN PT- 3
CEtab:=TCheckBox(GetControl('CETAB'));
Etab1:=GetControlText('XX_VARETAB1');
Etab2:=GetControlText('XX_VARETAB2');
If CEtab.Checked=True Then
   begin
   SetControltext('XX_RUPTURE1','PSA_ETABLISSEMENT');
   orderby:=' ORDER BY PSA_ETABLISSEMENT,PVM_MEDTRAVGU';
   end
Else
   begin
   SetControltext('XX_RUPTURE1','');
   orderby:=' ORDER BY PVM_MEDTRAVGU';
   end;
If GetControlText('ALPHA')='X' Then OrderBy:=OrderBy+',PSA_LIBELLE'   //PT- 2
Else
    begin
    If GetControlText('TRIDATE')='X' Then OrderBy:=OrderBy+',PVM_DATEVISITE'
    Else OrderBy:=OrderBy+',PSA_SALARIE';
    end;
// Dans cette procedure on selectionne toutes les dernières visites par salarié dans l'intervalle
// de dates choisi par l'utilisateur.
A1:=THEdit (GetControl ('PVM_DATEVISITE'));
A2:=THEdit (GetControl ('PVM_DATEVISITE_'));
T1:=StrToDatetime(A1.text);
T2:=strToDatetime(A2.text);
SQL:='SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_DATENAISSANCE,PSA_LIBELLEEMPLOI,PSA_DATEENTREE,PVM_SALARIE,PVM_DATEVISITE,PVM_APTE,PVM_INAPTE,PVM_APTESR,PVM_MEDTRAVGU,PVM_DATEINA,PVM_RECLASSE,PVM_DATEREC,PVM_TYPEVISITMED,PVM_DATEVISITMED'+ //PT- 2   //PT4 //PT6
     ' ,PSA_ETABLISSEMENT, ANN_NOM1,ANN_NOM2,ANN_NOM3 FROM VISITEMEDTRAV V1 LEFT JOIN SALARIES on PVM_SALARIE=PSA_SALARIE'+
     ' left join ANNUAIRE on PVM_MEDTRAVGU=ANN_GUIDPER WHERE '+
     ' (V1.PVM_DATEVISITE IN (SELECT MAX(V2.PVM_DATEVISITE) FROM VISITEMEDTRAV V2 WHERE V1.PVM_SALARIE=V2.PVM_SALARIE GROUP BY V2.PVM_SALARIE) AND (V1.PVM_DATEVISITE>"'+UsDateTime(T1)+'"'+
     ' OR V1.PVM_DATEVISITE="'+UsDateTime(T1)+'") AND (V1.PVM_DATEVISITE<"'+UsDateTime(T2)+'" OR V1.PVM_DATEVISITE="'+UsDateTime(T2)+'")) '+StWhere;

if GetControlText('RESULTAT')='TOUS' then result:=' ';
if GetControlText('RESULTAT')='A' then result:=' AND PVM_APTE="X" ';
if GetControlText('RESULTAT')='S' then result:=' AND PVM_APTESR="X" ';  //PT6
if GetControlText('RESULTAT')='I' then result:=' AND PVM_INAPTE="X" ';
if GetControlText('RESULTAT')='R' then result:=' AND PVM_RECLASSE="X" ';
Org:=GetControlText ('PVM_MEDTRAVGU');
SQLEtab:=' AND PSA_ETABLISSEMENT>="'+Etab1+'" AND  PSA_ETABLISSEMENT<="'+Etab2+'" ';

//DEB PT5
Habilitation := '';
if (Assigned(MonHabilitation)) and (MonHabilitation.LeSQL<>'') then
begin
  Habilitation := ' AND ' + MonHabilitation.LeSQL;
end;
//FIN PT5

If Org='' then
    TFQRS1(Ecran).WhereSQL:=SQL+result+SQLEtab+Habilitation+orderby  //PT5
Else
    Begin
    SQLOrg:=' AND PVM_MEDTRAVGU="'+Org+'" ';        // DB2
    TFQRS1(Ecran).WhereSQL:=SQL+result+SQLETAB+SQLOrg+Habilitation+orderby; //PT5
    end;
end;

procedure TOF_EDITMEDTRAV.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_EDITMEDTRAV.OnArgument (S : String ) ;
var    Q1,Q2:TQuery;
       Etab1,Etab2:String;
       Date1,Date2:THEdit;
       TCDefaut,Check:TCheckBox;
begin
Inherited ;
Etab1:='';Etab2:='';
Q1:=OpenSQL('SELECT ET_ETABLISSEMENT FROM ETABLISS ORDER BY ET_ETABLISSEMENT',True);
If Not Q1.Eof Then Etab1:=Q1.FindField('ET_ETABLISSEMENT').AsString; // // PortageCWAS
Ferme(Q1);
SetControlText('XX_VARETAB1',Etab1);
Q2:=OpenSQL('SELECT ET_ETABLISSEMENT FROM ETABLISS ORDER BY ET_ETABLISSEMENT DESC',True);
If Not Q2.Eof Then Etab2:=Q2.FindField('ET_ETABLISSEMENT').AsString;   // // PortageCWAS
Ferme(Q2);
SetControlText('XX_VARETAB2',Etab2);
SetControlText('RESULTAT','TOUS');
Date1:=THEdit(GetControl('PVM_DATEVISITE'));
If Date1<>NIL Then Date1.OnElipsisClick := DateElipsisclick;
If Date1<>NIL Then Date1.OnExit := VerifDate1;
Date2:=THEdit(GetControl('PVM_DATEVISITE_'));
If Date2<>NIL Then Date2.OnElipsisClick := DateElipsisclick;
If Date2<>NIL Then Date2.OnExit := VerifDate2;
TCDefaut:=TCheckBox(GetControl('ALPHA'));
If TCDefaut<>Nil then TCDefaut.OnClick:=ChangeTri;
TCDefaut:=TCheckBox(GetControl('TRIDATE'));
If TCDefaut<>Nil then TCDefaut.OnClick:=ChangeTri;
Check := TCheckBox(GetControl('CKSORTIE'));
If Check <> Nil Then Check.OnCLick := CkSortieClick;
end ;

// PT- 1
procedure TOF_EDITMEDTRAV.DateElipsisclick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_EDITMEDTRAV.VerifDate1(Sender:TObject);
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

procedure TOF_EDITMEDTRAV.VerifDate2(Sender:TObject);
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

procedure TOF_EDITMEDTRAV.ChangeTri(Sender:TObject);
var Date,Alpha:TCheckBox;
begin
If Sender=Nil Then Exit;
Date:=TCheckBox(GetControl('TRIDATE'));
Alpha:=TCheckBox(GetControl('ALPHA'));
If Date.Checked=True Then
   begin
   Alpha.Enabled:=False;
   Alpha.Checked:=False;
   end
Else Alpha.Enabled:=True;
If Alpha.Checked=True Then
   begin
   Date.Enabled:=False;
   Date.Checked:=False;
   end
Else Date.Enabled:=True;
end;

procedure TOF_EDITMEDTRAV.CkSortieClick(Sender : TObject);
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
  registerclasses ( [ TOF_EDITMEDTRAV ] ) ;
end.
