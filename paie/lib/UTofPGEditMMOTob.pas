{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 20/11/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGEDITMMOTOB ()
Mots clefs ... : TOF;PGEDITMMOTOB
*****************************************************************
PT1 : 12/09/2003 : JL V_42 FQ 10809 Ajout contrôle pour DATE et code PCS
                                    + gestion nouveau code
PT2 : 13/10/2003 : JL V_42 FQ 10905 Rajout StPages pour édition crières en CWAS
PT3 : 17/10/2003 : JL V_42 FQ 10912 Format date entrée et date sortie modifié car pas de format dans état
PT4 : 27/09/2006 : JL V_75 FQ 13532 Gestion motif entrée pour apprentissage 
PT14 : 23/02/2007 JL V_70 FQ 13943 Corrections requete effectif et entrée
PT5 : 11/05/2007 JL V_720 FQ 14222 Corrections date effectif

}
Unit UTofPGEditMMOTob;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}EdtREtat,
{$ELSE}
       UtilEAgl,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,UTOB ,HTB97,HSysMenu,LookUp,ParamDat,PGEdtEtat,EntPaie;

Type
  TOF_PGEDITMMOTOB = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnUpdate;override ;
    private
    Etab,TypeElipsis,SexeAvModif:String;
    GMMO:THGrid;
    DateDeb,DateDeb2,DateFin,EntreeAvModif,SortieAvModif:TDateTime;
    TobEtat : Tob;
    procedure AfficheGrille;
    procedure LancerEtatAvecTob(Sender:TObject);
    procedure GMMOColEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GMMOElipsisClick(Sender: TObject);
    procedure AjoutLigne(Sender:TObject);
    procedure SupprimerLigne(Sender:TObject);
    procedure Effectif(Sender : Tobject);
    procedure DateElipsisclick(Sender: TObject);
    procedure ExitDateDeb(Sender:TObject);
    procedure OnCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure OnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
  end ;

Implementation

Procedure TOF_PGEDITMMOTOB.OnUpdate;
begin
  Inherited ;
end;

procedure TOF_PGEDITMMOTOB.OnArgument (S : String ) ;
var BEtat,BPlus,Bmoins:TToolBarButton97;
    Combo:THValComboBox;
    MultiCombo : THMultiValComboBox;
    THDate:THEdit;
    DateFinPrec : TDateTime;
begin
  Inherited ;
Etab:=ReadTokenPipe(S,';');
DateDeb:=StrToDate(ReadTokenPipe(S,';'));
DateDeb2:=StrToDate(ReadTokenPipe(S,';'));
DateFin:=StrToDate(ReadTokenPipe(S,';'));
SetControlProperty('THVETABLISS','Value',Etab);
SetControlText('ETAB',Etab);
PGGlbEtabMMO:=Etab;
SetControltext('XX_VARIABLEDEB',DateToStr(DateDeb));
SetControlText('XX_VARIABLEFIN',DateToStr(DateFin));
DateFinPrec := FinDeMois(PlusMois(DateDeb,-1)); //PT5
SetControlcaption('LIBEFF','Effectif au ' +DateToStr(DateFinPrec));
SetControlCaption('LIBFIN','Effectif au ' +DatEtoStr(DateFin) );
SetControlText('XX_EFFTOTAL',ReadTokenPipe(S,';'));
SetControlText('XX_NBENTREE',ReadTokenPipe(S,';'));
SetControlText('XX_NBSORTIES',ReadTokenPipe(S,';'));
SetControlText('XX_EFFFINAL',ReadTokenPipe(S,';'));
SetControlText('XX_H',ReadTokenPipe(S,';'));
SetControlText('XX_F',ReadTokenPipe(S,';'));
SetControlText('XX_NBINTERIM',ReadTokenPipe(S,';'));
SetControlText('XX_VARIABLEEFF',ReadTokenPipe(S,';'));
TypeElipsis:='';
GMMO:=THGrid(GetControl('GMMO'));
if GMMO<>NIL then
   begin
   GMMO.colformats[2] := 'CB=PGMMONATIONALITE';
   GMMO.colformats[3] := '0000';
   GMMO.ColLengths[3]:=4;
   GMMO.ColTypes[3]   := 'I';
   GMMO.colformats[4] := 'CB=PGSEXE';
   GMMO.ColEditables[5] := False;
   GMMO.ColTypes[7] := 'D';
   GMMO.colformats[7] :=  ShortDateFormat;
   GMMO.colformats[8] := 'CB=PGMMONATURECONTRAT';
   GMMO.ColTypes[9] := 'D';
   GMMO.colformats[9] :=  ShortDateFormat;
   GMMO.colformats[10] := 'CB=PGMMOMOTIFSORTIE';
  // GMMO.ColEditables[8]:=False;
   GMMO.ColAligns[3]:=taRightJustify ;
   GMMO.OnColEnter := GMMOColEnter;
   GMMO.OnElipsisClick := GMMOElipsisClick;
   GMMO.OnCellEnter := OnCellEnter;
   GMMO.OnCellExit := OnCellExit;
   end;
AfficheGrille;
Combo:=THValComboBox(getControl('THVETABLISS'));
If combo<>Nil then Combo.OnChange:=Effectif;
MultiCombo:=THMultiValComboBox(getControl('ETABCOMPL'));
If MultiCombo<>Nil then MultiCombo.OnChange:=Effectif;
BEtat:=TToolBarButton97(GetControl('BETAT'));
If BEtat<>Nil then BEtat.OnClick:=LancerEtatAvecTob;
BPlus:=TToolBarButton97(GetControl('BPLUS'));
If BPlus<>Nil then BPlus.OnClick:=AjoutLigne;
BMoins:=TToolBarButton97(GetControl('BMOINS'));
If BMoins<>Nil then BMoins.OnClick:=SupprimerLigne;
THDate:=THEdit(GetControl('XX_VARIABLEDEB'));
If THDate<>Nil Then
   begin
   THDate.OnElipsisClick:=DateElipsisClick;
   THDate.OnExit:=ExitDateDeb;
   end;
THDate:=THEdit(GetControl('XX_VARIABLEFIN'));
If THDate<>Nil Then THDate.OnElipsisClick:=DateElipsisClick;
THDate:=THEdit(GetControl('XX_VAREDITION'));
If THDate<>Nil Then THDate.OnElipsisClick:=DateElipsisClick;
end ;

procedure TOF_PGEDITMMOTOB.AfficheGrille;
var Q:Tquery;
    TobMMO,TobPaieMMO1,TobPaieMMO2,TPM1,TPM2:Tob;
    i:Integer;
    Salarie,Nationalite,MotiFentree,MotifSortie,St : String;
    dateNaissance : TDateTime;
    aa,mm,jj:Word;
    HMTrad    : THSystemMenu;
    E,EtabCompl,Etabliss,WhereEtab : String;
   WhereEtabPaie2,WhereEtabPaie1,WhereEtabSal : String;
   DateEff,MDateEff : TDateTime;
   NbContrat : Integer;
begin
     DateDeb := StrToDate(GetControlText('XX_VARIABLEDEB'));
     DateFin :=StrToDate(GetControlText('XX_VARIABLEFIN'));
     MDateEff := PlusMois(DateDeb,-1);
     DateEff := FinDeMois(MDateEff);
     DateDeb2 := DebutDeMois(DateEff);

E := GetControlText('THVETABLISS');
For i:=1 to GMMO.RowCount-1 do
    begin
    GMMO.Rows[i].Clear;
    end;
GMMO.RowCount:=2;
GMMO.FixedRows:=1;
NbContrat := 0;
//Salariés avec contrat
EtabCompl := GetControlText('ETABCOMPL');
Etabliss := ReadTokenPipe(EtabCompl,';');
While Etabliss <> '' do
begin
     WhereEtab := WhereEtab + ' OR PCI_ETABLISSEMENT="'+Etabliss+'"';
     WhereEtabPaie1 := WhereEtabPaie1 + ' OR P1.PPU_ETABLISSEMENT="'+Etabliss+'"';
     WhereEtabPaie2 := WhereEtabPaie2 + ' OR P2.PPU_ETABLISSEMENT="'+Etabliss+'"';
     WhereEtabSal := WhereEtabSal + ' OR PSA_ETABLISSEMENT="'+Etabliss+'"';
     Etab := ReadTokenPipe(EtabCompl,';');
     Etabliss := ReadTokenPipe(EtabCompl,';');
end;
WhereEtab := '(PCI_ETABLISSEMENT="'+E+'"' + WhereEtab+')';
WhereEtabSal := '(PSA_ETABLISSEMENT="'+E+'"' + WhereEtabSal+')';
WhereEtabPaie1 := '(P1.PPU_ETABLISSEMENT="'+E+'"' + WhereEtabPaie1+')';
WhereEtabPaie2 := '(P2.PPU_ETABLISSEMENT="'+E+'"' + WhereEtabPaie2+')';
Q := OpenSQL('SELECT PCI_MOTIFSORTIE,PSA_LIBELLE,PSA_PRENOM,PSA_NATIONALITE,PSA_SEXE,PSA_DATENAISSANCE,PSA_CODEEMPLOI,'+
'PCI_SALARIE,PCI_DEBUTCONTRAT,PCI_FINCONTRAT,PCI_TYPECONTRAT,PCI_MOTIFSORTIE FROM CONTRATTRAVAIL LEFT JOIN SALARIES ON PCI_SALARIE=PSA_SALARIE WHERE '+
WhereEtab+' AND ((PCI_DEBUTCONTRAT>="'+UsDateTime(DateDeb)+'" AND PCI_DEBUTCONTRAT<="'+UsDateTime(DateFin)+'") OR '+
'(PCI_FINCONTRAT>="'+UsDateTime(DateDeb)+'" AND PCI_FINCONTRAT<="'+UsDateTime(DateFin)+'"))',True);
TobMMO:=Tob.Create('Les salariés',Nil,-1);
TobMMO.LoadDetailDB('les salariés','','',Q,False);
Ferme(Q);
For i:=0 to tobMMO.Detail.Count-1 do
begin
     NbContrat := NbContrat+1;
    TobMMo.Detail[i].AddChampSup('ANNEE',False);
    DateNaissance:=TobMMO.detail[i].GetValue('PSA_DATENAISSANCE');
    decodedate(DateNaissance, aa,mm,jj);
    TobMMO.Detail[i].putValue('ANNEE',aa);
    If i>0 then GMMO.RowCount:=GMMO.RowCount+1;
    MotifEntree:='';
    MotifSortie:='';
    Salarie:=TobMMO.Detail[i].GetValue('PCI_SALARIE');
    GMMO.CellValues[0,i+1]:=TobMMO.Detail[i].GetValue('PSA_LIBELLE');
    GMMO.CellValues[1,i+1]:=TobMMO.Detail[i].GetValue('PSA_PRENOM');
    Nationalite := TobMMO.Detail[i].GetValue('PSA_NATIONALITE');
    If Nationalite = 'FRA' then GMMO.CellValues[2,i+1]:= 'F'
    Else
    begin
        Q := OpenSQL('SELECT PY_MEMBRECEE FROM PAYS WHERE PY_PAYS="'+Nationalite+'"',True);
        If Not Q.Eof then
        begin
                If Q.FindField('PY_MEMBRECEE').AsString = 'X' then GMMO.CellValues[2,i+1]:= 'C'
                Else GMMO.CellValues[2,i+1]:= 'A';
        end
        Else GMMO.CellValues[2,i+1]:= 'A';
        Ferme(Q);
    end;
    GMMO.CellValues[3,i+1]:=TobMMO.Detail[i].GetValue('ANNEE');
    GMMO.CellValues[4,i+1]:=TobMMO.Detail[i].GetValue('PSA_SEXE');
    If TobMMO.Detail[i].GetValue('PSA_CODEEMPLOI')<>'' Then
      begin
      if VH_Paie.PGPCS2003 then St := 'CO_TYPE="PZE"'  //PT1
     else St:='CO_TYPE = "PCS"';
      Q:=OpenSQL('SELECT CO_LIBRE FROM COMMUN WHERE '+St+' AND CO_ABREGE="'+TobMMO.Detail[i].GetValue('PSA_CODEEMPLOI')+'"',True);
      If not Q.eof then GMMO.CellValues[5,i+1]:=Q.FindField('CO_LIBRE').AsString;
      Ferme(Q);
      end;
    GMMO.CellValues[6,i+1]:=TobMMO.Detail[i].GetValue('PSA_CODEEMPLOI');
    GMMO.CellValues[7,i+1]:=DateToStr(TobMMO.Detail[i].GetValue('PCI_DEBUTCONTRAT'));
    If TobMMO.Detail[i].GetValue('PCI_TYPECONTRAT') = 'CCD' then MotifEntree := 'RD'
    else If TobMMO.Detail[i].GetValue('PCI_TYPECONTRAT') = 'CAA' then MotifEntree := 'RD'
    else If TobMMO.Detail[i].GetValue('PCI_TYPECONTRAT') = 'CAC' then MotifEntree := 'RD'
    else If TobMMO.Detail[i].GetValue('PCI_TYPECONTRAT') = 'CAP' then MotifEntree := 'RD'
    else MotifEntree := 'RI';
    GMMO.CellValues[8,i+1]:=MotifEntree;
    If TobMMO.Detail[i].GetValue('PCI_FINCONTRAT')<=DateFin then
    GMMO.CellValues[9,i+1]:=DateToStr(TobMMO.Detail[i].GetValue('PCI_FINCONTRAT'))
    Else GMMO.CellValues[9,i+1]:='';
    MotifSortie := TobMMO.Detail[i].GetValue('PCI_MOTIFSORTIE');
    Q:=OpenSQL('SELECT PMS_MAINOEUVRE FROM MOTIFSORTIEPAY WHERE PMS_CODE="'+MotifSortie+'"',True);
    If Not Q.Eof then MotifSortie := Q.FindField('PMS_MAINOEUVRE').AsString
    Else MotifSortie := '';
    Ferme(Q);
    end;
//    GMMO.CellValues[10,i+1]:=TobMMO.Detail[i].GetValue('PSA_MOTIFSORTIE');
    GMMO.CellValues[10,i+1]:=MotifSortie;
TobMMO.Free;


//Salariés sans contrat
Q:=OpenSQL('SELECT PPU_SALARIE,PPU_ETABLISSEMENT FROM PAIEENCOURS WHERE PPU_DATEDEBUT="'+UsDateTime(DateDeb)+'"',True);
TobPaieMMO1:=TOB.Create('les paies 1',Nil,-1);
TobPaieMMO1.LoadDetailDB('PAIEENCOURS','','',Q,False);
Ferme(Q);
Q:=OpenSQL('SELECT PPU_SALARIE,PPU_ETABLISSEMENT FROM PAIEENCOURS WHERE PPU_DATEDEBUT="'+UsDateTime(DateDeb2)+'"',True);
TobPaieMMO2:=TOB.Create('les paies 2',Nil,-1);
TobPaieMMO2.LoadDetailDB('PAIEENCOURS','','',Q,False);
Ferme(Q);
Q:=OpenSQL('SELECT PSA_SEXE,PSA_CODEEMPLOI,PSA_SALARIE,PSA_DATESORTIE,PSA_DATEENTREE, PSA_LIBELLE,PSA_PRENOM,PSA_DATENAISSANCE'+  //PT6
        ',PSA_NATIONALITE,PSA_MOTIFSORTIE,PSA_MOTIFENTREE FROM SALARIES left join PAIEENCOURS P1 on psa_SALARIE=P1.PPU_SALARIE'+
        ' AND P1.PPU_DATEDEBUT="'+UsDateTime(DateDeb2)+'"'+
        ' WHERE PSA_PRISEFFECTIF="X" AND ((PSA_DADSPROF<>"09" AND PSA_DADSPROF<>"10" AND PSA_DADSPROF<>"11") OR PSA_DADSPROF="") AND'+    //PT- 10 PT12
        ' (('+WhereEtabSal+' AND ((PSA_DATESORTIE >="'+UsDateTime(Datedeb)+'" AND PSA_DATESORTIE <="'+UsDateTime(DateFin)+'")'+
        ' OR (PSA_DATEENTREE >="'+UsDateTime(Datedeb)+'" AND  PSA_DATEENTREE <="'+UsDateTime(DateFin)+'")))'+
        ' OR (Exists (SELECT PPU_SALARIE FROM PAIEENCOURS P2 WHERE P2.PPU_SALARIE=PSA_SALARIE AND'+
        ' P2.PPU_DATEDEBUT="'+UsDateTime(Datedeb)+'" AND P1.PPU_ETABLISSEMENT<>P2.PPU_ETABLISSEMENT'+
        ' AND ('+WhereEtabPaie1+' OR '+WhereEtabPaie2+'))))'+
        ' AND PSA_SALARIE NOT IN (SELECT PCI_SALARIE FROM CONTRATTRAVAIL WHERE '+
        WhereEtab+' AND ((PCI_DEBUTCONTRAT>="'+UsDateTime(DateDeb)+'" AND PCI_DEBUTCONTRAT>="'+UsDateTime(DateFin)+'") OR '+
        '(PCI_FINCONTRAT<="'+UsDateTime(DateDeb)+'" AND PCI_FINCONTRAT>="'+UsDateTime(DateFin)+'")))'+
        ' GROUP BY PSA_SEXE,PSA_CODEEMPLOI,PSA_SALARIE,PSA_DATESORTIE,PSA_DATEENTREE, PSA_LIBELLE,PSA_PRENOM,PSA_DATENAISSANCE,PSA_NATIONALITE,PSA_MOTIFSORTIE,PSA_MOTIFENTREE'+
       ' ORDER BY PSA_SALARIE',True);
TobMMO:=Tob.Create('Les salariés',Nil,-1);
TobMMO.LoadDetailDB('SALARIES','','',Q,False);
Ferme(Q);
For i:=0 to tobMMO.Detail.Count-1 do
    begin
    TobMMo.Detail[i].AddChampSup('ANNEE',False);
    DateNaissance:=TobMMO.detail[i].GetValue('PSA_DATENAISSANCE');
    decodedate(DateNaissance, aa,mm,jj);
    TobMMO.Detail[i].putValue('ANNEE',aa);
    If (NbContrat>0) or (i>0) then GMMO.RowCount:=GMMO.RowCount+1;
    MotifEntree:='';
    MotifSortie:='';
    Salarie:=TobMMO.Detail[i].GetValue('PSA_SALARIE');
    GMMO.CellValues[0,i+NbContrat+1]:=TobMMO.Detail[i].GetValue('PSA_LIBELLE');
    GMMO.CellValues[1,i+NbContrat+1]:=TobMMO.Detail[i].GetValue('PSA_PRENOM');
    Nationalite := TobMMO.Detail[i].GetValue('PSA_NATIONALITE');
    If Nationalite = 'FRA' then GMMO.CellValues[2,i+NbContrat+1]:= 'F'
    Else
    begin
        Q := OpenSQL('SELECT PY_MEMBRECEE FROM PAYS WHERE PY_PAYS="'+Nationalite+'"',True);
        If Not Q.Eof then
        begin
                If Q.FindField('PY_MEMBRECEE').AsString = 'X' then GMMO.CellValues[2,i+NbContrat+1]:= 'C'
                Else GMMO.CellValues[2,i+NbContrat+1]:= 'A';
        end
        Else GMMO.CellValues[2,i+NbContrat+1]:= 'A';
        Ferme(Q);
    end;
    GMMO.CellValues[3,i+NbContrat+1]:=TobMMO.Detail[i].GetValue('ANNEE');
    GMMO.CellValues[4,i+NbContrat+1]:=TobMMO.Detail[i].GetValue('PSA_SEXE');
    If TobMMO.Detail[i].GetValue('PSA_CODEEMPLOI')<>'' Then
      begin
      if VH_Paie.PGPCS2003 then St := 'CO_TYPE="PZE"'  //PT1
     else St:='CO_TYPE = "PCS"';
      Q:=OpenSQL('SELECT CO_LIBRE FROM COMMUN WHERE '+St+' AND CO_ABREGE="'+TobMMO.Detail[i].GetValue('PSA_CODEEMPLOI')+'"',True);
      If not Q.eof then GMMO.CellValues[5,i+NbContrat+1]:=Q.FindField('CO_LIBRE').AsString;
      Ferme(Q);
      end;
    GMMO.CellValues[6,i+NbContrat+1]:=TobMMO.Detail[i].GetValue('PSA_CODEEMPLOI');
    TPM1:=TobPaieMMO1.FindFirst(['PPU_SALARIE'],[Salarie],False);
    TPM2:=TobPaieMMO2.FindFirst(['PPU_SALARIE'],[Salarie],False);
    If (TobMMO.Detail[i].GetValue('PSA_DATEENTREE')>=Datedeb) and (TobMMO.Detail[i].GetValue('PSA_DATEENTREE')<=DateFin) then
    GMMO.CellValues[7,i+NbContrat+1]:=DateToStr(TobMMO.Detail[i].GetValue('PSA_DATEENTREE'))
    Else
        begin
        If (TPM1<>Nil) and (TPM2<>Nil) then
           begin
           If TPM1.GetValue('PPU_ETABLISSEMENT')<>TPM2.GetValue('PPU_ETABLISSEMENT') then
              begin
              If TPM1.GetValue('PPU_ETABLISSEMENT')=Etab Then begin GMMO.CellValues[7,i+NbContrat+1]:=DateToStr(dateDeb);MotifEntree:='TE';end
              Else GMMO.CellValues[7,i+NbContrat+1]:=DateToStr(TobMMO.Detail[i].GetValue('PSA_DATEENTREE'));
              end
           Else GMMO.CellValues[7,i+NbContrat+1]:=DateToStr(TobMMO.Detail[i].GetValue('PSA_DATEENTREE'));
           end
        Else GMMO.CellValues[7,i+NbContrat+1]:=DateToStr(TobMMO.Detail[i].GetValue('PSA_DATEENTREE'));
        end;
     If MotifEntree = '' then
     begin
        MotifEntree := 'RI';
        Q := OpenSQL('SELECT PCI_SALARIE,PCI_TYPECONTRAT FROM CONTRATTRAVAIL WHERE PCI_DEBUTCONTRAT<="'+UsDateTime(DateFin)+'" '+
        'AND (PCI_FINCONTRAT="'+UsDateTime(IDate1900)+'" OR PCI_FINCONTRAT>="'+UsDateTime(DateDeb)+'") '+
        'ORDER BY PCI_DEBUTCONTRAT DESC',True);
        If Not Q.Eof then
        begin
                If Q.FindField('PCI_TYPECONTRAT').AsString = 'CCD' then MotifEntree := 'RD'
                //DEBUT PT4
                else If Q.FindField('PCI_TYPECONTRAT').AsString = 'CAA' then MotifEntree := 'RD'
                else If Q.FindField('PCI_TYPECONTRAT').AsString = 'CAC' then MotifEntree := 'RD'
                else If Q.FindField('PCI_TYPECONTRAT').AsString = 'CAP' then MotifEntree := 'RD';
                //FIN PT4
        end;
        Ferme(Q);
     end;
//    GMMO.CellValues[8,i+1]:=TobMMO.Detail[i].GetValue('PSA_MOTIFENTREE');
    GMMO.CellValues[8,i+NbContrat+1]:=MotifEntree;
    If (TobMMO.Detail[i].GetValue('PSA_DATESORTIE')>=Datedeb) and (TobMMO.Detail[i].GetValue('PSA_DATESORTIE')<=DateFin) then
    GMMO.CellValues[9,i+NbContrat+1]:=DateToStr(TobMMO.Detail[i].GetValue('PSA_DATESORTIE'))
    Else
        begin
        If (TPM1<>Nil) and (TPM2<>Nil) then
           begin
           If TPM1.GetValue('PPU_ETABLISSEMENT')<>TPM2.GetValue('PPU_ETABLISSEMENT') then
              begin
              If TPM2.GetValue('PPU_ETABLISSEMENT')=Etab Then begin GMMO.CellValues[9,i+NbContrat+1]:=DateToStr(DateDeb);MotifSortie:='TS';end
              Else GMMO.CellValues[9,i+NbContrat+1]:=('');
              end
           Else GMMO.CellValues[9,i+NbContrat+1]:='';
           end
        Else GMMO.CellValues[9,i+NbContrat+1]:='';
        end;
    If MotifSortie ='' then
    begin
           MotifSortie := TobMMO.Detail[i].GetValue('PSA_MOTIFSORTIE');
           Q:=OpenSQL('SELECT PMS_MAINOEUVRE FROM MOTIFSORTIEPAY WHERE PMS_CODE="'+MotifSortie+'"',True);
           If Not Q.Eof then MotifSortie := Q.FindField('PMS_MAINOEUVRE').AsString
           Else MotifSortie := '';
           Ferme(Q);
    end;
//    GMMO.CellValues[10,i+NbContrat+1]:=TobMMO.Detail[i].GetValue('PSA_MOTIFSORTIE');
    GMMO.CellValues[10,i+NbContrat+1]:=MotifSortie;
    end;
TobMMO.Free;
TobPaieMMO1.Free;
TobPaieMMO2.free;
if GMMO.RowCount>1 then GMMO.FixedRows:=1;
HMTrad.ResizeGridColumns (GMMO);
end;

procedure TOF_PGEDITMMOTOB.LancerEtatAvecTob(Sender:TObject);
var TobMMO,T:Tob;
    i:integer;
   Pages : TPageControl;
   StPages : String;
begin
Pages := TPageControl(GetControl('PAGES'));
TobMMO:=TOB.Create('SALARIES',Nil,-1) ;
TobMMO.AddChampSup('LIBELLE',False);
TobMMO.AddChampSup('PRENOM',False);
TobMMO.AddChampSup('NATIONALITE',False);
TobMMO.AddChampSup('DATENAISSANCE',False);
TobMMO.AddChampSup('SEXE',False);
TobMMO.AddChampSup('LIBEMPLOI',False);
TobMMO.AddChampSup('CODEEMPLOI',False);
TobMMO.AddChampSup('DATEENTREE',False);
TobMMO.AddChampSup('MOTIFENTREE',False);
TobMMO.AddChampSup('DATESORTIE',False);
TobMMO.AddChampSup('MOTIFSORTIE',False);
For i:=1 to GMMO.RowCount-1 do
    begin
    T:=Tob.Create('Fille MMO',TobMMO,i-1);
    T.AddChampSup('LIBELLE',False);
    T.AddChampSup('PRENOM',False);
    T.AddChampSup('NATIONALITE',False);
    T.AddChampSup('DATENAISSANCE',False);
    T.AddChampSup('SEXE',False);
    T.AddChampSup('LIBEMPLOI',False);
    T.AddChampSup('CODEEMPLOI',False);
    T.AddChampSup('DATEENTREE',False);
    T.AddChampSup('MOTIFENTREE',False);
    T.AddChampSup('DATESORTIE',False);
    T.AddChampSup('MOTIFSORTIE',False);
    T.putValue('LIBELLE',GMMO.CellValues[0,i]);
    T.putValue('PRENOM',GMMO.CellValues[1,i]);
    T.putValue('NATIONALITE',GMMO.CellValues[2,i]);
    T.putValue('DATENAISSANCE',GMMO.CellValues[3,i]);
    T.putValue('SEXE',GMMO.CellValues[4,i]);
    T.putValue('LIBEMPLOI',GMMO.CellValues[5,i]);
    T.putValue('CODEEMPLOI',GMMO.CellValues[6,i]);
    If IsValidDate(GMMO.CellValues[7,i]) then T.putValue('DATEENTREE',GMMO.CellValues[7,i])  //PT3
    Else T.putValue('DATEENTREE','');
    T.putValue('MOTIFENTREE',GMMO.CellValues[8,i]);
    If IsValidDate(GMMO.CellValues[9,i]) then T.putValue('DATESORTIE',GMMO.CellValues[9,i])  //PT3
    Else T.putValue('DATESORTIE','');
    T.putValue('MOTIFSORTIE',GMMO.CellValues[10,i]);
    end;
//TobMMO.GetGridDetail(GMMO,-1,'', 'PSA_LIBELLE;PSA_PRENOM;PSA_NATIONALITE;PSA_DATENAISSANCE;PSA_SEXE;CO_LIBRE;PSA_CODEEMPLOI;PSA_DATEENTREE;PSA_MOTIFENTREE;PSA_DATESORTIE;PSA_MOTIFSORTIE');
StPages := '';
{$IFDEF EAGLCLIENT}
StPages := AglGetCriteres (Pages, FALSE);  // PT2
{$ENDIF}
LanceEtatTOB('E','PAY','MOT',TOBMMO,True,False,False,Pages,'','',False,0,StPages);

TobMMO.Free;
end;

procedure TOF_PGEDITMMOTOB.GMMOColEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
If (ou=6) or (ou=7) or (ou=9) Then GMMO.ElipsisButton:=True
Else GMMO.ElipsisButton:=False;
TypeElipsis:='';
If ou=6 then TypeElipsis:='PCS';
If ou=7 then TypeElipsis:='Date';
If ou=9 then TypeElipsis:='Date';
end;

procedure TOF_PGEDITMMOTOB.GMMOElipsisClick(Sender: TObject);
var key : char;
    St:String;
    i:Integer;
    Q:TQuery;
begin
If TypeElipsis='Date' then
   begin
   key := '*';
   ParamDate (Ecran, Sender, Key);
   end;
If TypeElipsis='PCS' then
   begin
   if VH_Paie.PGPCS2003 then St := 'CO_TYPE="PZE"'  //PT1
   else St:='CO_TYPE = "PCS"';
   LookUpList (GMMO,'Les codes emploi','COMMUN','CO_ABREGE','CO_LIBRE',St,'CO_ABREGE',TRUE,-1);
   i:=GMMO.row;
   If GMMO.CellValues[6,i]<>'' Then
      begin
      Q:=OpenSQL('SELECT CO_LIBRE FROM COMMUN WHERE '+St+' AND CO_ABREGE="'+GMMO.CellValues[6,i]+'"',True);
      If not Q.eof then GMMO.CellValues[5,i]:=Q.FindField('CO_LIBRE').AsString;
      Ferme(Q);
      end;
   end;
end;

procedure TOF_PGEDITMMOTOB.AjoutLigne(Sender:TObject);
begin
If GMMO.RowCount>2 then GMMO.RowCount:=GMMO.RowCount+1
Else
   begin
   If GMMO.CellValues[0,1]<>'' then GMMO.RowCount:=GMMO.RowCount+1;
   end;
GMMO.CellValues[4,GMMO.RowCount-1]:='M';
GMMO.CellValues[7,GMMO.RowCount-1]:=DateToStr(DateFin);
//GMMO.CellValues[9,GMMO.RowCount-1]:=DateToStr(IDate1900);
SetControltext('XX_NBENTREE',IntToStr(StrToInt(GetControlText('XX_NBENTREE'))+1));
SetControltext('XX_EFFFINAL',IntToStr(StrToInt(GetControlText('XX_EFFFINAL'))+1));
SetControltext('XX_H',IntToStr(StrToInt(GetControlText('XX_H'))+1));
end;

procedure TOF_PGEDITMMOTOB.SupprimerLigne(Sender:TObject);
var i:Integer;
begin
i:=GMMO.row;
if i=0 then exit;
If GMMO.CellValues[4,i]='M' Then
   begin
   If (GMMO.CellValues[7,i]<>DateToStr(IDate1900)) and (GMMO.CellValues[7,i]<>'') then
      begin
      If (GMMO.CellValues[9,i]<>DateToStr(IDate1900)) and (GMMO.CellValues[9,i]<>'') then
         begin
         SetControltext('XX_NBENTREE',IntToStr(StrToInt(GetControlText('XX_NBENTREE'))-1));
         SetControltext('XX_NBSORTIES',IntToStr(StrToInt(GetControlText('XX_NBSORTIES'))-1));
         end
      Else
         begin
         SetControltext('XX_NBENTREE',IntToStr(StrToInt(GetControlText('XX_NBENTREE'))-1));
         SetControltext('XX_EFFFINAL',IntToStr(StrToInt(GetControlText('XX_EFFFINAL'))-1));
         SetControltext('XX_H',IntToStr(StrToInt(GetControlText('XX_H'))-1));
         end;
      end
   Else
      begin
      If (GMMO.CellValues[9,i]<>DateToStr(IDate1900)) and (GMMO.CellValues[9,i]<>'') then
         begin
         SetControltext('XX_NBSORTIES',IntToStr(StrToInt(GetControlText('XX_NBSORTIES'))-1));
         SetControltext('XX_EFFFINAL',IntToStr(StrToInt(GetControlText('XX_EFFFINAL'))+1));
         SetControltext('XX_H',IntToStr(StrToInt(GetControlText('XX_H'))+1));
         end
      Else
         begin
         end;
      end;
   end
Else
   begin
   If (GMMO.CellValues[7,i]<>DateToStr(IDate1900))and (GMMO.CellValues[7,i]<>'') then
      begin
      If (GMMO.CellValues[9,i]<>DateToStr(IDate1900)) and (GMMO.CellValues[9,i]<>'') then
         begin
         SetControltext('XX_NBENTREE',IntToStr(StrToInt(GetControlText('XX_NBENTREE'))-1));
         SetControltext('XX_NBSORTIES',IntToStr(StrToInt(GetControlText('XX_NBSORTIES'))-1));
         end
      Else
         begin
         SetControltext('XX_NBENTREE',IntToStr(StrToInt(GetControlText('XX_NBENTREE'))-1));
         SetControltext('XX_EFFFINAL',IntToStr(StrToInt(GetControlText('XX_EFFFINAL'))-1));
         SetControltext('XX_H',IntToStr(StrToInt(GetControlText('XX_F'))-1));
         end;
      end
   Else
      begin
      If (GMMO.CellValues[9,i]<>DateToStr(IDate1900)) and (GMMO.CellValues[9,i]<>'') then
         begin
         SetControltext('XX_NBSORTIES',IntToStr(StrToInt(GetControlText('XX_NBSORTIES'))-1));
         SetControltext('XX_EFFFINAL',IntToStr(StrToInt(GetControlText('XX_EFFFINAL'))+1));
         SetControltext('XX_H',IntToStr(StrToInt(GetControlText('XX_F'))+1));
         end
      Else
         begin
         end;
      end;
   end;
If (i = 1) and (GMMO.RowCount = 2) then GMMO.Rows[1].Clear
Else GMMO.DeleteRow(i);
end;

procedure TOF_PGEDITMMOTOB.Effectif(Sender: Tobject);
var DateEff,MDateEff,DateFin,DateDeb,DateDeb2 : TDateTime;
    etab,datelib1,datelib2,etablissement,St : String;
    QE,QSO,QEN,QET,QS,QInterim,Q : TQuery;
    TobEff,T : Tob;
    TotalEN,TotalSO,totalM,TotalF,Total : Integer;
    TotalENM,TotalSOM,TotalENF,TotalSOF : Integer;
    WhereEtab,E,WhereEtabPaie,WhereEtabSal,EtabCompl,WhereEtabPaie1,WhereEtabPaie2,WhereEtabinterim : String;
begin
        DateDeb := StrToDate(GetControlText('XX_VARIABLEDEB'));
        DateFin :=StrToDate(GetControlText('XX_VARIABLEFIN'));
        MDateEff := PlusMois(DateDeb,-1);
        DateEff := FinDeMois(MDateEff);
        DateDeb2 := DebutDeMois(DateEff);

        St := '(PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%'+GetControlText('ETAB')+'%") '+
        ' AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%DMO%" )  ' ;
        SetControlProperty('DECLARANT', 'Plus' ,St );

        QET:=OpenSql('SELECT PDA_DECLARANTATTES FROM DECLARANTATTEST '+
                   'WHERE (PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%'+GetControlText('ETAB')+'%") '+
                   'AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%DMO%" )  '+
                   'ORDER BY PDA_ETABLISSEMENT DESC',True);
        If Not QET.eof then
          Begin
          SetControlText('DECLARANT',QET.FindField('PDA_DECLARANTATTES').AsString);
//        AffichDeclarant(nil);
          End;
        Ferme(QET);
                E := GetControlText('ETAB');
                Total := 0;
                TotalM := 0;
                TotalF := 0;
                TotalEN := 0;
                TotalENM := 0;
                TotalENF := 0;
                TotalSO := 0;
                TotalSOM := 0;
                TotalSOF := 0;
                //NEW
                E := GetControlText('THVETABLISS');
                EtabCompl := GetControlText('ETABCOMPL');
                Etab := ReadTokenPipe(EtabCompl,';');
                While Etab <> '' do
                begin
                     WhereEtab := WhereEtab + ' OR PCI_ETABLISSEMENT="'+Etab+'"';
                     WhereEtabinterim := WhereEtabinterim + ' OR PEI_ETABLISSEMENT="'+Etab+'"';
                     WhereEtabPaie := WhereEtabPaie + ' OR PPU_ETABLISSEMENT="'+Etab+'"';
                     WhereEtabSal := WhereEtabSal + ' OR PSA_ETABLISSEMENT="'+Etab+'"';
                     WhereEtabPaie1 := WhereEtabPaie1 + ' AND P1.PPU_ETABLISSEMENT<>"'+Etab+'"';
                     WhereEtabPaie2 := WhereEtabPaie2 + ' OR P2.PPU_ETABLISSEMENT="'+Etab+'"';
                     Etab := ReadTokenPipe(EtabCompl,';');
                end;
                WhereEtabinterim := '(PEI_ETABLISSEMENT="'+E+'"' + WhereEtabinterim+')';
                WhereEtab := '(PCI_ETABLISSEMENT="'+E+'"' + WhereEtab+')';
                WhereEtabSal := '(PSA_ETABLISSEMENT="'+E+'"' + WhereEtabSal+')';
                WhereEtabPaie := '(PPU_ETABLISSEMENT="'+E+'"' + WhereEtabPaie+')';
                WhereEtabPaie1 := '(P1.PPU_ETABLISSEMENT<>"'+E+'"' + WhereEtabPaie1+')';
                WhereEtabPaie2 := '(P2.PPU_ETABLISSEMENT="'+E+'"' + WhereEtabPaie2+')';
                //Effectif
                         //Les contrats Modif requete PT14
                Q := OpenSQL('SELECT COUNT (DISTINCT PCI_SALARIE) EFFECTIF,PSA_SEXE FROM CONTRATTRAVAIL LEFT JOIN SALARIES ON PCI_SALARIE=PSA_SALARIE WHERE '+
                WhereEtab+' AND (PCI_DEBUTCONTRAT<"'+UsDateTime(DateDeb)+'") '+
                'AND (PCI_FINCONTRAT>="'+UsDateTime(DateDeb)+'" OR PCI_FINCONTRAT<="'+UsDateTime(IDate1900)+'"  OR PCI_FINCONTRAT IS NULL)'+
                ' AND PSA_PRISEFFECTIF="X" AND ((PSA_DADSPROF<>"09" AND PSA_DADSPROF<>"10" AND PSA_DADSPROF<>"11") OR PSA_DADSPROF="")'+
                ' GROUP BY PSA_SEXE',True);
                TobEff := Tob.Create('lesEffectifs',Nil,-1);
                TobEff.LoadDetailDB('LesEffectifs','','',Q,False);
                Ferme(Q);
                T := TobEff.FindFirst(['PSA_SEXE'],['M'],False);
                If T <> Nil then TotalM := TotalM + T.GetValue('EFFECTIF');
                T := TobEff.FindFirst(['PSA_SEXE'],['F'],False);
                If T <> Nil then TotalF := TotalF + T.GetValue('EFFECTIF');
                TobEff.Free;
                         //Les sans contrats Modif requete PT14
                Q := OpenSQL('SELECT COUNT(DISTINCT PSA_SALARIE) EFFECTIF,PSA_SEXE FROM SALARIES LEFT JOIN PAIEENCOURS ON PSA_SALARIE=PPU_SALARIE'+
                ' AND PPU_DATEDEBUT="'+UsDateTime(DateDeb)+'"'+
                ' WHERE PSA_PRISEFFECTIF="X" AND ((PSA_DADSPROF<>"09" AND PSA_DADSPROF<>"10" AND PSA_DADSPROF<>"11") OR PSA_DADSPROF="") AND'+
                ' PSA_DATEENTREE<"'+UsDateTime(DateDeb)+'"'+
                ' AND (PSA_DATESORTIE>"'+UsDateTime(DateDeb)+'" OR PSA_DATESORTIE IS NULL OR PSA_DATESORTIE="'+UsDateTime(idate1900)+'")'+
                ' AND ('+WhereEtabPaie+' OR ('+WhereEtabSal+' AND'+
                ' PPU_DATEDEBUT IS NULL))'+
                ' AND PSA_SALARIE NOT IN (SELECT PCI_SALARIE FROM CONTRATTRAVAIL WHERE '+
                WhereEtab+' AND (PCI_DEBUTCONTRAT<"'+UsDateTime(DateDeb)+'") '+
                'AND (PCI_FINCONTRAT>"'+UsDateTime(DateDeb)+'" OR PCI_FINCONTRAT<="'+UsDateTime(IDate1900)+'"  OR PCI_FINCONTRAT IS NULL))'+
                ' GROUP BY PSA_SEXE',True);
                TobEff := Tob.Create('lesEffectifs',Nil,-1);
                TobEff.LoadDetailDB('LesEffectifs','','',Q,False);
                Ferme(Q);
                T := TobEff.FindFirst(['PSA_SEXE'],['M'],False);
                If T <> Nil then TotalM := TotalM + T.GetValue('EFFECTIF');
                T := TobEff.FindFirst(['PSA_SEXE'],['F'],False);
                If T <> Nil then TotalF := TotalF + T.GetValue('EFFECTIF');
                TobEff.Free;
                Total := TotalM + TotalF;
                SetControlText('XX_EFFTOTAL',IntToStr(Total));

                //Les entrées
                        //Les contrats
                Q := OpenSQL('SELECT COUNT (PCI_SALARIE) NBENTREECT,PSA_SEXE FROM CONTRATTRAVAIL LEFT JOIN SALARIES ON PCI_SALARIE=PSA_SALARIE WHERE '+
                WhereEtab+' AND PCI_DEBUTCONTRAT>="'+UsDateTime(DateDeb)+'" AND PCI_DEBUTCONTRAT<="'+UsDateTime(DateFin)+'"'+
                ' GROUP BY PSA_SEXE',True);
                TobEff := Tob.Create('lesEffectifs',Nil,-1);
                TobEff.LoadDetailDB('LesEffectifs','','',Q,False);
                Ferme(Q);
                T := TobEff.FindFirst(['PSA_SEXE'],['M'],False);
                If T <> Nil then TotalENM := TotalENM + T.GetValue('NBENTREECT');
                T := TobEff.FindFirst(['PSA_SEXE'],['F'],False);
                If T <> Nil then TotalENF := TotalENF + T.GetValue('NBENTREECT');
                TobEff.Free;
                         //Les sans contrats
                Q := OpenSQL('SELECT COUNT (DISTINCT PSA_SALARIE) NBENTREE,PSA_SEXE FROM SALARIES LEFT JOIN PAIEENCOURS P1 ON PSA_SALARIE=P1.PPU_SALARIE'+//PT6
                ' WHERE PSA_PRISEFFECTIF="X" AND ((PSA_DADSPROF<>"09" AND PSA_DADSPROF<>"10" AND PSA_DADSPROF<>"11") OR PSA_DADSPROF="")'+ //PT12
                ' AND (('+WhereEtabSal+' AND PSA_DATEENTREE<="'+UsDateTime(DateFin)+'" AND PSA_DATEENTREE>="'+UsDateTime(DateDeb)+'")'+
                ' OR (Exists (SELECT PPU_SALARIE FROM PAIEENCOURS P2 WHERE P2.PPU_SALARIE=PSA_SALARIE AND'+
                ' P1.PPU_DATEDEBUT="'+UsDateTime(DateDeb2)+'" AND P2.PPU_DATEDEBUT="'+UsDateTime(DateDeb)+'"'+
                ' AND '+WhereEtabPaie1+' AND P1.PPU_ETABLISSEMENT IS NOT NULL AND '+WhereEtabPaie2+'))) '+
                ' AND PSA_SALARIE NOT IN (SELECT PCI_SALARIE FROM CONTRATTRAVAIL WHERE '+
                WhereEtab+' AND PCI_DEBUTCONTRAT>="'+UsDateTime(DateDeb)+'" AND PCI_DEBUTCONTRAT<="'+UsDateTime(DateFin)+'")'+ //PT14
                ' GROUP BY PSA_SEXE',True);
                TobEff := Tob.Create('lesEffectifs',Nil,-1);
                TobEff.LoadDetailDB('LesEffectifs','','',Q,False);
                Ferme(Q);
                T := TobEff.FindFirst(['PSA_SEXE'],['M'],False);
                If T <> Nil then TotalENM := TotalENM + T.GetValue('NBENTREE');
                T := TobEff.FindFirst(['PSA_SEXE'],['F'],False);
                If T <> Nil then TotalENF := TotalENF + T.GetValue('NBENTREE');
                TobEff.Free;
                TotalEn := TotalEnM + TotalEnF;
                SetControlText('XX_NBENTREE',IntToStr(TotalEn));
                //Les sorties
                EtabCompl := GetControlText('ETABCOMPL');
                Etab := ReadTokenPipe(EtabCompl,';');
                WhereEtabPaie2 := '';
                WhereEtabPaie1 := '';
                While Etab <> '' do
                begin
                     WhereEtabPaie1 := WhereEtabPaie1 + ' OR P1.PPU_ETABLISSEMENT="'+Etab+'"';
                     WhereEtabPaie2 := WhereEtabPaie2 + ' AND P2.PPU_ETABLISSEMENT<>"'+Etab+'"';
                     Etab := ReadTokenPipe(EtabCompl,';');
                end;
                WhereEtabPaie1 := '(P1.PPU_ETABLISSEMENT="'+E+'"' + WhereEtabPaie1+')';
                WhereEtabPaie2 := '(P2.PPU_ETABLISSEMENT<>"'+E+'"' + WhereEtabPaie2+')';
                Q := OpenSQL('SELECT COUNT (PCI_SALARIE) NBSORTIECT,PSA_SEXE FROM CONTRATTRAVAIL LEFT JOIN SALARIES ON PCI_SALARIE=PSA_SALARIE WHERE '+
                'PSA_PRISEFFECTIF="X" AND ((PSA_DADSPROF<>"09" AND PSA_DADSPROF<>"10" AND PSA_DADSPROF<>"11") OR PSA_DADSPROF="") AND '+
                WhereEtab+' AND PCI_FINCONTRAT>="'+UsDateTime(DateDeb)+'" '+
                'AND PCI_FINCONTRAT<"'+UsDateTime(DateFin)+'"'+
                ' GROUP BY PSA_SEXE',True);
                TobEff := Tob.Create('lesEffectifs',Nil,-1);
                TobEff.LoadDetailDB('LesEffectifs','','',Q,False);
                Ferme(Q);
                T := TobEff.FindFirst(['PSA_SEXE'],['M'],False);
                If T <> Nil then TotalSOM := TotalSOM + T.GetValue('NBSORTIECT');
                T := TobEff.FindFirst(['PSA_SEXE'],['F'],False);
                If T <> Nil then TotalSOF := TotalSOF + T.GetValue('NBSORTIECT');
                TobEff.Free;
                Q := OpenSQL('SELECT COUNT (DISTINCT PSA_SALARIE) NBSORTIE,PSA_SEXE FROM SALARIES LEFT JOIN PAIEENCOURS P1 ON PSA_SALARIE=P1.PPU_SALARIE'+  //PT6
                ' WHERE PSA_PRISEFFECTIF="X" AND ((PSA_DADSPROF<>"09" AND PSA_DADSPROF<>"10" AND PSA_DADSPROF<>"11") OR PSA_DADSPROF="")'+  //PT12
                ' AND (('+WhereEtabSal+' AND PSA_DATESORTIE<="'+UsDateTime(DateFin)+'" AND PSA_DATESORTIE>="'+UsDateTime(DateDeb)+'")'+
                ' OR (EXISTS (SELECT PPU_SALARIE FROM PAIEENCOURS P2 WHERE P2.PPU_SALARIE=PSA_SALARIE AND'+
                ' P1.PPU_DATEDEBUT="'+UsDateTime(DateDeb2)+'" AND P2.PPU_DATEDEBUT="'+UsDateTime(DateDeb)+'"'+
                ' AND '+WhereEtabPaie1+' AND P2.PPU_ETABLISSEMENT IS NOT NULL AND '+WhereEtabPaie2+')))'+
                ' AND PSA_SALARIE NOT IN (SELECT PCI_SALARIE FROM CONTRATTRAVAIL WHERE '+
                WhereEtab+' AND (PCI_FINCONTRAT>="'+UsDateTime(datedeb)+'" AND PCI_FINCONTRAT<"'+UsDateTime(datefin)+'"))'+
                ' GROUP BY PSA_SEXE',True);
                TobEff := Tob.Create('lesEffectifs',Nil,-1);
                TobEff.LoadDetailDB('LesEffectifs','','',Q,False);
                Ferme(Q);
                T := TobEff.FindFirst(['PSA_SEXE'],['M'],False);
                If T <> Nil then TotalSOM := TotalSOM + T.GetValue('NBSORTIE');
                T := TobEff.FindFirst(['PSA_SEXE'],['F'],False);
                If T <> Nil then TotalSOF := TotalSOF + T.GetValue('NBSORTIE');
                TobEff.Free;
                TotalSo := TotalSoF + TotalSoM;
                SetControlText('XX_NBSORTIES',IntToStr(TotalSo));
                SetControlText('XX_EFFFINAL',IntToStr(Total + TotalEn - TotalSO));
                SetControlText('XX_H',IntToStr(TotalM + TotalEnM - TotalSOM));
                SetControlText('XX_F',IntToStr(TotalF + TotalEnF - TotalSOF));
               //FIN NEW
                //Debut PT- 8
                QInterim := OpenSQL('SELECT COUNT (DISTINCT PEI_INTERIMAIRE) AS INTERIMAIRE FROM EMPLOIINTERIM LEFT JOIN INTERIMAIRES ON PSI_INTERIMAIRE=PEI_INTERIMAIRE'+
                ' WHERE '+WhereEtabinterim+' AND PEI_DEBUTEMPLOI<="'+UsDateTime(DateFin)+'" AND (PEI_FINEMPLOI="'+UsdateTime(IDate1900)+'" OR PEI_FINEMPLOI>="'+UsDateTime(DateFin)+'") AND PSI_TYPEINTERIM="INT"',True);
                if not QInterim.eof then SetControlText('XX_NBINTERIM',QInterim.FindField('INTERIMAIRE').AsString); // PortageCWAS
                Ferme(QInterim);
                // FIN PT- 8
        AfficheGrille;
end;




procedure TOF_PGEDITMMOTOB.DateElipsisclick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGEDITMMOTOB.ExitDateDeb(Sender:TObject);
var DateFinPrec : TDateTime;
begin
DateDeb:=StrToDate(GetControlText('XX_VARIABLEDEB'));
DateFin:=FinDeMois(DateDeb);
SetControlText('XX_VARIABLEFIN',DateToStr(DateFin));
DateFinPrec := FinDeMois(PlusMois(DateDeb,-1)); //PT5
SetControlcaption('LIBEFF','Effectif au ' +DateToStr(DateFinPrec)); 
SetControlCaption('LIBFIN','Effectif au ' +DatEtoStr(DateFin) );
Effectif(Nil);
end;

procedure TOF_PGEDITMMOTOB.OnCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
If ACol=4 then SexeAvModif := GMMO.CellValues[ACol,ARow];
If ACol=7  then
begin
        If IsValidDate(GMMO.CellValues[ACol,ARow]) then EntreeAvModif := StrTodate( GMMO.CellValues[ACol,ARow])
        else EntreeAvModif := IDate1900;
end;
If ACol=9  then
begin
        If IsValidDate(GMMO.CellValues[ACol,ARow]) then SortieAvModif := StrTodate(GMMO.CellValues[ACol,ARow])
        else SortieAvModif := IDate1900;
end;
end;

procedure TOF_PGEDITMMOTOB.OnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var DateApModif,DebutMois,FinMois : TDateTime;
    SexeApModif,St,TestDate : String;
    NBh,NbF,NbEntree,NbSortie,NbFin : Integer;
begin
If ACol = 4 then
begin
        SexeApModif := GMMO.CellValues[ACol,ARow];
        If SexeAvModif <> SexeApModif then
        begin
                NbF := StrToInt(GetControlText('XX_F'));
                NbH := StrToInt(GetControlText('XX_H'));
                If SexeApModif ='F' then begin NbF := NbF + 1;NbH := NbH - 1; end
                else begin NbF := NbF - 1;NbH := NbH + 1; end;
                SetControlText('XX_H',IntToStr(NbH));
                SetControlText('XX_F',IntToStr(NbF));
        end;
end;
If ACol = 7 then
begin
        DebutMois := StrTodate(GetControlText('XX_VARIABLEDEB'));
        FinMois := StrTodate(GetControlText('XX_VARIABLEFIN'));
        TestDate := GMMO.CellValues[ACol,ARow];
        If IsValidDate(GMMO.CellValues[ACol,ARow]) then DateApModif := StrTodate(GMMO.CellValues[ACol,ARow])
        else
        begin
                If (TestDate <> '  /  /    ') and (TestDate<>'') then  //PT1
                begin
                        PGIBox(TestDate+' n''est pas une date correcte',Ecran.Caption);
                        GMMO.Col := ACol;
                        GMMO.Row := ARow;
                        Exit;
                end;
                GMMO.CellValues[ACOL,ARow] := '';
                DateApModif := IDate1900;
        end;
        If EntreeAvModif <> DateApModif then
        begin
                If (EntreeAvModif >= DebutMois) and (EntreeAvModif <= FinMois) then
                begin
                        If (DateApModif < DebutMois) or (DateApModif > FinMois) then
                        begin
                                NbFin := StrToInt(GetControlText('XX_EFFFINAL'));
                                NbEntree := StrToInt(GetControlText('XX_NBENTREE'));
                                NbEntree := NbEntree - 1;
                                NbFin := NbFin - 1;
                                SetControlText('XX_EFFFINAL',IntToStr(NbFin));
                                SetControlText('XX_NBENTREE',IntToStr(NbEntree));
                        end;
                end;
                If (DateApModif >= DebutMois) and (DateApModif <= FinMois) then
                begin
                        If (EntreeAvModif < DebutMois) or (EntreeAvModif > FinMois) then
                        begin
                                NbFin := StrToInt(GetControlText('XX_EFFFINAL'));
                                NbEntree := StrToInt(GetControlText('XX_NBENTREE'));
                                NbEntree := NbEntree + 1;
                                NbFin := NbFin + 1;
                                SetControlText('XX_EFFFINAL',IntToStr(NbFin));
                                SetControlText('XX_NBENTREE',IntToStr(NbEntree));
                        end;
                end;
        end;
end;
If ACol = 9 then
begin
        DebutMois := StrTodate(GetControlText('XX_VARIABLEDEB'));
        FinMois := StrTodate(GetControlText('XX_VARIABLEFIN'));
        TestDate := GMMO.CellValues[ACol,ARow];
        If IsValidDate(GMMO.CellValues[ACol,ARow]) then DateApModif := StrTodate(GMMO.CellValues[ACol,ARow])
        else
        begin
                If (TestDate <> '  /  /    ') and (TestDate<>'') then  //PT1
                begin
                        PGIBox(TestDate+' n''est pas une date correcte',Ecran.Caption);
                        GMMO.Col := ACol;
                        GMMO.Row := ARow;
                        Exit;
                end;
                GMMO.CellValues[ACol,ARow] := '';
                DateApModif := IDate1900;
        end;
        If SortieAvModif <> DateApModif then
        begin
                If (SortieAvModif >= DebutMois) and (SortieAvModif <= FinMois) then
                begin
                        If (DateApModif < DebutMois) or (DateApModif > FinMois) then
                        begin
                                NbFin := StrToInt(GetControlText('XX_EFFFINAL'));
                                NbSortie := StrToInt(GetControlText('XX_NBSORTIES'));
                                NbSortie := NbSortie - 1;
                                NbFin := NbFin + 1;
                                SetControlText('XX_EFFFINAL',IntToStr(NbFin));
                                SetControlText('XX_NBSORTIES',IntToStr(NbSortie));
                        end;
                end;
                If (DateApModif >= DebutMois) and (DateApModif <= FinMois) then
                begin
                        If (SortieAvModif < DebutMois) or (SortieAvModif > FinMois) then
                        begin
                                NbFin := StrToInt(GetControlText('XX_EFFFINAL'));
                                NbSortie := StrToInt(GetControlText('XX_NBSORTIES'));
                                NbSortie := NbSortie + 1;
                                NbFin := NbFin - 1;
                                SetControlText('XX_EFFFINAL',IntToStr(NbFin));
                                SetControlText('XX_NBSORTIES',IntToStr(NbSortie));
                        end;
                end;
        end;
end;
If ACol = 6 then  //PT1
begin
            If GMMO.CellValues[ACol,ARow] <> '' then
            begin
                if VH_Paie.PGPCS2003 then St := 'CO_TYPE="PZE"'
                else St:='CO_TYPE = "PCS"';
                If Not ExisteSQL('SELECT CO_CODE FROM COMMUN WHERE '+St+' AND CO_ABREGE="'+GMMO.CellValues[ACol,ARow]+'"') then
                begin
                        GMMO.Col := ACol;
                        GMMO.Row := ARow;
                        PGIBox('Ce code n''existe pas',Ecran.Caption);
                end;
            end;
end;
end;

Initialization
  registerclasses ( [ TOF_PGEDITMMOTOB ] ) ;
end.
