{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 16/11/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PGASSECIC_SPECTACLE
*****************************************************************
PT1 : 29/03/2002 : JL 571 : - Mise en commentaire appel procedure ChangeListe
                            - Nouvel gestion avec données de la table deportsal
PT2 : 14/09/2005 : JL V_60 : FQ 12572 Correction requête mul pour oracle
PT3 : 30/08/2006 : JL V_70 FQ 13472 Ajout impression motif fin contart
PT4 : 28/09/2006 : JL V_75 FQ 13551 Ajout libellé de l'emploi au lieu code
PT5 : 09/10/2006 : JL V_75 Gestion des réalisateurs
PT6 : 06/06/2007 : JL V_72 Gestion affichage salari&és avec buleltin de paie
PT7 : 10/10/2007 : RM V_80 Mise à jour pour prendre en compte le formulaire V4
PT8 : 03/12/2007 : RM V_80 FQ14020 Compléter la saisie du matricule salarié
PT9 : 07/01/2008 : RM V_80 FQ15095 Correction alimentation de la TobEdtion
}
unit UTOFPGMULASSEDIC_spectacle;

Interface
uses     UTOF,
{$IFNDEF EAGLCLIENT}
         Mul,db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}EdtREtat, hdb,
{$ELSE}
       eMul,UtilEAGL,
{$ENDIF}
         Hctrls,HEnt1,StdCtrls,Classes,HTB97,ParamDat,HMsgBox, Forms,Controls,
         sysutils,PGOutils2,EntPaie,P5Def,UTOB,HQry,HStatus,ed_tools,ComCtrls;



Type
  TOF_PGMULASSEDIC_SPECTACLE = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (stArgument : String ) ; override ;
    private
    Argument:String;
   CListe:THValComboBox;
   EMatricule:THEdit;
   procedure ExitEdit(Sender: TObject);
 //  procedure ChangeListe(Sender:TObject);  //PT1
   procedure DateElipsisclick(Sender: TObject);
   Procedure OnClickSalarieSortie(Sender: TObject);
end;
Type
  TOF_PGMULEDITASSEDICSPECTACLE = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (stArgument : String ) ; override ;
    private
    procedure EditionAttest(Sender : TOBject);
    procedure ExitEdit(Sender: TObject);  //PT8
end;
Type
  TOF_PGMULSALSPECTACLE = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (stArgument : String ) ; override ;
    private
    procedure SuppAttestation(Sender : TObject);
    procedure EditionAttest(Sender : TOBject);
end;


Implementation


procedure TOF_PGMULASSEDIC_SPECTACLE.OnLoad ;
var Clause,Fait : String;
    //PT7 Inclus:TCheckBox;
    AD,AF,MD,MF,jj : Word;
    DD,DF : TDateTime;
    AnneeD,AnneeF,MoisF,MoisD : String;
    Q : TQuery;
    DateArret : TDateTime;
begin
  Inherited ;
if Argument = 'AS' then
       Ecran.Caption := 'Attestation ASSEDIC spectacles';
if Argument = 'CP' then
       Ecran.Caption := 'Congés spectacles';
UpdateCaption(TFmul(Ecran));
Clause:='';
If CListe.Value='A' Then Clause:='PSE_INTERMITTENT="X"'
Else Clause:='';
if TCheckBox(GetControl('CKSORTIE'))<>nil then //PT17
Begin
     if (GetControlText('CKSORTIE')='X') and
     (IsValidDate(GetControlText('DATEARRET'))) then
     Begin
          DateArret:= StrtoDate(GetControlText('DATEARRET'));
          Clause:= Clause + ' AND PSA_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE '+
          '(PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR'+
          ' PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR'+
          ' PSA_DATESORTIE IS NULL) AND'+
          ' PSA_DATEENTREE <="'+UsDateTime(DateArret)+'")';
     End;
End;
Fait := '';
DD := StrToDate(GetControlText('DATEDEBUT'));
DF := StrToDate(GetControlText('DATEFIN'));
DecodeDate(DD, AD, MD, jj);
DecodeDate(DF, AF, MF, jj);
If (AD = AF) and (MD = MF) then
begin
     Q := OpenSQL('Select CO_CODE FROM COMMUN Where CO_TYPE="PGA" AND CO_LIBELLE="'+IntToStr(AD)+'"',True);
     If not Q.Eof then AnneeD := Q.FindField('CO_CODE').AsString
     else AnneeD := '';
     Ferme(Q);
     MoisD := ColleZeroDevant(MD,2);
     If GetCheckBoxState('ATTESTFAIT') = CbChecked then Fait := 'PSE_SALARIE IN (SELECT PGA_SALARIE FROM PGAEM WHERE '+
     'PGA_MOIS="'+MoisD+'" AND PGA_ANNEE="'+AnneeD+'")'
     else Fait := 'PSE_SALARIE NOT IN (SELECT PGA_SALARIE FROM PGAEM WHERE '+
     'PGA_MOIS="'+MoisD+'" AND PGA_ANNEE="'+AnneeD+'")';
end
else
begin
     Q := OpenSQL('Select CO_CODE FROM COMMUN Where CO_TYPE="PGA" AND CO_LIBELLE="'+IntToStr(AD)+'"',True);
     If not Q.Eof then AnneeD := Q.FindField('CO_CODE').AsString
     else AnneeD := '';
     Ferme(Q);
     Q := OpenSQL('Select CO_CODE FROM COMMUN Where CO_TYPE="PGA" AND CO_LIBELLE="'+IntToStr(AF)+'"',True);
     If not Q.Eof then AnneeF := Q.FindField('CO_CODE').AsString
     else AnneeF := '';
     Ferme(Q);
     MoisF := ColleZeroDevant(MF,2);
     If GetCheckBoxState('ATTESTFAIT') = CbChecked then Fait := 'PSE_SALARIE IN (SELECT PGA_SALARIE FROM PGAEM WHERE '+
     'PGA_MOIS>="'+MoisD+'" AND PGA_ANNEE>="'+AnneeD+'" AND PGA_MOIS<="'+MoisF+'" AND PGA_ANNEE<="'+AnneeF+'")'
     else Fait := 'PSE_SALARIE NOT IN (SELECT PGA_SALARIE FROM PGAEM WHERE '+
     'PGA_MOIS>="'+MoisD+'" AND PGA_ANNEE>="'+AnneeD+'" AND PGA_MOIS<="'+MoisF+'" AND PGA_ANNEE<="'+AnneeF+'")'
end;
If Clause <> '' then Clause := Clause + ' AND '+ Fait
else Clause := Fait;
//DEBUT PT6
If GetCheckBoxState('CBULLETIN') = CbChecked then
begin
  Clause := Clause + ' AND PSA_SALARIE IN (SELECT DISTINCT(PPU_SALARIE) FROM PAIEENCOURS WHERE PPU_DATEFIN>="'+UsDateTime(StrToDate(GetControlText('DATEDEBUT')))+'"'+
  ' AND PPU_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('DATEFIN')))+'")';
end;
//FIN PT6
SetControlText('XX_WHERE',Clause);
end ;

procedure TOF_PGMULASSEDIC_SPECTACLE.OnArgument (stArgument : String ) ;
Var Num : Integer;
    EWhere,Defaut:THEdit;
    DateAttestation:TDateTime;
    Check : TCheckBox;
begin
Inherited ;
Argument :=stArgument;
CListe:=THValComboBox(GetControl('CLISTE'));
//CListe.Value:='';
//If CListe<>Nil Then CListe.OnChange:=ChangeListe;  //PT1
CListe.Value:='A';
EWhere:=THEdit(GetControl('XX_WHERE'));
EWhere.Text:='PSE_INTERMITTENT="X"';
EMatricule:=ThEdit(getcontrol('PSA_SALARIE'));
If EMatricule<>nil then EMatricule.OnExit:=ExitEdit;
For Num := 1 to VH_Paie.PGNbreStatOrg do
         begin
         if Num >4 then Break;
         VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
         end;
VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ;
DateAttestation:=PlusMois(Date,-1) ;
SetControlText('DATEDEBUT',DateToStr(DebutDeMois(DateAttestation)));
SetControlText('DATEFIN',DateToStr(FinDeMois(DAteAttestation)));
Defaut:=THEdit(GetControl('DATEDEBUT'));
If Defaut<>Nil Then Defaut.OnElipsisClick:=DateElipsisClick;
Defaut:=THEdit(GetControl('DATEFIN'));
If Defaut<>Nil Then Defaut.OnElipsisClick:=DateElipsisClick;
Check:= TCheckBox(GetControl('CKSORTIE'));  //PT17
         if Check <> nil then Check.OnClick:= OnClickSalarieSortie;
end;

procedure TOF_PGMULASSEDIC_SPECTACLE.OnClickSalarieSortie(Sender: TObject);
begin
SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
end;


{procedure TOF_PGMULASSEDIC_SPECTACLE.ChangeListe(Sender:TObject);
Var EWhere:THEdit;
begin
EWhere:=THEdit(GetControl('XX_WHERE'));
If CListe.Value='A' Then EWhere.Text:=' PSA_DADSPROF="05"'
Else EWhere.Text:='';
end;}

procedure TOF_PGMULASSEDIC_SPECTACLE.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

procedure TOF_PGMULASSEDIC_SPECTACLE.DateElipsisclick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

{TOF_PGMULEDITASSEDICSPECTACLE}
procedure TOF_PGMULEDITASSEDICSPECTACLE.OnLoad;
begin
Inherited ;
end;

procedure TOF_PGMULEDITASSEDICSPECTACLE.OnArgument (stArgument : String ) ;
var DateAttestation : TDateTime;
    Q,QAttes : TQuery;
    aa,mm,jj : Word;
    BImp : TToolBarButton97;
    Defaut : THEdit;  //PT8
begin
Inherited ;
   Defaut:=ThEdit(getcontrol('PGA_SALARIE'));      //PT8
   If Defaut<>nil then Defaut.OnExit:=ExitEdit;    //PT8

   DateAttestation:=PlusMois(V_PGI.DATEEntree,-1) ;
   DecodeDate(DateAttestation, aa, mm, jj);
   Q := OpenSQL('SELECT CO_CODE FROM COMMUN WHERE CO_TYPE="PGA" AND CO_LIBELLE="'+IntToStr(aa)+'"',True);
   If Not Q.Eof then SetControlText('PGA_ANNEE',Q.FindField('CO_CODE').AsString);
   Ferme(Q);
   SetControlText('PGA_MOIS',IntToStr(mm));
   BImp := TToolBarButton97(Getcontrol('BImprimer'));
   if BImp <> Nil then BImp.OnClick := EditionAttest;
   SetControlVisible('BOuvrir',False);
   QAttes := OpenSql('SELECT PDA_DECLARANTATTES FROM DECLARANTATTEST ' +
      'WHERE (PDA_ETABLISSEMENT = "") ' +
      'AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%ASP%" )  ' +
      'ORDER BY PDA_ETABLISSEMENT DESC', True);
   if not QAttes.eof then
   begin
        SetControlText('DECLARANT', QAttes.FindField('PDA_DECLARANTATTES').AsString);
   end;
   Ferme(QAttes);
end;

procedure TOF_PGMULEDITASSEDICSPECTACLE.EditionAttest(Sender : TOBject);
var Q : TQuery;
    TobEtab,TobEdition,TEt,T : Tob;
    Salarie,Etab,NumDoc ,Categ,St: String;
     {$IFNDEF EAGLCLIENT}
  Liste: THDBGrid;
  {$ELSE}
  Liste: THGrid;
  {$ENDIF}
  Q_Mul: THQuery;
  i : Integer;
  Pages : TPageControl;
  ImpSal,ImpAssedic,ImpDupli : Boolean;
  EtatsAImprimer,Etat : String;
  mm,yy : word;
  DatePeriodeEC : TDateTime;
  An : String;
  CategIntermittent : String;
  AjoutV4,Lparent,BrutSap1,Taux1,Cotisation1 : String; //PT7

begin
  {$IFNDEF EAGLCLIENT}
  Liste := THDBGrid(GetControl('FLISTE'));
  {$ELSE}
  Liste := THGrid(GetControl('FLISTE'));
  {$ENDIF}
  if Liste = nil then Exit;
   Q_Mul := THQuery(Ecran.FindComponent('Q'));
  if Q_Mul = nil then Exit;
  if ((Liste.nbSelected) <= 0) and (not Liste.AllSelected) then
  begin
      PGIBox('Aucune attestation sélectionnée','Edition des attestations');
      Exit;
  end;
  Pages := TPageControl(GetControl('PAGES'));
  Q := OpenSQL('SELECT * FROM ETABLISS LEFT JOIN ETABCOMPL ON ET_ETABLISSEMENT=ETB_ETABLISSEMENT',True);
  TobEtab := Tob.Create('LesEtab',Nil,-1);
  TobEtab.LoadDetailDB('lesEtabs','','',Q,False);
  Ferme(Q);
  TobEdition := Tob.Create('Edition',Nil,-1);
  if ((Liste.nbSelected) > 0) and (not Liste.AllSelected) then
  begin
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', Liste.nbSelected, FALSE, TRUE);
    InitMove(Liste.nbSelected, '');
    for i := 0 to Liste.NbSelected - 1 do
    begin
        Liste.GotoLeBOOKMARK(i);
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(Liste.Row - 1);
        {$ENDIF}
        T := Tob.Create('LigneEdition',TobEdition,-1);
        Salarie := TFmul(Ecran).Q.FindField('PGA_SALARIE').asstring;
        NumDoc := TFmul(Ecran).Q.FindField('PGA_NUMATTEST').asstring;
        Q := OpenSQL('SELECT * FROM PGAEM LEFT JOIN SALARIES ON PGA_SALARIE=PSA_SALARIE '+
        'LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+
        'WHERE PGA_SALARIE="'+Salarie+'" AND PGA_NUMATTEST="'+NumDoc+'"',True);

        //PT7 ajout du UpperCase lignes ci-dessous **
        T.AddChampSupValeur('PSA_SALARIE',Salarie);
        T.AddChampSupValeur('ENUMDOC',UpperCase(NumDoc));
        T.AddChampSupValeur('PGA_NUMATTEST',UpperCase(NumDoc));
        If not Q.Eof then
        begin
            Etab := Q.FindField('PSA_ETABLISSEMENT').AsString;
            Categ := Q.FindField('PGA_ISCATEG').AsString;
            CategIntermittent := Q.FindField('PSE_ISCATEG').AsString;
            //PT7 Debut Modif ===>
            If Q.FindField('PSA_NOMJF').AsString = '' then
            Begin
               T.AddChampSupValeur('LNOM',UpperCase(Q.FindField('PSA_LIBELLE').AsString));
               T.AddChampSupValeur('ENOMUSAGE',UpperCase(Q.FindField('PSA_SURNOM').AsString));
            End
            Else
            Begin
               T.AddChampSupValeur('LNOM',UpperCase(Q.FindField('PSA_NOMJF').AsString));
               T.AddChampSupValeur('ENOMUSAGE',UpperCase(Q.FindField('PSA_LIBELLE').AsString));
            End;
            //PT7 Fin Modif <===
            T.AddChampSupValeur('LPRENOM',UpperCase(Q.FindField('PSA_PRENOM').AsString));
            T.AddChampSupValeur('CATTESTINIT','-');
            T.AddChampSupValeur('ATTESTCOMPL','-');
            T.AddChampSupValeur('CRECAPPOS','-');
            T.AddChampSupValeur('RECAPNEG','-');
            iF Q.FindField('PGA_TYPEAEM').AsString = 'COM' then T.PutValue('ATTESTCOMPL','X')
            else iF Q.FindField('PGA_TYPEAEM').AsString = 'NEG' then T.PutValue('RECAPNEG','X')
            else iF Q.FindField('PGA_TYPEAEM').AsString = 'POS' then T.PutValue('CRECAPPOS','X')
            else T.PutValue('CATTESTINIT','X');
            T.AddChampSupValeur('EEMPLOI',UpperCase(Q.FindField('PGA_LIBELLEEMPLOI').AsString)); //PT4
            //T.AddChampSupValeur('ERETRAITE',Q.FindField('PSE_ISRETRAITE').AsString);           //PT7
            T.AddChampSupValeur('ERETRAITE',UpperCase(Q.FindField('PGA_ISRETRAITE').AsString));  //PT7
            T.AddChampSupValeur('ENUMASSEDIC',Q.FindField('PSE_ISNUMASSEDIC').AsString);
            //T.AddChampSupValeur('PSA_NUMEROSS',Q.FindField('PSA_NUMEROSS').AsString);  //PT7
            T.AddChampSupValeur('ENUMIDENTIFIANT',UpperCase(Q.FindField('PSA_NUMEROSS').AsString));//PT7
            T.AddChampSupValeur('EDATENAISSANCE',DateToStr(Q.FindField('PSA_DATENAISSANCE').AsDateTime));
            T.AddChampSupValeur('ETELSALARIE',Q.FindField('PSA_TELEPHONE').AsString);
            //T.AddChampSupValeur('PSA_CODEPOSTAL',Q.FindField('PSA_CODEPOSTAL').AsString);   //PT7
            T.AddChampSupValeur('ECPSAL',UpperCase(Q.FindField('PSA_CODEPOSTAL').AsString));  //PT7
            //T.AddChampSupValeur('PSA_VILLE',Q.FindField('PSA_VILLE').AsString);             //PT7
            T.AddChampSupValeur('ECOMMUNESAL',UpperCase(Q.FindField('PSA_VILLE').AsString));  //PT7
            T.AddChampSupValeur('EADRSAL1',UpperCase(Q.FindField('PSA_ADRESSE1').AsString));
            T.AddChampSupValeur('EADRSAL2',UpperCase(Q.FindField('PSA_ADRESSE2').AsString) + UpperCase(Q.FindField('PSA_ADRESSE3').AsString));
            //T.AddChampSupValeur('ENOMUSAGE',Q.FindField('PSA_SURNOM').AsString);          //PT7
            //T.AddChampSupValeur('EDATENAISSANCE',DateToStr(Q.FindField('PSA_DATENAISSANCE').AsDateTime));  //PT7
            T.AddChampSupValeur('MOIS',Q.FindField('PGA_MOIS').AsString);
            //DEBUT PT3
            T.AddChampSupValeur('CFINCDD', Q.FindField('PGA_FINCDD').AsString);
            T.AddChampSupValeur('CRUPTEMPL', Q.FindField('PGA_RUPTEMP').AsString);
            T.AddChampSupValeur('CRUPTSAL', Q.FindField('PGA_RUPTSAL').AsString);
            //FIN PT3
            T.AddChampSupValeur('ANNEE',RechDom('PGANNEE',Q.FindField('PGA_ANNEE').AsString,False));
            T.AddChampSupValeur('NUMATTESTINIT',UpperCase(Q.FindField('PGA_NUMATTESTINIT').AsString));
//            EtabSalarie := Q.FindField('PSA_ETABLISSEMENT').AsString;
            If Q.FindField('PGA_CADRE').AsString='X' then
            begin
                 T.AddChampSupValeur('CSALCADRE','X');
                 T.AddChampSupValeur('CSALNONCADRE','-');
            end
            else
            begin
                 T.AddChampSupValeur('CSALCADRE','-');
                 T.AddChampSupValeur('CSALNONCADRE','X');
            end;
            T.AddChampSupValeur('EDATEDEB',Q.FindField('PGA_DATEDEBUT').AsDateTime);
            T.AddChampSupValeur('EDATEFIN',Q.FindField('PGA_DATEFIN').AsDateTime);

            MM := StrToInt(Q.FindField('PGA_MOIS').AsString);
            An := RechDom('PGANNEE',Q.FindField('PGA_ANNEE').AsString,False);
            If (Isnumeric(An)) and (length(An) = 4) then
            begin
                yy := StrToInt(An);
                DatePeriodeEC := EncodeDate(yy,mm,01);
                DatePeriodeEC := FinDeMois(DatePeriodeEC);
                If (Q.FindField('PGA_DATEFIN').AsDateTime > DatePeriodeEC) or (Q.FindField('PGA_DATEFIN').AsDateTime <= IDate1900) then
                begin
                     T.AddChampSupValeur('CENCOURS','X');
                end
                else T.AddChampSupValeur('CENCOURS','-');
            end
            else T.AddChampSupValeur('CENCOURS','-');
            T.AddChampSupValeur('EDATE',Q.FindField('PGA_DATEATTEST').AsDateTime);
            //PT7 Debut modif =====>
            AjoutV4 := Q.FindField('PGA_LIENPARENTE').AsString;
            Lparent := Trim(ReadTokenPipe(AjoutV4,';'));
            BrutSap1:= Trim(ReadTokenPipe(AjoutV4,';'));
            Taux1   := Trim(ReadTokenPipe(AjoutV4,';'));
            Cotisation1:= Trim(ReadTokenPipe(AjoutV4,';'));
            If BrutSap1 = '' Then BrutSap1 := '0';
            If Taux1 = '' Then Taux1 := '0';
            If Cotisation1 = '' Then Cotisation1 := '0';

            //if Q.FindField('PGA_LIENPARENTE').AsString <> '' then
            //begin
            //     T.AddChampSupValeur('LIENPARENTE',Q.FindField('PGA_LIENPARENTE').AsString);
            //     T.AddChampSupValeur('LELIENPARENTE',Q.FindField('PGA_LIENPARENTE').AsString);
            //     T.AddChampSupValeur('PARENTE','X');
            //end
            //else
            //begin
            //     T.AddChampSupValeur('LIENPARENTE','');
            //     T.AddChampSupValeur('LELIENPARENTE','');
            //     T.AddChampSupValeur('PARENTE','-');
            //end;
            If Lparent <> '' Then
            Begin
               T.AddChampSupValeur('LIENPARENTE',UpperCase(Lparent));
               T.AddChampSupValeur('PARENTE','X');
               T.AddChampSupValeur('PARENTENON','-');
            End
            Else
            Begin
               T.AddChampSupValeur('LIENPARENTE','');
               T.AddChampSupValeur('PARENTE','-');
               T.AddChampSupValeur('PARENTENON','X');
            End;
            //PT7 Fin modif <=====
            //PT9 Debut Ajout ==>
            T.AddChampSupValeur('CREALISATEUR','-');
            T.AddChampSupValeur('CARTISTE','-');
            T.AddChampSupValeur('CTECHNICIEN','-');
            T.AddChampSupValeur('COUVRIER','-');
            //PT9 Fin Ajout <==
            If (Categ = 'OUV') or (Categ = 'TEC') then
            begin
            //PT7    T.AddChampSupValeur('HEURESOUV',Q.FindField('PGA_NBHEURES').AsFloat);
                T.AddChampSupValeur('HEURESART',0);     //PT7
                T.AddChampSupValeur('CACHETS',0);       //PT7
                T.PutValue('CARTISTE','-');             //PT9 T.AddChampSupValeur
                T.PutValue('CREALISATEUR','-');         //PT9 T.AddChampSupValeur
                If Categ = 'TEC' then
                begin
                   T.PutValue('CTECHNICIEN','X');       //PT9 T.AddChampSupValeur
                   T.PutValue('COUVRIER','-');          //PT9 T.AddChampSupValeur
                end
                else
                begin
                   T.PutValue('COUVRIER','X');          //PT9 T.AddChampSupValeur
                   T.PutValue('CTECHNICIEN','-');       //PT9 T.AddChampSupValeur
                end;
            end
            else
            begin
            //PT7 Debut Modif ==>
                If (Categ = 'ART') or (Categ = 'REA') then
                begin
                   T.PutValue('CTECHNICIEN','-');    //PT9 T.AddChampSupValeur
                   T.PutValue('COUVRIER','-');       //PT9 T.AddChampSupValeur
                   If (Q.FindField('PGA_NBJOURS').AsFloat) < 5 Then
                   Begin
                      T.AddChampSupValeur('HEURESART',Q.FindField('PGA_NBCACHETS').AsFloat);
                      T.AddChampSupValeur('CACHETS',0);
                   End
                   Else
                   Begin
                      T.AddChampSupValeur('HEURESART',0);
                      T.AddChampSupValeur('CACHETS',Q.FindField('PGA_NBCACHETS').AsFloat);
                   End;
                   If Categ = 'ART' then
                   begin
                      T.PutValue('CARTISTE','X');    //PT9 T.AddChampSupValeur
                      T.PutValue('CREALISATEUR','-');//PT9 T.AddChampSupValeur
                   end
                   else
                   begin
                      T.PutValue('CREALISATEUR','X');//PT9 T.AddChampSupValeur
                      T.PutValue('CARTISTE','-');    //PT9 T.AddChampSupValeur
                   End;
                //PT9 End
                //PT9 Else
                //PT9 Begin
                //PT9    T.AddChampSupValeur('CARTISTE','-');
                //PT9    T.AddChampSupValeur('CREALISATEUR','-');
                //PT9    T.AddChampSupValeur('CTECHNICIEN','-');
                //PT9    T.AddChampSupValeur('COUVRIER','-');
                End;
            //    T.AddChampSupValeur('HEURESART',Q.FindField('PGA_NBHEURES').AsFloat);
            //    if Categ = 'ART' then T.AddChampSupValeur('CARTISTE','X')   //PT5
            //    else T.AddChampSupValeur('CARTISTE','-');
            //    T.AddChampSupValeur('CREALISATEUR','-');
            //    T.AddChampSupValeur('CTECHNICIEN','-');
            //    T.AddChampSupValeur('COUVRIER','-');
            end;
            //PT7 Fin Modif <==

            //PT7 T.AddChampSupValeur('CACHETS',Q.FindField('PGA_NBCACHETS').AsFloat);
            T.AddChampSupValeur('HEURESOUV',Q.FindField('PGA_NBHEURES').AsFloat);
            T.AddChampSupValeur('NBJOURSTRAV',Q.FindField('PGA_NBJOURS').AsFloat);
            T.AddChampSupValeur('ESALBTRUTSAV',Q.FindField('PGA_SALAIRESAVANT').AsFloat);
            T.AddChampSupValeur('ESALBTRUTSAP',Q.FindField('PGA_SALAIRESAPRES').AsFloat);
            T.AddChampSupValeur('ETAUX',Q.FindField('PGA_TAUX').AsFloat);
            T.AddChampSupValeur('ECOTISATION',Q.FindField('PGA_CONTRIBUTIONS').AsFloat);
            //PT7 Ajout ===>
            T.AddChampSupValeur('ESALBTRUTSAP1',StrToFloat(BrutSap1));
            T.AddChampSupValeur('ETAUX1',StrToFloat(Taux1));
            T.AddChampSupValeur('ECOTISATION1',StrToFloat(Cotisation1));
            T.AddChampSupValeur('ECOTISTOTAL', StrToFloat(Cotisation1) + Q.FindField('PGA_CONTRIBUTIONS').AsFloat);
            T.AddChampSupValeur('NUMOBJET',UpperCase(Q.FindField('PGA_NUMOBJETAEM').AsString));
            //PT7 Fin Ajout <===
        end;
        Ferme(Q);
        TEt := TobEtab.FindFirst(['ET_ETABLISSEMENT'],[Etab],False);

        T.AddChampSupValeur('ESOCIETE',UpperCase(TEt.getValue('ET_LIBELLE')));
        //PT7 T.AddChampSupValeur('EADRETAB1',TEt.getValue('ET_ADRESSE1'));
        //PT7 T.AddChampSupValeur('EADRETAB2',TEt.getValue('ET_ADRESSE2') + TEt.getValue('ET_ADRESSE3'));
        //PT7 T.AddChampSupValeur('JURIDIQUE',RechDom('TTFORMEJURIDIQUE',TEt.getValue('ET_JURIDIQUE'),False));  //PT9
        //PT7 T.AddChampSupValeur('ET_ACTIVITE',TEt.getValue('ET_ACTIVITE'));
        //PT7 T.AddChampSupValeur('ET_TELEPHONE',TEt.getValue('ET_TELEPHONE'));
        T.AddChampSupValeur('ETELETAB',TEt.getValue('ET_TELEPHONE'));  //PT7
        T.AddChampSupValeur('EFAXETAB',TEt.getValue('ET_FAX'));
        T.AddChampSupValeur('ESIRET',TEt.getValue('ET_SIRET'));
        T.AddChampSupValeur('EAPE',UpperCase(TEt.getValue('ET_APE')));
        //PT7 T.AddChampSupValeur('CODEPOSTAL',TEt.getValue('ET_CODEPOSTAL'));
        T.AddChampSupValeur('ECPETAB',TEt.getValue('ET_CODEPOSTAL'));  //PT7
        //PT7 T.AddChampSupValeur('ET_VILLE',TEt.getValue('ET_VILLE'));
        T.AddChampSupValeur('ECOMMUNE',UpperCase(TEt.getValue('ET_VILLE')));  //PT7
        T.AddChampSupValeur('ENUMCENTREREC',UpperCase(TEt.getValue('ETB_ISNUMRECOUV')));
        //PT7 T.AddChampSupValeur('EMAILETAB','');
        T.AddChampSupValeur('EMAILETAB',UpperCase(TEt.getValue('ET_EMAIL')));  //PT7

        If TEt.GetValue('ETB_ISLICSPEC')='X' Then
        begin
              T.AddChampSupValeur('CLICENCE','X');
              T.AddChampSupValeur('ENUMLICENCE',UpperCase(TEt.GetValue('ETB_ISNUMLIC')));
        end
        Else
        begin
              T.AddChampSupValeur('CLICENCENON','-');
              T.AddChampSupValeur('ENUMLICENCE','');
        end;
        If TEt.GetValue('ETB_ISLABELP')='X' Then
        begin
              T.AddChampSupValeur('CLABEL','X');
              T.AddChampSupValeur('CLABELNON','-');
              T.AddChampSupValeur('ENUMLABEL',UpperCase(TEt.GetValue('ETB_ISNUMLAB')));
        end
        Else
        begin
              T.AddChampSupValeur('CLABELNON','X');
              T.AddChampSupValeur('CLABEL','-');
              T.AddChampSupValeur('ENUMLABEL','');
        end;
        If TEt.GetValue('ETB_ISOCCAS')='X' Then
        begin
            T.AddChampSupValeur('CORGANISATEUR','X');
            T.AddChampSupValeur('CORGANISATEURNON','-');
        end
        Else
        begin
             T.AddChampSupValeur('CORGANISATEURNON','X');
             T.AddChampSupValeur('CORGANISATEUR','-');
        end;
        if (TEt.GetValue('ETB_ISNUMCPAY')) <> '' then
        begin
           T.AddChampSupValeur('ENUMAFFILIATION',UpperCase(TEt.GetValue('ETB_ISNUMCPAY')));
           T.AddChampSupValeur('CAFFILCP','X');
           T.AddChampSupValeur('CAFFILCPNON','-');
        end
        else
        begin
             T.AddChampSupValeur('CAFFILCPNON','X');
             T.AddChampSupValeur('CAFFILCP','-');
             T.AddChampSupValeur('ENUMAFFILIATION','');
        end;
        St := RechDom('PGDECLARANTATTEST', GetControlText('DECLARANT'), False);
        T.AddChampSupValeur('ENOMEMPL',UpperCase(RechDom('PGDECLARANTNOM', GetControlText('DECLARANT'), False)));
        T.AddChampSupValeur('EPRENOMEMPL',UpperCase(RechDom('PGDECLARANTPRENOM', GetControlText('DECLARANT'), False)));
        T.AddChampSupValeur('ELIEU',UpperCase(RechDom('PGDECLARANTVILLE', GetControlText('DECLARANT'), False)));
        T.AddChampSupValeur('PERSAJOINDRE',UpperCase(RechDom('PGDECLARANTATTEST', GetControlText('DECLARANT'), False)));
        T.AddChampSupValeur('TELPERSAJOINDRE', RechDom('PGDECLARANTTEL', GetControlText('DECLARANT'), False));

        St := RechDom('PGDECLARANTQUAL', GetControlText('DECLARANT'), False);
        if St = 'AUT' then T.AddChampSupValeur('EQUALITEEMPL',UpperCase(RechDom('PGDECLARANTAUTRE', GetControlText('DECLARANT'), False)))
        else T.AddChampSupValeur('EQUALITEEMPL',UpperCase(RechDom('PGQUALDECLARANT2', St, False)));
        ExecuteSQL('UPDATE PGAEM SET PGA_FAIT="X" WHERE PGA_SALARIE="'+Salarie+'" AND PGA_NUMATTEST="'+NumDoc+'"');
    end;
  end
  else if liste.AllSelected then
  begin
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
    InitMove(TFmul(Ecran).Q.RecordCount, '');
    Q_Mul.First;
    while not Q_Mul.EOF do
    begin
        T := Tob.Create('LigneEdition',TobEdition,-1);
        Salarie := TFmul(Ecran).Q.FindField('PGA_SALARIE').asstring;
        NumDoc := TFmul(Ecran).Q.FindField('PGA_NUMATTEST').asstring;
        Q := OpenSQL('SELECT * FROM PGAEM LEFT JOIN SALARIES ON PGA_SALARIE=PSA_SALARIE '+
        'LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+
        'WHERE PGA_SALARIE="'+Salarie+'" AND PGA_NUMATTEST="'+NumDoc+'"',True);

        //PT7 ajout du UpperCase lignes ci-dessous **
        T.AddChampSupValeur('PSA_SALARIE',Salarie);
        T.AddChampSupValeur('ENUMDOC',UpperCase(NumDoc));
        T.AddChampSupValeur('PGA_NUMATTEST',UpperCase(NumDoc));
        If not Q.Eof then
        begin
            Etab := Q.FindField('PSA_ETABLISSEMENT').AsString;
            Categ := Q.FindField('PGA_ISCATEG').AsString;
            CategIntermittent := Q.FindField('PSE_ISCATEG').AsString;
            //PT7 Debut Modif ===>
            If Q.FindField('PSA_NOMJF').AsString = '' then
            Begin
               T.AddChampSupValeur('LNOM',UpperCase(Q.FindField('PSA_LIBELLE').AsString));
               T.AddChampSupValeur('ENOMUSAGE',UpperCase(Q.FindField('PSA_SURNOM').AsString));
            End
            Else
            Begin
               T.AddChampSupValeur('LNOM',UpperCase(Q.FindField('PSA_NOMJF').AsString));
               T.AddChampSupValeur('ENOMUSAGE',UpperCase(Q.FindField('PSA_LIBELLE').AsString));
            End;
            //PT7 Fin Modif <===
            T.AddChampSupValeur('LPRENOM',UpperCase(Q.FindField('PSA_PRENOM').AsString));
            T.AddChampSupValeur('CATTESTINIT','-');
            T.AddChampSupValeur('ATTESTCOMPL','-');
            T.AddChampSupValeur('CRECAPPOS','-');
            T.AddChampSupValeur('RECAPNEG','-');
            iF Q.FindField('PGA_TYPEAEM').AsString = 'COM' then T.PutValue('ATTESTCOMPL','X')
            else iF Q.FindField('PGA_TYPEAEM').AsString = 'NEG' then T.PutValue('RECAPNEG','X')
            else iF Q.FindField('PGA_TYPEAEM').AsString = 'POS' then T.PutValue('CRECAPPOS','X')
            else T.PutValue('CATTESTINIT','X');
            T.AddChampSupValeur('EEMPLOI',UpperCase(Q.FindField('PGA_LIBELLEEMPLOI').AsString)); //PT4
            //T.AddChampSupValeur('ERETRAITE',Q.FindField('PSE_ISRETRAITE').AsString);           //PT7
            T.AddChampSupValeur('ERETRAITE',UpperCase(Q.FindField('PGA_ISRETRAITE').AsString));  //PT7
            T.AddChampSupValeur('ENUMASSEDIC',Q.FindField('PSE_ISNUMASSEDIC').AsString);
            //T.AddChampSupValeur('PSA_NUMEROSS',Q.FindField('PSA_NUMEROSS').AsString);  //PT7
            T.AddChampSupValeur('ENUMIDENTIFIANT',UpperCase(Q.FindField('PSA_NUMEROSS').AsString));//PT7
            T.AddChampSupValeur('EDATENAISSANCE',DateToStr(Q.FindField('PSA_DATENAISSANCE').AsDateTime));
            T.AddChampSupValeur('ETELSALARIE',Q.FindField('PSA_TELEPHONE').AsString);
            //T.AddChampSupValeur('PSA_CODEPOSTAL',Q.FindField('PSA_CODEPOSTAL').AsString);   //PT7
            T.AddChampSupValeur('ECPSAL',UpperCase(Q.FindField('PSA_CODEPOSTAL').AsString));  //PT7
            //T.AddChampSupValeur('PSA_VILLE',Q.FindField('PSA_VILLE').AsString);             //PT7
            T.AddChampSupValeur('ECOMMUNESAL',UpperCase(Q.FindField('PSA_VILLE').AsString));  //PT7
            T.AddChampSupValeur('EADRSAL1',UpperCase(Q.FindField('PSA_ADRESSE1').AsString));
            T.AddChampSupValeur('EADRSAL2',UpperCase(Q.FindField('PSA_ADRESSE2').AsString) + UpperCase(Q.FindField('PSA_ADRESSE3').AsString));
            //T.AddChampSupValeur('ENOMUSAGE',Q.FindField('PSA_SURNOM').AsString);          //PT7
            //T.AddChampSupValeur('EDATENAISSANCE',DateToStr(Q.FindField('PSA_DATENAISSANCE').AsDateTime));  //PT7
            T.AddChampSupValeur('MOIS',Q.FindField('PGA_MOIS').AsString);
            //DEBUT PT3
            T.AddChampSupValeur('CFINCDD', Q.FindField('PGA_FINCDD').AsString);
            T.AddChampSupValeur('CRUPTEMPL', Q.FindField('PGA_RUPTEMP').AsString);
            T.AddChampSupValeur('CRUPTSAL', Q.FindField('PGA_RUPTSAL').AsString);
            //FIN PT3
            T.AddChampSupValeur('ANNEE',RechDom('PGANNEE',Q.FindField('PGA_ANNEE').AsString,False));
            T.AddChampSupValeur('NUMATTESTINIT',UpperCase(Q.FindField('PGA_NUMATTESTINIT').AsString));
//            EtabSalarie := Q.FindField('PSA_ETABLISSEMENT').AsString;
            If Q.FindField('PGA_CADRE').AsString='X' then
            begin
                 T.AddChampSupValeur('CSALCADRE','X');
                 T.AddChampSupValeur('CSALNONCADRE','-');
            end
            else
            begin
                 T.AddChampSupValeur('CSALCADRE','-');
                 T.AddChampSupValeur('CSALNONCADRE','X');
            end;
            T.AddChampSupValeur('EDATEDEB',Q.FindField('PGA_DATEDEBUT').AsDateTime);
            T.AddChampSupValeur('EDATEFIN',Q.FindField('PGA_DATEFIN').AsDateTime);

            MM := StrToInt(Q.FindField('PGA_MOIS').AsString);
            An := RechDom('PGANNEE',Q.FindField('PGA_ANNEE').AsString,False);
            If (Isnumeric(An)) and (length(An) = 4) then
            begin
                yy := StrToInt(An);
                DatePeriodeEC := EncodeDate(yy,mm,01);
                DatePeriodeEC := FinDeMois(DatePeriodeEC);
                If (Q.FindField('PGA_DATEFIN').AsDateTime > DatePeriodeEC) or (Q.FindField('PGA_DATEFIN').AsDateTime <= IDate1900) then
                begin
                     T.AddChampSupValeur('CENCOURS','X');
                end
                else T.AddChampSupValeur('CENCOURS','-');
            end
            else T.AddChampSupValeur('CENCOURS','-');
            T.AddChampSupValeur('EDATE',Q.FindField('PGA_DATEATTEST').AsDateTime);
            //PT7 Debut modif =====>
            AjoutV4 := Q.FindField('PGA_LIENPARENTE').AsString;
            Lparent := Trim(ReadTokenPipe(AjoutV4,';'));
            BrutSap1:= Trim(ReadTokenPipe(AjoutV4,';'));
            Taux1   := Trim(ReadTokenPipe(AjoutV4,';'));
            Cotisation1:= Trim(ReadTokenPipe(AjoutV4,';'));
            If BrutSap1 = '' Then BrutSap1 := '0';
            If Taux1 = '' Then Taux1 := '0';
            If Cotisation1 = '' Then Cotisation1 := '0';

            //if Q.FindField('PGA_LIENPARENTE').AsString <> '' then
            //begin
            //     T.AddChampSupValeur('LIENPARENTE',Q.FindField('PGA_LIENPARENTE').AsString);
            //     T.AddChampSupValeur('LELIENPARENTE',Q.FindField('PGA_LIENPARENTE').AsString);
            //     T.AddChampSupValeur('PARENTE','X');
            //end
            //else
            //begin
            //     T.AddChampSupValeur('LIENPARENTE','');
            //     T.AddChampSupValeur('LELIENPARENTE','');
            //     T.AddChampSupValeur('PARENTE','-');
            //end;
            If Lparent <> '' Then
            Begin
               T.AddChampSupValeur('LIENPARENTE',UpperCase(Lparent));
               T.AddChampSupValeur('PARENTE','X');
               T.AddChampSupValeur('PARENTENON','-');
            End
            Else
            Begin
               T.AddChampSupValeur('LIENPARENTE','');
               T.AddChampSupValeur('PARENTE','-');
               T.AddChampSupValeur('PARENTENON','X');
            End;
            //PT7 Fin modif <=====
            //PT9 Debut Ajout ==>
            T.AddChampSupValeur('CREALISATEUR','-');
            T.AddChampSupValeur('CARTISTE','-');
            T.AddChampSupValeur('CTECHNICIEN','-');
            T.AddChampSupValeur('COUVRIER','-');
            //PT9 Fin Ajout <==
            If (Categ = 'OUV') or (Categ = 'TEC') then
            begin
            //PT7    T.AddChampSupValeur('HEURESOUV',Q.FindField('PGA_NBHEURES').AsFloat);
                T.AddChampSupValeur('HEURESART',0);     //PT7
                T.AddChampSupValeur('CACHETS',0);       //PT7
                T.PutValue('CARTISTE','-');             //PT9 T.AddChampSupValeur
                T.PutValue('CREALISATEUR','-');         //PT9 T.AddChampSupValeur
                If Categ = 'TEC' then
                begin
                   T.PutValue('CTECHNICIEN','X');       //PT9 T.AddChampSupValeur
                   T.PutValue('COUVRIER','-');          //PT9 T.AddChampSupValeur
                end
                else
                begin
                   T.PutValue('COUVRIER','X');          //PT9 T.AddChampSupValeur
                   T.PutValue('CTECHNICIEN','-');       //PT9 T.AddChampSupValeur
                end;
            end
            else
            begin
            //PT7 Debut Modif ==>
                If (Categ = 'ART') or (Categ = 'REA') then
                begin
                   T.PutValue('CTECHNICIEN','-');    //PT9 T.AddChampSupValeur
                   T.PutValue('COUVRIER','-');       //PT9 T.AddChampSupValeur
                   If (Q.FindField('PGA_NBJOURS').AsFloat) < 5 Then
                   Begin
                      T.AddChampSupValeur('HEURESART',Q.FindField('PGA_NBCACHETS').AsFloat);
                      T.AddChampSupValeur('CACHETS',0);
                   End
                   Else
                   Begin
                      T.AddChampSupValeur('HEURESART',0);
                      T.AddChampSupValeur('CACHETS',Q.FindField('PGA_NBCACHETS').AsFloat);
                   End;
                   If Categ = 'ART' then
                   begin
                      T.PutValue('CARTISTE','X');    //PT9 T.AddChampSupValeur
                      T.PutValue('CREALISATEUR','-');//PT9 T.AddChampSupValeur
                   end
                   else
                   begin
                      T.PutValue('CREALISATEUR','X');//PT9 T.AddChampSupValeur
                      T.PutValue('CARTISTE','-');    //PT9 T.AddChampSupValeur
                   End;
                //PT9 End
                //PT9 Else
                //PT9 Begin
                //PT9    T.AddChampSupValeur('CARTISTE','-');
                //PT9    T.AddChampSupValeur('CREALISATEUR','-');
                //PT9    T.AddChampSupValeur('CTECHNICIEN','-');
                //PT9    T.AddChampSupValeur('COUVRIER','-');
                End;
            //    T.AddChampSupValeur('HEURESART',Q.FindField('PGA_NBHEURES').AsFloat);
            //    if Categ = 'ART' then T.AddChampSupValeur('CARTISTE','X')   //PT5
            //    else T.AddChampSupValeur('CARTISTE','-');
            //    T.AddChampSupValeur('CREALISATEUR','-');
            //    T.AddChampSupValeur('CTECHNICIEN','-');
            //    T.AddChampSupValeur('COUVRIER','-');
            end;
            //PT7 Fin Modif <==

            //PT7 T.AddChampSupValeur('CACHETS',Q.FindField('PGA_NBCACHETS').AsFloat);
            T.AddChampSupValeur('HEURESOUV',Q.FindField('PGA_NBHEURES').AsFloat);
            T.AddChampSupValeur('NBJOURSTRAV',Q.FindField('PGA_NBJOURS').AsFloat);
            T.AddChampSupValeur('ESALBTRUTSAV',Q.FindField('PGA_SALAIRESAVANT').AsFloat);
            T.AddChampSupValeur('ESALBTRUTSAP',Q.FindField('PGA_SALAIRESAPRES').AsFloat);
            T.AddChampSupValeur('ETAUX',Q.FindField('PGA_TAUX').AsFloat);
            T.AddChampSupValeur('ECOTISATION',Q.FindField('PGA_CONTRIBUTIONS').AsFloat);
            //PT7 Ajout ===>
            T.AddChampSupValeur('ESALBTRUTSAP1',StrToFloat(BrutSap1));
            T.AddChampSupValeur('ETAUX1',StrToFloat(Taux1));
            T.AddChampSupValeur('ECOTISATION1',StrToFloat(Cotisation1));
            T.AddChampSupValeur('ECOTISTOTAL', StrToFloat(Cotisation1) + Q.FindField('PGA_CONTRIBUTIONS').AsFloat);
            T.AddChampSupValeur('NUMOBJET',UpperCase(Q.FindField('PGA_NUMOBJETAEM').AsString));
            //PT7 Fin Ajout <===
        end;
        Ferme(Q);
        TEt := TobEtab.FindFirst(['ET_ETABLISSEMENT'],[Etab],False);

        T.AddChampSupValeur('ESOCIETE',UpperCase(TEt.getValue('ET_LIBELLE')));
        //PT7 T.AddChampSupValeur('EADRETAB1',TEt.getValue('ET_ADRESSE1'));
        //PT7 T.AddChampSupValeur('EADRETAB2',TEt.getValue('ET_ADRESSE2') + TEt.getValue('ET_ADRESSE3'));
        //PT7 T.AddChampSupValeur('JURIDIQUE',RechDom('TTFORMEJURIDIQUE',TEt.getValue('ET_JURIDIQUE'),False));  //PT9
        //PT7 T.AddChampSupValeur('ET_ACTIVITE',TEt.getValue('ET_ACTIVITE'));
        //PT7 T.AddChampSupValeur('ET_TELEPHONE',TEt.getValue('ET_TELEPHONE'));
        T.AddChampSupValeur('ETELETAB',TEt.getValue('ET_TELEPHONE'));  //PT7
        T.AddChampSupValeur('EFAXETAB',TEt.getValue('ET_FAX'));
        T.AddChampSupValeur('ESIRET',TEt.getValue('ET_SIRET'));
        T.AddChampSupValeur('EAPE',UpperCase(TEt.getValue('ET_APE')));
        //PT7 T.AddChampSupValeur('CODEPOSTAL',TEt.getValue('ET_CODEPOSTAL'));
        T.AddChampSupValeur('ECPETAB',TEt.getValue('ET_CODEPOSTAL'));  //PT7
        //PT7 T.AddChampSupValeur('ET_VILLE',TEt.getValue('ET_VILLE'));
        T.AddChampSupValeur('ECOMMUNE',UpperCase(TEt.getValue('ET_VILLE')));  //PT7
        T.AddChampSupValeur('ENUMCENTREREC',UpperCase(TEt.getValue('ETB_ISNUMRECOUV')));
        //PT7 T.AddChampSupValeur('EMAILETAB','');
        T.AddChampSupValeur('EMAILETAB',UpperCase(TEt.getValue('ET_EMAIL')));  //PT7

        If TEt.GetValue('ETB_ISLICSPEC')='X' Then
        begin
              T.AddChampSupValeur('CLICENCE','X');
              T.AddChampSupValeur('ENUMLICENCE',UpperCase(TEt.GetValue('ETB_ISNUMLIC')));
        end
        Else
        begin
              T.AddChampSupValeur('CLICENCENON','-');
              T.AddChampSupValeur('ENUMLICENCE','');
        end;
        If TEt.GetValue('ETB_ISLABELP')='X' Then
        begin
              T.AddChampSupValeur('CLABEL','X');
              T.AddChampSupValeur('CLABELNON','-');
              T.AddChampSupValeur('ENUMLABEL',UpperCase(TEt.GetValue('ETB_ISNUMLAB')));
        end
        Else
        begin
              T.AddChampSupValeur('CLABELNON','X');
              T.AddChampSupValeur('CLABEL','-');
              T.AddChampSupValeur('ENUMLABEL','');
        end;
        If TEt.GetValue('ETB_ISOCCAS')='X' Then
        begin
            T.AddChampSupValeur('CORGANISATEUR','X');
            T.AddChampSupValeur('CORGANISATEURNON','-');
        end
        Else
        begin
             T.AddChampSupValeur('CORGANISATEURNON','X');
             T.AddChampSupValeur('CORGANISATEUR','-');
        end;
        if (TEt.GetValue('ETB_ISNUMCPAY')) <> '' then
        begin
           T.AddChampSupValeur('ENUMAFFILIATION',UpperCase(TEt.GetValue('ETB_ISNUMCPAY')));
           T.AddChampSupValeur('CAFFILCP','X');
           T.AddChampSupValeur('CAFFILCPNON','-');
        end
        else
        begin
             T.AddChampSupValeur('CAFFILCPNON','X');
             T.AddChampSupValeur('CAFFILCP','-');
             T.AddChampSupValeur('ENUMAFFILIATION','');
        end;
        St := RechDom('PGDECLARANTATTEST', GetControlText('DECLARANT'), False);
        T.AddChampSupValeur('ENOMEMPL',UpperCase(RechDom('PGDECLARANTNOM', GetControlText('DECLARANT'), False)));
        T.AddChampSupValeur('EPRENOMEMPL',UpperCase(RechDom('PGDECLARANTPRENOM', GetControlText('DECLARANT'), False)));
        T.AddChampSupValeur('ELIEU',UpperCase(RechDom('PGDECLARANTVILLE', GetControlText('DECLARANT'), False)));
        T.AddChampSupValeur('PERSAJOINDRE',UpperCase(RechDom('PGDECLARANTATTEST', GetControlText('DECLARANT'), False)));
        T.AddChampSupValeur('TELPERSAJOINDRE', RechDom('PGDECLARANTTEL', GetControlText('DECLARANT'), False));

        St := RechDom('PGDECLARANTQUAL', GetControlText('DECLARANT'), False);
        if St = 'AUT' then T.AddChampSupValeur('EQUALITEEMPL',UpperCase(RechDom('PGDECLARANTAUTRE', GetControlText('DECLARANT'), False)))
        else T.AddChampSupValeur('EQUALITEEMPL',UpperCase(RechDom('PGQUALDECLARANT2', St, False)));
        ExecuteSQL('UPDATE PGAEM SET PGA_FAIT="X" WHERE PGA_SALARIE="'+Salarie+'" AND PGA_NUMATTEST="'+NumDoc+'"');
      Q_Mul.Next;
    end;
  end;
  FiniMoveProgressForm;
  TobEdition.Detail.Sort('PGA_NUMATTEST');

  ImpSal := False;
  ImpAssedic := False;
  ImpDupli := False;
  If THMultiValComboBox(GetControl('EXEMPLAIRES')).Tous = True then
  begin
       ImpSal := True;
       ImpAssedic := True;
       ImpDupli := True;
  end
  else
  begin
       EtatsAImprimer := GetControlText('EXEMPLAIRES');
       While EtatsAImprimer <> '' do
       begin
            Etat := ReadTokenPipe(EtatsAImprimer,';');
            If Etat = 'SAL' then ImpSal := True
            else If Etat = 'ASS' then ImpAssedic := True
            else If Etat = 'EMP' then ImpDupli := True;
       end;
  end;
  If GetCheckBoxState('CALPHA') = CbChecked then TobEdition.Detail.Sort('LNOM')
  else TobEdition.Detail.Sort('ENUMDOC');
  If ImpSal then LanceEtatTOB('E','PAT','AS1',TobEdition,True,False,False,Pages,'','',False,0,'');
  If ImpAssedic then LanceEtatTOB('E','PAT','AS2',TobEdition,True,False,False,Pages,'','',False,0,'');
  If ImpDupli then LanceEtatTOB('E','PAT','AS2',TobEdition,True,False,False,Pages,'','',True,0,'');
  TobEdition.Free;
end;

//PT8 AJout de la procedure ===========================>
Procedure TOF_PGMULEDITASSEDICSPECTACLE.ExitEdit(Sender: TObject);
var edit : thedit;
Begin
edit:=THEdit(Sender);
If edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
End;

{TOF_PGMULSALSPECTACLE}
procedure TOF_PGMULSALSPECTACLE.OnLoad;
begin
Inherited ;
end;

procedure TOF_PGMULSALSPECTACLE.SuppAttestation(Sender : TObject);
var NumAttest : String;
     {$IFNDEF EAGLCLIENT}
  Liste: THDBGrid;
  {$ELSE}
  Liste: THGrid;
  {$ENDIF}
  Q_Mul: THQuery;
  i,Rep : Integer;
begin
 {$IFNDEF EAGLCLIENT}
  Liste := THDBGrid(GetControl('FLISTE'));
  {$ELSE}
  Liste := THGrid(GetControl('FLISTE'));
  {$ENDIF}
  if Liste = nil then Exit;
   Q_Mul := THQuery(Ecran.FindComponent('Q'));
  if Q_Mul = nil then Exit;

  if ((Liste.nbSelected) <= 0) and (not Liste.AllSelected) then
 begin
      PGIBox('Aucune attestation sélectionnée','Suppression des attestations');
      Exit;
 end;

 If Liste.AllSelected then
 begin
      Rep := PGIAsk('Attention vous allez supprimer l''ensemble des attestations du salarié voulez-vous continuer ');
      If Rep <> mrYes then Exit;
 end;
  if ((Liste.nbSelected) > 0) and (not Liste.AllSelected) then
  begin
    for i := 0 to Liste.NbSelected - 1 do
    begin
        Liste.GotoLeBOOKMARK(i);
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(Liste.Row - 1);
        {$ENDIF}
        NumAttest := TFmul(Ecran).Q.FindField('PGA_NUMATTEST').asstring;
        ExecuteSQL('DELETE FROM PGAEM WHERE PGA_NUMATTEST="'+NumAttest+'"');
        end;
  end
  else if liste.AllSelected then
  begin
    Q_Mul.First;
    while not Q_Mul.EOF do
    begin
          NumAttest := TFmul(Ecran).Q.FindField('PGA_NUMATTEST').asstring;
          ExecuteSQL('DELETE FROM PGAEM WHERE PGA_NUMATTEST="'+NumAttest+'"');
          Q_Mul.Next;
    end;
  end;
  TFMul(Ecran).BChercheClick(TFMUL(Ecran).BCherche);
end;

procedure TOF_PGMULSALSPECTACLE.OnArgument (stArgument : String ) ;
var DateAttestation : TDateTime;
    aa,mm,jj : Word;
    Q,QAttes : TQuery;
    EtabSalarie,Salarie : String;
    Bt : TToolBarButton97;

begin
Inherited ;
   Bt := TToolBarButton97(Getcontrol('BImprimer'));
   if Bt <> Nil then Bt.OnClick := EditionAttest;
   Salarie := ReadTokenPipe(stArgument,';');
   SetControlText('PGA_SALARIE',Salarie);
   EtabSalarie := '';
   Q := openSQL('SELECT PSA_ETABLISSEMENT FROM SALARIES WHERE PSA_SALARIE="'+Salarie+'"',True);
   If Not Q.Eof then EtabSalarie := Q.FindField('PSA_ETABLISSEMENT').AsString;
   Ferme(Q);
   DateAttestation:=PlusMois(V_PGI.DATEEntree,-1) ;
   DecodeDate(DateAttestation, aa, mm, jj);
   Q := OpenSQL('SELECT CO_CODE FROM COMMUN WHERE CO_TYPE="PGA" AND CO_LIBELLE="'+IntToStr(aa)+'"',True);
   If Not Q.Eof then SetControlText('PGA_ANNEE',Q.FindField('CO_CODE').AsString);
   Ferme(Q);
   SetControlText('PGA_MOIS','');
   Bt := TToolBarButton97(getControl('BDelete'));
   If Bt <> Nil then Bt.OnClick := SuppAttestation;
   TFMul(Ecran).Caption := 'Attestations du salarié : '+ Salarie + ' ' + RechDom('PGSALARIE',Salarie,False);
   SetControlCaption('LIBDECL','...');
   QAttes := OpenSql('SELECT PDA_DECLARANTATTES FROM DECLARANTATTEST ' +
      'WHERE (PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%' + EtabSalarie + '%") ' +
      'AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%ASP%" )  ' +
      'ORDER BY PDA_ETABLISSEMENT DESC', True);
   if not QAttes.eof then SetControlText('DECLARANT', QAttes.FindField('PDA_DECLARANTATTES').AsString);
   Ferme(QAttes);

end;

procedure TOF_PGMULSALSPECTACLE.EditionAttest(Sender : TOBject);
var Q : TQuery;
    TobEtab,TobEdition,TEt,T : Tob;
    Salarie,Etab,NumDoc ,Categ,St: String;
     {$IFNDEF EAGLCLIENT}
  Liste: THDBGrid;
  {$ELSE}
  Liste: THGrid;
  {$ENDIF}
  Q_Mul: THQuery;
  i : Integer;
  Pages : TPageControl;
  ImpSal,ImpAssedic,ImpDupli : Boolean;
  EtatsAImprimer,Etat : String;
  mm,yy : word;
  DatePeriodeEC : TDateTime;
  An,CategIntermittent : String;
  AjoutV4,Lparent,BrutSap1,Taux1,Cotisation1 : String; //PT7
begin
{$IFNDEF EAGLCLIENT}
  Liste := THDBGrid(GetControl('FLISTE'));
  {$ELSE}
  Liste := THGrid(GetControl('FLISTE'));
  {$ENDIF}
  if Liste = nil then Exit;
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
  if Q_Mul = nil then Exit;
  if ((Liste.nbSelected) <= 0) and (not Liste.AllSelected) then
  begin
      PGIBox('Aucune attestation sélectionnée','Edition des attestations');
      Exit;
  end;
  Pages := TPageControl(GetControl('PAGES'));
  Q := OpenSQL('SELECT * FROM ETABLISS LEFT JOIN ETABCOMPL ON ET_ETABLISSEMENT=ETB_ETABLISSEMENT',True);
  TobEtab := Tob.Create('LesEtab',Nil,-1);
  TobEtab.LoadDetailDB('lesEtabs','','',Q,False);
  Ferme(Q);
  TobEdition := Tob.Create('Edition',Nil,-1);
  if ((Liste.nbSelected) > 0) and (not Liste.AllSelected) then
  begin
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', Liste.nbSelected, FALSE, TRUE);
    InitMove(Liste.nbSelected, '');
    for i := 0 to Liste.NbSelected - 1 do
    begin
        Liste.GotoLeBOOKMARK(i);
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(Liste.Row - 1);
        {$ENDIF}
        T := Tob.Create('LigneEdition',TobEdition,-1);
        Salarie := TFmul(Ecran).Q.FindField('PGA_SALARIE').asstring;
        NumDoc := TFmul(Ecran).Q.FindField('PGA_NUMATTEST').asstring;
        Q := OpenSQL('SELECT * FROM PGAEM LEFT JOIN SALARIES ON PGA_SALARIE=PSA_SALARIE '+
        'LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+
        'WHERE PGA_SALARIE="'+Salarie+'" AND PGA_NUMATTEST="'+NumDoc+'"',True);

        //PT7 ajout du UpperCase lignes ci-dessous **
        T.AddChampSupValeur('PSA_SALARIE',Salarie);
        T.AddChampSupValeur('ENUMDOC',UpperCase(NumDoc));
        T.AddChampSupValeur('PGA_NUMATTEST',UpperCase(NumDoc));
        If not Q.Eof then
        begin
            Etab := Q.FindField('PSA_ETABLISSEMENT').AsString;
            Categ := Q.FindField('PGA_ISCATEG').AsString;
            CategIntermittent := Q.FindField('PSE_ISCATEG').AsString;
            //PT7 Debut Modif ===>
            If Q.FindField('PSA_NOMJF').AsString = '' then
            Begin
               T.AddChampSupValeur('LNOM',UpperCase(Q.FindField('PSA_LIBELLE').AsString));
               T.AddChampSupValeur('ENOMUSAGE',UpperCase(Q.FindField('PSA_SURNOM').AsString));
            End
            Else
            Begin
               T.AddChampSupValeur('LNOM',UpperCase(Q.FindField('PSA_NOMJF').AsString));
               T.AddChampSupValeur('ENOMUSAGE',UpperCase(Q.FindField('PSA_LIBELLE').AsString));
            End;
            //PT7 Fin Modif <===
            T.AddChampSupValeur('LPRENOM',UpperCase(Q.FindField('PSA_PRENOM').AsString));
            T.AddChampSupValeur('CATTESTINIT','-');
            T.AddChampSupValeur('ATTESTCOMPL','-');
            T.AddChampSupValeur('CRECAPPOS','-');
            T.AddChampSupValeur('RECAPNEG','-');
            iF Q.FindField('PGA_TYPEAEM').AsString = 'COM' then T.PutValue('ATTESTCOMPL','X')
            else iF Q.FindField('PGA_TYPEAEM').AsString = 'NEG' then T.PutValue('RECAPNEG','X')
            else iF Q.FindField('PGA_TYPEAEM').AsString = 'POS' then T.PutValue('CRECAPPOS','X')
            else T.PutValue('CATTESTINIT','X');
            T.AddChampSupValeur('EEMPLOI',UpperCase(Q.FindField('PGA_LIBELLEEMPLOI').AsString)); //PT4
            //T.AddChampSupValeur('ERETRAITE',Q.FindField('PSE_ISRETRAITE').AsString);           //PT7
            T.AddChampSupValeur('ERETRAITE',UpperCase(Q.FindField('PGA_ISRETRAITE').AsString));  //PT7
            T.AddChampSupValeur('ENUMASSEDIC',Q.FindField('PSE_ISNUMASSEDIC').AsString);
            //T.AddChampSupValeur('PSA_NUMEROSS',Q.FindField('PSA_NUMEROSS').AsString);  //PT7
            T.AddChampSupValeur('ENUMIDENTIFIANT',UpperCase(Q.FindField('PSA_NUMEROSS').AsString));//PT7
            T.AddChampSupValeur('EDATENAISSANCE',DateToStr(Q.FindField('PSA_DATENAISSANCE').AsDateTime));
            T.AddChampSupValeur('ETELSALARIE',Q.FindField('PSA_TELEPHONE').AsString);
            //T.AddChampSupValeur('PSA_CODEPOSTAL',Q.FindField('PSA_CODEPOSTAL').AsString);   //PT7
            T.AddChampSupValeur('ECPSAL',UpperCase(Q.FindField('PSA_CODEPOSTAL').AsString));  //PT7
            //T.AddChampSupValeur('PSA_VILLE',Q.FindField('PSA_VILLE').AsString);             //PT7
            T.AddChampSupValeur('ECOMMUNESAL',UpperCase(Q.FindField('PSA_VILLE').AsString));  //PT7
            T.AddChampSupValeur('EADRSAL1',UpperCase(Q.FindField('PSA_ADRESSE1').AsString));
            T.AddChampSupValeur('EADRSAL2',UpperCase(Q.FindField('PSA_ADRESSE2').AsString) + UpperCase(Q.FindField('PSA_ADRESSE3').AsString));
            //T.AddChampSupValeur('ENOMUSAGE',Q.FindField('PSA_SURNOM').AsString);          //PT7
            //T.AddChampSupValeur('EDATENAISSANCE',DateToStr(Q.FindField('PSA_DATENAISSANCE').AsDateTime));  //PT7
            T.AddChampSupValeur('MOIS',Q.FindField('PGA_MOIS').AsString);
            //DEBUT PT3
            T.AddChampSupValeur('CFINCDD', Q.FindField('PGA_FINCDD').AsString);
            T.AddChampSupValeur('CRUPTEMPL', Q.FindField('PGA_RUPTEMP').AsString);
            T.AddChampSupValeur('CRUPTSAL', Q.FindField('PGA_RUPTSAL').AsString);
            //FIN PT3
            T.AddChampSupValeur('ANNEE',RechDom('PGANNEE',Q.FindField('PGA_ANNEE').AsString,False));
            T.AddChampSupValeur('NUMATTESTINIT',UpperCase(Q.FindField('PGA_NUMATTESTINIT').AsString));
//            EtabSalarie := Q.FindField('PSA_ETABLISSEMENT').AsString;
            If Q.FindField('PGA_CADRE').AsString='X' then
            begin
                 T.AddChampSupValeur('CSALCADRE','X');
                 T.AddChampSupValeur('CSALNONCADRE','-');
            end
            else
            begin
                 T.AddChampSupValeur('CSALCADRE','-');
                 T.AddChampSupValeur('CSALNONCADRE','X');
            end;
            T.AddChampSupValeur('EDATEDEB',Q.FindField('PGA_DATEDEBUT').AsDateTime);
            T.AddChampSupValeur('EDATEFIN',Q.FindField('PGA_DATEFIN').AsDateTime);

            MM := StrToInt(Q.FindField('PGA_MOIS').AsString);
            An := RechDom('PGANNEE',Q.FindField('PGA_ANNEE').AsString,False);
            If (Isnumeric(An)) and (length(An) = 4) then
            begin
                yy := StrToInt(An);
                DatePeriodeEC := EncodeDate(yy,mm,01);
                DatePeriodeEC := FinDeMois(DatePeriodeEC);
                If (Q.FindField('PGA_DATEFIN').AsDateTime > DatePeriodeEC) or (Q.FindField('PGA_DATEFIN').AsDateTime <= IDate1900) then
                begin
                     T.AddChampSupValeur('CENCOURS','X');
                end
                else T.AddChampSupValeur('CENCOURS','-');
            end
            else T.AddChampSupValeur('CENCOURS','-');
            T.AddChampSupValeur('EDATE',Q.FindField('PGA_DATEATTEST').AsDateTime);
            //PT7 Debut modif =====>
            AjoutV4 := Q.FindField('PGA_LIENPARENTE').AsString;
            Lparent := Trim(ReadTokenPipe(AjoutV4,';'));
            BrutSap1:= Trim(ReadTokenPipe(AjoutV4,';'));
            Taux1   := Trim(ReadTokenPipe(AjoutV4,';'));
            Cotisation1:= Trim(ReadTokenPipe(AjoutV4,';'));
            If BrutSap1 = '' Then BrutSap1 := '0';
            If Taux1 = '' Then Taux1 := '0';
            If Cotisation1 = '' Then Cotisation1 := '0';

            //if Q.FindField('PGA_LIENPARENTE').AsString <> '' then
            //begin
            //     T.AddChampSupValeur('LIENPARENTE',Q.FindField('PGA_LIENPARENTE').AsString);
            //     T.AddChampSupValeur('LELIENPARENTE',Q.FindField('PGA_LIENPARENTE').AsString);
            //     T.AddChampSupValeur('PARENTE','X');
            //end
            //else
            //begin
            //     T.AddChampSupValeur('LIENPARENTE','');
            //     T.AddChampSupValeur('LELIENPARENTE','');
            //     T.AddChampSupValeur('PARENTE','-');
            //end;
            If Lparent <> '' Then
            Begin
               T.AddChampSupValeur('LIENPARENTE',UpperCase(Lparent));
               T.AddChampSupValeur('PARENTE','X');
               T.AddChampSupValeur('PARENTENON','-');
            End
            Else
            Begin
               T.AddChampSupValeur('LIENPARENTE','');
               T.AddChampSupValeur('PARENTE','-');
               T.AddChampSupValeur('PARENTENON','X');
            End;
            //PT7 Fin modif <=====
            //PT9 Debut Ajout ==>
            T.AddChampSupValeur('CREALISATEUR','-');
            T.AddChampSupValeur('CARTISTE','-');
            T.AddChampSupValeur('CTECHNICIEN','-');
            T.AddChampSupValeur('COUVRIER','-');
            //PT9 Fin Ajout <==
            If (Categ = 'OUV') or (Categ = 'TEC') then
            begin
            //PT7    T.AddChampSupValeur('HEURESOUV',Q.FindField('PGA_NBHEURES').AsFloat);
                T.AddChampSupValeur('HEURESART',0);     //PT7
                T.AddChampSupValeur('CACHETS',0);       //PT7
                T.PutValue('CARTISTE','-');             //PT9 T.AddChampSupValeur
                T.PutValue('CREALISATEUR','-');         //PT9 T.AddChampSupValeur
                If Categ = 'TEC' then
                begin
                   T.PutValue('CTECHNICIEN','X');       //PT9 T.AddChampSupValeur
                   T.PutValue('COUVRIER','-');          //PT9 T.AddChampSupValeur
                end
                else
                begin
                   T.PutValue('COUVRIER','X');          //PT9 T.AddChampSupValeur
                   T.PutValue('CTECHNICIEN','-');       //PT9 T.AddChampSupValeur
                end;
            end
            else
            begin
            //PT7 Debut Modif ==>
                If (Categ = 'ART') or (Categ = 'REA') then
                begin
                   T.PutValue('CTECHNICIEN','-');    //PT9 T.AddChampSupValeur
                   T.PutValue('COUVRIER','-');       //PT9 T.AddChampSupValeur
                   If (Q.FindField('PGA_NBJOURS').AsFloat) < 5 Then
                   Begin
                      T.AddChampSupValeur('HEURESART',Q.FindField('PGA_NBCACHETS').AsFloat);
                      T.AddChampSupValeur('CACHETS',0);
                   End
                   Else
                   Begin
                      T.AddChampSupValeur('HEURESART',0);
                      T.AddChampSupValeur('CACHETS',Q.FindField('PGA_NBCACHETS').AsFloat);
                   End;
                   If Categ = 'ART' then
                   begin
                      T.PutValue('CARTISTE','X');    //PT9 T.AddChampSupValeur
                      T.PutValue('CREALISATEUR','-');//PT9 T.AddChampSupValeur
                   end
                   else
                   begin
                      T.PutValue('CREALISATEUR','X');//PT9 T.AddChampSupValeur
                      T.PutValue('CARTISTE','-');    //PT9 T.AddChampSupValeur
                   End;
                //PT9 End
                //PT9 Else
                //PT9 Begin
                //PT9    T.AddChampSupValeur('CARTISTE','-');
                //PT9    T.AddChampSupValeur('CREALISATEUR','-');
                //PT9    T.AddChampSupValeur('CTECHNICIEN','-');
                //PT9    T.AddChampSupValeur('COUVRIER','-');
                End;
            //    T.AddChampSupValeur('HEURESART',Q.FindField('PGA_NBHEURES').AsFloat);
            //    if Categ = 'ART' then T.AddChampSupValeur('CARTISTE','X')   //PT5
            //    else T.AddChampSupValeur('CARTISTE','-');
            //    T.AddChampSupValeur('CREALISATEUR','-');
            //    T.AddChampSupValeur('CTECHNICIEN','-');
            //    T.AddChampSupValeur('COUVRIER','-');
            end;
            //PT7 Fin Modif <==

            //PT7 T.AddChampSupValeur('CACHETS',Q.FindField('PGA_NBCACHETS').AsFloat);
            T.AddChampSupValeur('HEURESOUV',Q.FindField('PGA_NBHEURES').AsFloat);
            T.AddChampSupValeur('NBJOURSTRAV',Q.FindField('PGA_NBJOURS').AsFloat);
            T.AddChampSupValeur('ESALBTRUTSAV',Q.FindField('PGA_SALAIRESAVANT').AsFloat);
            T.AddChampSupValeur('ESALBTRUTSAP',Q.FindField('PGA_SALAIRESAPRES').AsFloat);
            T.AddChampSupValeur('ETAUX',Q.FindField('PGA_TAUX').AsFloat);
            T.AddChampSupValeur('ECOTISATION',Q.FindField('PGA_CONTRIBUTIONS').AsFloat);
            //PT7 Ajout ===>
            T.AddChampSupValeur('ESALBTRUTSAP1',StrToFloat(BrutSap1));
            T.AddChampSupValeur('ETAUX1',StrToFloat(Taux1));
            T.AddChampSupValeur('ECOTISATION1',StrToFloat(Cotisation1));
            T.AddChampSupValeur('ECOTISTOTAL', StrToFloat(Cotisation1) + Q.FindField('PGA_CONTRIBUTIONS').AsFloat);
            T.AddChampSupValeur('NUMOBJET',UpperCase(Q.FindField('PGA_NUMOBJETAEM').AsString));
            //PT7 Fin Ajout <===
        end;
        Ferme(Q);
        TEt := TobEtab.FindFirst(['ET_ETABLISSEMENT'],[Etab],False);

        T.AddChampSupValeur('ESOCIETE',UpperCase(TEt.getValue('ET_LIBELLE')));
        //PT7 T.AddChampSupValeur('EADRETAB1',TEt.getValue('ET_ADRESSE1'));
        //PT7 T.AddChampSupValeur('EADRETAB2',TEt.getValue('ET_ADRESSE2') + TEt.getValue('ET_ADRESSE3'));
        //PT7 T.AddChampSupValeur('JURIDIQUE',RechDom('TTFORMEJURIDIQUE',TEt.getValue('ET_JURIDIQUE'),False));  //PT9
        //PT7 T.AddChampSupValeur('ET_ACTIVITE',TEt.getValue('ET_ACTIVITE'));
        //PT7 T.AddChampSupValeur('ET_TELEPHONE',TEt.getValue('ET_TELEPHONE'));
        T.AddChampSupValeur('ETELETAB',TEt.getValue('ET_TELEPHONE'));  //PT7
        T.AddChampSupValeur('EFAXETAB',TEt.getValue('ET_FAX'));
        T.AddChampSupValeur('ESIRET',TEt.getValue('ET_SIRET'));
        T.AddChampSupValeur('EAPE',UpperCase(TEt.getValue('ET_APE')));
        //PT7 T.AddChampSupValeur('CODEPOSTAL',TEt.getValue('ET_CODEPOSTAL'));
        T.AddChampSupValeur('ECPETAB',TEt.getValue('ET_CODEPOSTAL'));  //PT7
        //PT7 T.AddChampSupValeur('ET_VILLE',TEt.getValue('ET_VILLE'));
        T.AddChampSupValeur('ECOMMUNE',UpperCase(TEt.getValue('ET_VILLE')));  //PT7
        T.AddChampSupValeur('ENUMCENTREREC',UpperCase(TEt.getValue('ETB_ISNUMRECOUV')));
        //PT7 T.AddChampSupValeur('EMAILETAB','');
        T.AddChampSupValeur('EMAILETAB',UpperCase(TEt.getValue('ET_EMAIL')));  //PT7

        If TEt.GetValue('ETB_ISLICSPEC')='X' Then
        begin
              T.AddChampSupValeur('CLICENCE','X');
              T.AddChampSupValeur('ENUMLICENCE',UpperCase(TEt.GetValue('ETB_ISNUMLIC')));
        end
        Else
        begin
              T.AddChampSupValeur('CLICENCENON','-');
              T.AddChampSupValeur('ENUMLICENCE','');
        end;
        If TEt.GetValue('ETB_ISLABELP')='X' Then
        begin
              T.AddChampSupValeur('CLABEL','X');
              T.AddChampSupValeur('CLABELNON','-');
              T.AddChampSupValeur('ENUMLABEL',UpperCase(TEt.GetValue('ETB_ISNUMLAB')));
        end
        Else
        begin
              T.AddChampSupValeur('CLABELNON','X');
              T.AddChampSupValeur('CLABEL','-');
              T.AddChampSupValeur('ENUMLABEL','');
        end;
        If TEt.GetValue('ETB_ISOCCAS')='X' Then
        begin
            T.AddChampSupValeur('CORGANISATEUR','X');
            T.AddChampSupValeur('CORGANISATEURNON','-');
        end
        Else
        begin
             T.AddChampSupValeur('CORGANISATEURNON','X');
             T.AddChampSupValeur('CORGANISATEUR','-');
        end;
        if (TEt.GetValue('ETB_ISNUMCPAY')) <> '' then
        begin
           T.AddChampSupValeur('ENUMAFFILIATION',UpperCase(TEt.GetValue('ETB_ISNUMCPAY')));
           T.AddChampSupValeur('CAFFILCP','X');
           T.AddChampSupValeur('CAFFILCPNON','-');
        end
        else
        begin
             T.AddChampSupValeur('CAFFILCPNON','X');
             T.AddChampSupValeur('CAFFILCP','-');
             T.AddChampSupValeur('ENUMAFFILIATION','');
        end;
        St := RechDom('PGDECLARANTATTEST', GetControlText('DECLARANT'), False);
        T.AddChampSupValeur('ENOMEMPL',UpperCase(RechDom('PGDECLARANTNOM', GetControlText('DECLARANT'), False)));
        T.AddChampSupValeur('EPRENOMEMPL',UpperCase(RechDom('PGDECLARANTPRENOM', GetControlText('DECLARANT'), False)));
        T.AddChampSupValeur('ELIEU',UpperCase(RechDom('PGDECLARANTVILLE', GetControlText('DECLARANT'), False)));
        T.AddChampSupValeur('PERSAJOINDRE',UpperCase(RechDom('PGDECLARANTATTEST', GetControlText('DECLARANT'), False)));
        T.AddChampSupValeur('TELPERSAJOINDRE', RechDom('PGDECLARANTTEL', GetControlText('DECLARANT'), False));

        St := RechDom('PGDECLARANTQUAL', GetControlText('DECLARANT'), False);
        if St = 'AUT' then T.AddChampSupValeur('EQUALITEEMPL',UpperCase(RechDom('PGDECLARANTAUTRE', GetControlText('DECLARANT'), False)))
        else T.AddChampSupValeur('EQUALITEEMPL',UpperCase(RechDom('PGQUALDECLARANT2', St, False)));
        ExecuteSQL('UPDATE PGAEM SET PGA_FAIT="X" WHERE PGA_SALARIE="'+Salarie+'" AND PGA_NUMATTEST="'+NumDoc+'"');
    end;
  end
  else if liste.AllSelected then
  begin
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
    InitMove(TFmul(Ecran).Q.RecordCount, '');
    Q_Mul.First;
    while not Q_Mul.EOF do
    begin
        T := Tob.Create('LigneEdition',TobEdition,-1);
        Salarie := TFmul(Ecran).Q.FindField('PGA_SALARIE').asstring;
        NumDoc := TFmul(Ecran).Q.FindField('PGA_NUMATTEST').asstring;
        Q := OpenSQL('SELECT * FROM PGAEM LEFT JOIN SALARIES ON PGA_SALARIE=PSA_SALARIE '+
        'LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+
        'WHERE PGA_SALARIE="'+Salarie+'" AND PGA_NUMATTEST="'+NumDoc+'"',True);

        //PT7 ajout du UpperCase lignes ci-dessous **
        T.AddChampSupValeur('PSA_SALARIE',Salarie);
        T.AddChampSupValeur('ENUMDOC',UpperCase(NumDoc));
        T.AddChampSupValeur('PGA_NUMATTEST',UpperCase(NumDoc));
        If not Q.Eof then
        begin
            Etab := Q.FindField('PSA_ETABLISSEMENT').AsString;
            Categ := Q.FindField('PGA_ISCATEG').AsString;
            CategIntermittent := Q.FindField('PSE_ISCATEG').AsString;
            //PT7 Debut Modif ===>
            If Q.FindField('PSA_NOMJF').AsString = '' then
            Begin
               T.AddChampSupValeur('LNOM',UpperCase(Q.FindField('PSA_LIBELLE').AsString));
               T.AddChampSupValeur('ENOMUSAGE',UpperCase(Q.FindField('PSA_SURNOM').AsString));
            End
            Else
            Begin
               T.AddChampSupValeur('LNOM',UpperCase(Q.FindField('PSA_NOMJF').AsString));
               T.AddChampSupValeur('ENOMUSAGE',UpperCase(Q.FindField('PSA_LIBELLE').AsString));
            End;
            //PT7 Fin Modif <===
            T.AddChampSupValeur('LPRENOM',UpperCase(Q.FindField('PSA_PRENOM').AsString));
            T.AddChampSupValeur('CATTESTINIT','-');
            T.AddChampSupValeur('ATTESTCOMPL','-');
            T.AddChampSupValeur('CRECAPPOS','-');
            T.AddChampSupValeur('RECAPNEG','-');
            iF Q.FindField('PGA_TYPEAEM').AsString = 'COM' then T.PutValue('ATTESTCOMPL','X')
            else iF Q.FindField('PGA_TYPEAEM').AsString = 'NEG' then T.PutValue('RECAPNEG','X')
            else iF Q.FindField('PGA_TYPEAEM').AsString = 'POS' then T.PutValue('CRECAPPOS','X')
            else T.PutValue('CATTESTINIT','X');
            T.AddChampSupValeur('EEMPLOI',UpperCase(Q.FindField('PGA_LIBELLEEMPLOI').AsString)); //PT4
            //T.AddChampSupValeur('ERETRAITE',Q.FindField('PSE_ISRETRAITE').AsString);           //PT7
            T.AddChampSupValeur('ERETRAITE',UpperCase(Q.FindField('PGA_ISRETRAITE').AsString));  //PT7
            T.AddChampSupValeur('ENUMASSEDIC',Q.FindField('PSE_ISNUMASSEDIC').AsString);
            //T.AddChampSupValeur('PSA_NUMEROSS',Q.FindField('PSA_NUMEROSS').AsString);  //PT7
            T.AddChampSupValeur('ENUMIDENTIFIANT',UpperCase(Q.FindField('PSA_NUMEROSS').AsString));//PT7
            T.AddChampSupValeur('EDATENAISSANCE',DateToStr(Q.FindField('PSA_DATENAISSANCE').AsDateTime));
            T.AddChampSupValeur('ETELSALARIE',Q.FindField('PSA_TELEPHONE').AsString);
            //T.AddChampSupValeur('PSA_CODEPOSTAL',Q.FindField('PSA_CODEPOSTAL').AsString);   //PT7
            T.AddChampSupValeur('ECPSAL',UpperCase(Q.FindField('PSA_CODEPOSTAL').AsString));  //PT7
            //T.AddChampSupValeur('PSA_VILLE',Q.FindField('PSA_VILLE').AsString);             //PT7
            T.AddChampSupValeur('ECOMMUNESAL',UpperCase(Q.FindField('PSA_VILLE').AsString));  //PT7
            T.AddChampSupValeur('EADRSAL1',UpperCase(Q.FindField('PSA_ADRESSE1').AsString));
            T.AddChampSupValeur('EADRSAL2',UpperCase(Q.FindField('PSA_ADRESSE2').AsString) + UpperCase(Q.FindField('PSA_ADRESSE3').AsString));
            //T.AddChampSupValeur('ENOMUSAGE',Q.FindField('PSA_SURNOM').AsString);          //PT7
            //T.AddChampSupValeur('EDATENAISSANCE',DateToStr(Q.FindField('PSA_DATENAISSANCE').AsDateTime));  //PT7
            T.AddChampSupValeur('MOIS',Q.FindField('PGA_MOIS').AsString);
            //DEBUT PT3
            T.AddChampSupValeur('CFINCDD', Q.FindField('PGA_FINCDD').AsString);
            T.AddChampSupValeur('CRUPTEMPL', Q.FindField('PGA_RUPTEMP').AsString);
            T.AddChampSupValeur('CRUPTSAL', Q.FindField('PGA_RUPTSAL').AsString);
            //FIN PT3
            T.AddChampSupValeur('ANNEE',RechDom('PGANNEE',Q.FindField('PGA_ANNEE').AsString,False));
            T.AddChampSupValeur('NUMATTESTINIT',UpperCase(Q.FindField('PGA_NUMATTESTINIT').AsString));
//            EtabSalarie := Q.FindField('PSA_ETABLISSEMENT').AsString;
            If Q.FindField('PGA_CADRE').AsString='X' then
            begin
                 T.AddChampSupValeur('CSALCADRE','X');
                 T.AddChampSupValeur('CSALNONCADRE','-');
            end
            else
            begin
                 T.AddChampSupValeur('CSALCADRE','-');
                 T.AddChampSupValeur('CSALNONCADRE','X');
            end;
            T.AddChampSupValeur('EDATEDEB',Q.FindField('PGA_DATEDEBUT').AsDateTime);
            T.AddChampSupValeur('EDATEFIN',Q.FindField('PGA_DATEFIN').AsDateTime);

            MM := StrToInt(Q.FindField('PGA_MOIS').AsString);
            An := RechDom('PGANNEE',Q.FindField('PGA_ANNEE').AsString,False);
            If (Isnumeric(An)) and (length(An) = 4) then
            begin
                yy := StrToInt(An);
                DatePeriodeEC := EncodeDate(yy,mm,01);
                DatePeriodeEC := FinDeMois(DatePeriodeEC);
                If (Q.FindField('PGA_DATEFIN').AsDateTime > DatePeriodeEC) or (Q.FindField('PGA_DATEFIN').AsDateTime <= IDate1900) then
                begin
                     T.AddChampSupValeur('CENCOURS','X');
                end
                else T.AddChampSupValeur('CENCOURS','-');
            end
            else T.AddChampSupValeur('CENCOURS','-');
            T.AddChampSupValeur('EDATE',Q.FindField('PGA_DATEATTEST').AsDateTime);
            //PT7 Debut modif =====>
            AjoutV4 := Q.FindField('PGA_LIENPARENTE').AsString;
            Lparent := Trim(ReadTokenPipe(AjoutV4,';'));
            BrutSap1:= Trim(ReadTokenPipe(AjoutV4,';'));
            Taux1   := Trim(ReadTokenPipe(AjoutV4,';'));
            Cotisation1:= Trim(ReadTokenPipe(AjoutV4,';'));
            If BrutSap1 = '' Then BrutSap1 := '0';
            If Taux1 = '' Then Taux1 := '0';
            If Cotisation1 = '' Then Cotisation1 := '0';

            //if Q.FindField('PGA_LIENPARENTE').AsString <> '' then
            //begin
            //     T.AddChampSupValeur('LIENPARENTE',Q.FindField('PGA_LIENPARENTE').AsString);
            //     T.AddChampSupValeur('LELIENPARENTE',Q.FindField('PGA_LIENPARENTE').AsString);
            //     T.AddChampSupValeur('PARENTE','X');
            //end
            //else
            //begin
            //     T.AddChampSupValeur('LIENPARENTE','');
            //     T.AddChampSupValeur('LELIENPARENTE','');
            //     T.AddChampSupValeur('PARENTE','-');
            //end;
            If Lparent <> '' Then
            Begin
               T.AddChampSupValeur('LIENPARENTE',UpperCase(Lparent));
               T.AddChampSupValeur('PARENTE','X');
               T.AddChampSupValeur('PARENTENON','-');
            End
            Else
            Begin
               T.AddChampSupValeur('LIENPARENTE','');
               T.AddChampSupValeur('PARENTE','-');
               T.AddChampSupValeur('PARENTENON','X');
            End;
            //PT7 Fin modif <=====
            //PT9 Debut Ajout ==>
            T.AddChampSupValeur('CREALISATEUR','-');
            T.AddChampSupValeur('CARTISTE','-');
            T.AddChampSupValeur('CTECHNICIEN','-');
            T.AddChampSupValeur('COUVRIER','-');
            //PT9 Fin Ajout <==
            If (Categ = 'OUV') or (Categ = 'TEC') then
            begin
            //PT7    T.AddChampSupValeur('HEURESOUV',Q.FindField('PGA_NBHEURES').AsFloat);
                T.AddChampSupValeur('HEURESART',0);     //PT7
                T.AddChampSupValeur('CACHETS',0);       //PT7
                T.PutValue('CARTISTE','-');             //PT9 T.AddChampSupValeur
                T.PutValue('CREALISATEUR','-');         //PT9 T.AddChampSupValeur
                If Categ = 'TEC' then
                begin
                   T.PutValue('CTECHNICIEN','X');       //PT9 T.AddChampSupValeur
                   T.PutValue('COUVRIER','-');          //PT9 T.AddChampSupValeur
                end
                else
                begin
                   T.PutValue('COUVRIER','X');          //PT9 T.AddChampSupValeur
                   T.PutValue('CTECHNICIEN','-');       //PT9 T.AddChampSupValeur
                end;
            end
            else
            begin
            //PT7 Debut Modif ==>
                If (Categ = 'ART') or (Categ = 'REA') then
                begin
                   T.PutValue('CTECHNICIEN','-');    //PT9 T.AddChampSupValeur
                   T.PutValue('COUVRIER','-');       //PT9 T.AddChampSupValeur
                   If (Q.FindField('PGA_NBJOURS').AsFloat) < 5 Then
                   Begin
                      T.AddChampSupValeur('HEURESART',Q.FindField('PGA_NBCACHETS').AsFloat);
                      T.AddChampSupValeur('CACHETS',0);
                   End
                   Else
                   Begin
                      T.AddChampSupValeur('HEURESART',0);
                      T.AddChampSupValeur('CACHETS',Q.FindField('PGA_NBCACHETS').AsFloat);
                   End;
                   If Categ = 'ART' then
                   begin
                      T.PutValue('CARTISTE','X');    //PT9 T.AddChampSupValeur
                      T.PutValue('CREALISATEUR','-');//PT9 T.AddChampSupValeur
                   end
                   else
                   begin
                      T.PutValue('CREALISATEUR','X');//PT9 T.AddChampSupValeur
                      T.PutValue('CARTISTE','-');    //PT9 T.AddChampSupValeur
                   End;
                //PT9 End
                //PT9 Else
                //PT9 Begin
                //PT9    T.AddChampSupValeur('CARTISTE','-');
                //PT9    T.AddChampSupValeur('CREALISATEUR','-');
                //PT9    T.AddChampSupValeur('CTECHNICIEN','-');
                //PT9    T.AddChampSupValeur('COUVRIER','-');
                End;
            //    T.AddChampSupValeur('HEURESART',Q.FindField('PGA_NBHEURES').AsFloat);
            //    if Categ = 'ART' then T.AddChampSupValeur('CARTISTE','X')   //PT5
            //    else T.AddChampSupValeur('CARTISTE','-');
            //    T.AddChampSupValeur('CREALISATEUR','-');
            //    T.AddChampSupValeur('CTECHNICIEN','-');
            //    T.AddChampSupValeur('COUVRIER','-');
            end;
            //PT7 Fin Modif <==

            //PT7 T.AddChampSupValeur('CACHETS',Q.FindField('PGA_NBCACHETS').AsFloat);
            T.AddChampSupValeur('HEURESOUV',Q.FindField('PGA_NBHEURES').AsFloat);
            T.AddChampSupValeur('NBJOURSTRAV',Q.FindField('PGA_NBJOURS').AsFloat);
            T.AddChampSupValeur('ESALBTRUTSAV',Q.FindField('PGA_SALAIRESAVANT').AsFloat);
            T.AddChampSupValeur('ESALBTRUTSAP',Q.FindField('PGA_SALAIRESAPRES').AsFloat);
            T.AddChampSupValeur('ETAUX',Q.FindField('PGA_TAUX').AsFloat);
            T.AddChampSupValeur('ECOTISATION',Q.FindField('PGA_CONTRIBUTIONS').AsFloat);
            //PT7 Ajout ===>
            T.AddChampSupValeur('ESALBTRUTSAP1',StrToFloat(BrutSap1));
            T.AddChampSupValeur('ETAUX1',StrToFloat(Taux1));
            T.AddChampSupValeur('ECOTISATION1',StrToFloat(Cotisation1));
            T.AddChampSupValeur('ECOTISTOTAL', StrToFloat(Cotisation1) + Q.FindField('PGA_CONTRIBUTIONS').AsFloat);
            T.AddChampSupValeur('NUMOBJET',UpperCase(Q.FindField('PGA_NUMOBJETAEM').AsString));
            //PT7 Fin Ajout <===
        end;
        Ferme(Q);
        TEt := TobEtab.FindFirst(['ET_ETABLISSEMENT'],[Etab],False);

        T.AddChampSupValeur('ESOCIETE',UpperCase(TEt.getValue('ET_LIBELLE')));
        //PT7 T.AddChampSupValeur('EADRETAB1',TEt.getValue('ET_ADRESSE1'));
        //PT7 T.AddChampSupValeur('EADRETAB2',TEt.getValue('ET_ADRESSE2') + TEt.getValue('ET_ADRESSE3'));
        //PT7 T.AddChampSupValeur('JURIDIQUE',RechDom('TTFORMEJURIDIQUE',TEt.getValue('ET_JURIDIQUE'),False));  //PT9
        //PT7 T.AddChampSupValeur('ET_ACTIVITE',TEt.getValue('ET_ACTIVITE'));
        //PT7 T.AddChampSupValeur('ET_TELEPHONE',TEt.getValue('ET_TELEPHONE'));
        T.AddChampSupValeur('ETELETAB',TEt.getValue('ET_TELEPHONE'));  //PT7
        T.AddChampSupValeur('EFAXETAB',TEt.getValue('ET_FAX'));
        T.AddChampSupValeur('ESIRET',TEt.getValue('ET_SIRET'));
        T.AddChampSupValeur('EAPE',UpperCase(TEt.getValue('ET_APE')));
        //PT7 T.AddChampSupValeur('CODEPOSTAL',TEt.getValue('ET_CODEPOSTAL'));
        T.AddChampSupValeur('ECPETAB',TEt.getValue('ET_CODEPOSTAL'));  //PT7
        //PT7 T.AddChampSupValeur('ET_VILLE',TEt.getValue('ET_VILLE'));
        T.AddChampSupValeur('ECOMMUNE',UpperCase(TEt.getValue('ET_VILLE')));  //PT7
        T.AddChampSupValeur('ENUMCENTREREC',UpperCase(TEt.getValue('ETB_ISNUMRECOUV')));
        //PT7 T.AddChampSupValeur('EMAILETAB','');
        T.AddChampSupValeur('EMAILETAB',UpperCase(TEt.getValue('ET_EMAIL')));  //PT7

        If TEt.GetValue('ETB_ISLICSPEC')='X' Then
        begin
              T.AddChampSupValeur('CLICENCE','X');
              T.AddChampSupValeur('ENUMLICENCE',UpperCase(TEt.GetValue('ETB_ISNUMLIC')));
        end
        Else
        begin
              T.AddChampSupValeur('CLICENCENON','-');
              T.AddChampSupValeur('ENUMLICENCE','');
        end;
        If TEt.GetValue('ETB_ISLABELP')='X' Then
        begin
              T.AddChampSupValeur('CLABEL','X');
              T.AddChampSupValeur('CLABELNON','-');
              T.AddChampSupValeur('ENUMLABEL',UpperCase(TEt.GetValue('ETB_ISNUMLAB')));
        end
        Else
        begin
              T.AddChampSupValeur('CLABELNON','X');
              T.AddChampSupValeur('CLABEL','-');
              T.AddChampSupValeur('ENUMLABEL','');
        end;
        If TEt.GetValue('ETB_ISOCCAS')='X' Then
        begin
            T.AddChampSupValeur('CORGANISATEUR','X');
            T.AddChampSupValeur('CORGANISATEURNON','-');
        end
        Else
        begin
             T.AddChampSupValeur('CORGANISATEURNON','X');
             T.AddChampSupValeur('CORGANISATEUR','-');
        end;
        if (TEt.GetValue('ETB_ISNUMCPAY')) <> '' then
        begin
           T.AddChampSupValeur('ENUMAFFILIATION',UpperCase(TEt.GetValue('ETB_ISNUMCPAY')));
           T.AddChampSupValeur('CAFFILCP','X');
           T.AddChampSupValeur('CAFFILCPNON','-');
        end
        else
        begin
             T.AddChampSupValeur('CAFFILCPNON','X');
             T.AddChampSupValeur('CAFFILCP','-');
             T.AddChampSupValeur('ENUMAFFILIATION','');
        end;
        St := RechDom('PGDECLARANTATTEST', GetControlText('DECLARANT'), False);
        T.AddChampSupValeur('ENOMEMPL',UpperCase(RechDom('PGDECLARANTNOM', GetControlText('DECLARANT'), False)));
        T.AddChampSupValeur('EPRENOMEMPL',UpperCase(RechDom('PGDECLARANTPRENOM', GetControlText('DECLARANT'), False)));
        T.AddChampSupValeur('ELIEU',UpperCase(RechDom('PGDECLARANTVILLE', GetControlText('DECLARANT'), False)));
        T.AddChampSupValeur('PERSAJOINDRE',UpperCase(RechDom('PGDECLARANTATTEST', GetControlText('DECLARANT'), False)));
        T.AddChampSupValeur('TELPERSAJOINDRE', RechDom('PGDECLARANTTEL', GetControlText('DECLARANT'), False));

        St := RechDom('PGDECLARANTQUAL', GetControlText('DECLARANT'), False);
        if St = 'AUT' then T.AddChampSupValeur('EQUALITEEMPL',UpperCase(RechDom('PGDECLARANTAUTRE', GetControlText('DECLARANT'), False)))
        else T.AddChampSupValeur('EQUALITEEMPL',UpperCase(RechDom('PGQUALDECLARANT2', St, False)));
        ExecuteSQL('UPDATE PGAEM SET PGA_FAIT="X" WHERE PGA_SALARIE="'+Salarie+'" AND PGA_NUMATTEST="'+NumDoc+'"');
      Q_Mul.Next;
    end;
  end;
  FiniMoveProgressForm;
  TobEdition.Detail.Sort('PGA_NUMATTEST');

  ImpSal := False;
  ImpAssedic := False;
  ImpDupli := False;
  If THMultiValComboBox(GetControl('EXEMPLAIRES')).Tous = True then
  begin
       ImpSal := True;
       ImpAssedic := True;
       ImpDupli := True;
  end
  else
  begin
       EtatsAImprimer := GetControlText('EXEMPLAIRES');
       While EtatsAImprimer <> '' do
       begin
            Etat := ReadTokenPipe(EtatsAImprimer,';');
            If Etat = 'SAL' then ImpSal := True
            else If Etat = 'ASS' then ImpAssedic := True
            else If Etat = 'EMP' then ImpDupli := True;
       end;
  end;
  TobEdition.Detail.Sort('ENUMDOC');
  If ImpSal then LanceEtatTOB('E','PAT','AS1',TobEdition,True,False,False,Pages,'','',False,0,'');
  If ImpAssedic then LanceEtatTOB('E','PAT','AS2',TobEdition,True,False,False,Pages,'','',False,0,'');
  If ImpDupli then LanceEtatTOB('E','PAT','AS2',TobEdition,True,False,False,Pages,'','',True,0,'');
  TobEdition.Free;
end;

Initialization
  registerclasses ( [ TOF_PGMULASSEDIC_SPECTACLE,TOF_PGMULEDITASSEDICSPECTACLE,TOF_PGMULSALSPECTACLE] ) ;
end.
