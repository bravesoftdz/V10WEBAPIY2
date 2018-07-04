{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 23/09/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULCONGESSPEC ()
Mots clefs ... : TOF;PGMULCONGESSPEC
*****************************************************************}
{
PT1 26/01/2007  V_80 FC Mise en place du filtage habilitation pour les lookuplist
                        pour les critères code salarié uniquement
PT2 03/12/2007 V_72  RM FQ 14020 : Compléter le matricule salarié par des zeros automatiquement
}
Unit UTofPgMulCongesSpec;

Interface

Uses Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Mul,Fe_Main,HDB,
{$ELSE}
     eMul,MainEAGL,
{$ENDIF}
     sysutils,HCtrls,HEnt1,UTOF,UTob,HTB97,HStatus,ed_tools,ParamDat,HQry,LookUp,pgOutils,
     EntPaie,PGOutils2 ;  //PT2

Type
  TOF_PGMULCONGESSPEC = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    {$IFNDEF EAGLCLIENT}
    Liste : THDBGrid;
    {$ELSE}
    Liste : THGrid;
    {$ENDIF}
    procedure CalculPeriode(Sender : TObject);
    procedure GenerationDuFichier(Sender : TObject);
    procedure DateElipsisClick(Sender : TObject);
    procedure GrilleDblClick(Sender : TObject);
    procedure CreerPeriode(Sender : TObject);
    procedure SalarieElipsisClick(Sender : TObject);
    Procedure ExitEdit(Sender: TObject);  //PT2
  end ;

Implementation

procedure TOF_PGMULCONGESSPEC.OnLoad ;
begin
        If GetControlText('ETAB') <> '' then
        begin
                SetControlText('XX_WHERE','PCS_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES '+
                'WHERE PSA_ETABLISSEMENT="'+Getcontroltext('ETAB')+'")');
        end
        Else SetControltext('XX_WHERE','');
end;

procedure TOF_PGMULCONGESSPEC.OnArgument (S : String ) ;
var DateDebut,DateFin : TDateTime;
    BCalcul,BFichier : TToolBarButton97;
    THDate,Edit : THEdit;
    aa,mm,jj,Mois : word;
    Arg : String;

begin
  Inherited ;
        Arg := ReadTokenPipe(S,';');
        If Arg = 'CALCUL' then
        begin
                SetControlVisible('BINSERT',False);
                SetControlVisible('BOUVRIR',False);
                BCalcul := TToolBarButton97(GetControl('BCALCULER'));
                If BCalcul <> Nil Then BCalcul.OnClick := CalculPeriode;
        end;
        If Arg = 'SAISIE' then
        begin
                SetControlVisible('BCALCULER',False);
                {$IFNDEF EAGLCLIENT}
                Liste := THDBGrid(GetControl('FListe'));
                {$ELSE}
                Liste := THGrid(GetControl('FListe'));
                {$ENDIF}
                If Liste <> Nil then Liste.OnDblClick := GrilleDblClick;
                TFMul(Ecran).BInsert.OnClick := Creerperiode;
                Ecran.Caption := 'Saisie consultation congés spectacles';
        end;
        If Arg = 'FICHIER' then
        begin
                SetControlVisible('BINSERT',False);
                BFichier := TToolBarButton97(GetControl('BOUVRIR'));
                If BFichier <> Nil Then BFichier.OnClick := GenerationDuFichier;
                BFichier.Hint := 'Création du fichier';
                SetControlVisible('BCALCULER',False);
                Ecran.Caption := 'Génération pré-fichier Caisse Congés Spectacles';
        end;
        UpdateCaption(TFMul(Ecran));
        THDate := THEdit(GetControl('PCS_DATEDEBUT'));
        If THDate <> Nil Then THDate.OnElipsisClick := DateElipsisClick;
        THDate := THEdit(GetControl('PCS_DATEFIN'));
        If THDate <> Nil Then THDate.OnElipsisClick := DateElipsisClick;
        SetControlCaption('LIBSAL','');
        SetControltext('ETAB','');
        DecodeDate(V_PGI.DateEntree,aa,mm,jj);
        Mois:=1;
        if arg='FICHIER' then
        begin
                If mm<4 Then
                begin
                        aa:=aa-1;
                        Mois:=10;
                end;
                If (mm<7) and (mm>3) Then Mois:=1;
                If (mm<10) and (mm>6) Then Mois:=4;
                If mm>9 Then Mois:=7;
        end
        else
        begin
                If (mm<7) and (mm>3) Then Mois:=4;
                If (mm<10) and (mm>6) Then Mois:=7;
                If mm>9 Then Mois:=10;
        end;
        DateDebut:=EncodeDate(aa,Mois,1);
        DateFin:=PlusMois(DateDebut,2);
        DateFin:=FinDeMois(DateFin);
        SetControlText('PCS_DATEDEBUT',DateToStr(Datedebut));
        SetControltext('PCS_DATEFIN',dateToStr(DateFin));
        Edit := THEdit(GetControl('PCS_SALARIE'));
        If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
        If Edit <> Nil then Edit.OnExit := ExitEdit;    //PT2
end ;

procedure TOF_PGMULCONGESSPEC.CalculPeriode(Sender : TObject);
var TobCongesSpec,TobPaie,TC,TVerif,TobCumulSS,TobCumulBrut,TCS,TCB : Tob;
    Q : TQuery;
    NumSS,CleSS,NumCS,CleCS,NomNaiss,NomEpoux,Prenom,Pseudo : String;
    Emploi,JJd,MMd,AAd,Datedebut,DateFin,JJf,MMf,AAf : String;
    PayeLe,aa,mm,jj : String;
    Longueur,i : Integer;
    verifSal,StAnd : String;
    VerifDateDebut,VerifDateFin : TDateTime;
    BaseConge,Salaire : Double;
begin
        StAnd := '';
        If GetControlText('ETAB') <> '' then StAnd := ' AND PSA_ETABLISSEMENT="'+Getcontroltext('ETAB')+'"';
        If GetControltext('PCS_SALARIE') <> '' then StAnd := StAnd + ' AND PSA_SALARIE="'+GetcontrolText('PCS_SALARIE')+'"';
        Q := OpenSQL('SELECT PPU_SALARIE,PPU_DATEDEBUT,PPU_DATEFIN,PPU_PAYELE,PPU_CBRUT,PPU_CBASESS,PSA_NUMEROSS,PSE_ISCONGSPEC'+
        ',PSA_PRENOM,PSA_LIBELLE,PSA_NOMJF,PSA_SURNOM,PPU_LIBELLEEMPLOI,PSA_DADSCAT'+
        ' FROM PAIEENCOURS'+
        ' LEFT JOIN SALARIES ON PPU_SALARIE=PSA_SALARIE'+
        ' LEFT JOIN DEPORTSAL ON PPU_SALARIE=PSE_SALARIE'+
        ' WHERE PSE_INTERMITTENT="X" AND '+
        ' PPU_DATEDEBUT>="'+UsDatetime(StrToDate(GetControlText('PCS_DATEDEBUT')))+'"'+
        ' AND PPU_DATEFIN<="'+UsDatetime(StrToDate(GetControlText('PCS_DATEFIN')))+'"'+StAnd,True);
        TobPaie := Tob.Create('Les paies',Nil,-1);
        TobPaie.LoadDetailDB('PAIEENCOURS','','',Q,False);
        Ferme(Q);
        Q := OpenSQL('SELECT SUM (PHC_MONTANT) BRUT FROM HISTOCUMSAL LEFT JOIN DEPORTSAL ON PHC_SALARIE=PSE_SALARIE '+
        'WHERE PHC_DATEDEBUT>="'+UsDatetime(StrToDate(GetControlText('PCS_DATEDEBUT')))+'"'+
        ' AND PHC_DATEFIN<="'+UsDatetime(StrToDate(GetControlText('PCS_DATEFIN')))+'"'+
        ' AND PHC_CUMULPAIE="01"'+
        ' GROUP BY PHC_DATEDEBUT,PHC_DATEFIN,PHC_SALARIE',True);
        TobCumulBrut := Tob.Create('Les cumuls',Nil,-1);
        TobCumulBrut.LoadDetailDB('Les cumuls','','',Q,False);
        Ferme(Q);
        Q := OpenSQL('SELECT SUM (PHC_MONTANT) BASESS FROM HISTOCUMSAL LEFT JOIN DEPORTSAL ON PHC_SALARIE=PSE_SALARIE '+
        'WHERE PHC_DATEDEBUT>="'+UsDatetime(StrToDate(GetControlText('PCS_DATEDEBUT')))+'"'+
        ' AND PHC_DATEFIN<="'+UsDatetime(StrToDate(GetControlText('PCS_DATEFIN')))+'"'+
        ' AND PHC_CUMULPAIE="11"'+
        ' GROUP BY PHC_DATEDEBUT,PHC_DATEFIN,PHC_SALARIE',True);
        TobCumulSS := Tob.Create('Les cumuls',Nil,-1);
        TobCumulSS.LoadDetailDB('Les cumuls','','',Q,False);
        Ferme(Q);
        Q := OpenSQL('SELECT * FROM CONGESSPEC WHERE'+
        ' PCS_DATEDEBUT>="'+UsDatetime(StrToDate(GetControlText('PCS_DATEDEBUT')))+'"'+
        ' AND PCS_DATEFIN<="'+UsDatetime(StrToDate(GetControlText('PCS_DATEFIN')))+'"',True);
        TobCongesSpec := Tob.create('CONGESSPEC',Nil,-1);
        TobCongesSpec.LoadDetailDB('CONGESSPEC','','',Q,False);
        Ferme(Q);
        InitMoveProgressForm (NIL,'Début du traitement', 'Veuillez patienter SVP ...',TobPaie.Detail.Count-1,FALSE,TRUE);
        InitMove(TobPaie.Detail.Count,'');
        for i := 0 to TobPaie.Detail.Count - 1 do
        begin
                TC := Tob.Create('CONGESSPEC',TobCongesSpec,-1);
                TC.PutValue('PCS_SALARIE',TobPaie.Detail[i].GetValue('PPU_SALARIE'));
                VerifSal := TobPaie.Detail[i].GetValue('PPU_SALARIE');
                TC.PutValue('PCS_DATEDEBUT',TobPaie.Detail[i].GetValue('PPU_DATEDEBUT'));
                VerifDateDebut := TobPaie.Detail[i].GetValue('PPU_DATEDEBUT');
                TC.PutValue('PCS_DATEFIN',TobPaie.Detail[i].GetValue('PPU_DATEFIN'));
                VerifDateFin := TobPaie.Detail[i].GetValue('PPU_DATEFIN');
                TVerif := TobCongesSpec.FindFirst(['PCS_SALARIE','PCS_DATEDEBUT','PCS_DATEFIN'],[VerifSal,VerifDateDebut,VerifDateFin],False);
                If TVerif <> Nil Then TVerif.DeleteDB(False);
                NumSS := TobPaie.Detail[i].GetValue('PSA_NUMEROSS');
                Longueur := Length(NumSS);
                If Longueur>13 Then CleSS := Copy(NumSS,14,2)
                Else CleSS := '';
                If NumSS <> '' Then SetLength(NumSS,13);
                TC.PutValue('PCS_NUMEROSS',NumSS);
                TC.PutValue('PCS_CLENUMSS',CleSS);
                NumCS := TobPaie.Detail[i].GetValue('PSE_ISCONGSPEC');
                Longueur := Length(NumCS);
                If Longueur>1 Then CleCS := NumCS[1]
                Else CleCS := '';
                TC.PutValue('PCS_CLENUMCONG',CleCS);
                If Longueur>1 Then NumCS := Copy(NumCS,2,Longueur-1)
                Else NumCS := '';
                TC.PutValue('PCS_NUMCONGES',NumCS);
                If TobPaie.Detail[i].GetValue('PSA_NOMJF') <> '' Then
                begin
                        NomNaiss := TobPaie.Detail[i].GetValue('PSA_NOMJF');
                        NomEpoux := TobPaie.Detail[i].GetValue('PSA_LIBELLE');
                        If Length(NomNaiss)>20 Then SetLength(NomNaiss,20);
                        If Length(NomEpoux)>20 Then SetLength(NomEpoux,20);
                end
                Else
                begin
                        NomNaiss := TobPaie.Detail[i].GetValue('PSA_LIBELLE');
                        NomEpoux := '';
                        If Length(NomNaiss)>20 Then SetLength(NomNaiss,20);
                end;
                TC.PutValue('PCS_NOMNAISS',NomNaiss);
                TC.PutValue('PCS_NOMEPOUX',NomEpoux);
                Prenom := TobPaie.Detail[i].GetValue('PSA_PRENOM');
                If Length(Prenom)>12 Then SetLength(Prenom,12);
                TC.PutValue('PCS_PRENOM',Prenom);
                Pseudo := TobPaie.Detail[i].GetValue('PSA_SURNOM');
                If Length(Pseudo)>20 Then SetLength(Pseudo,20);
                TC.PutValue('PCS_PSEUDO',Pseudo);
                Emploi := TobPaie.Detail[i].GetValue('PPU_LIBELLEEMPLOI');
//                Emploi := RechDom('PGLIBEMPLOI',Emploi,False);
                If Length(Emploi)>15 Then SetLength(Emploi,15);
                TC.PutValue('PCS_EMPLOI',Emploi);
                If (TobPaie.Detail[i].GetValue('PSA_DADSCAT') = '01') or (TobPaie.Detail[i].GetValue('PSA_DADSCAT') = '01') then
                TC.PutValue('PCS_CADRE','O')
                else TC.PutValue('PCS_CADRE','N');
                DateDebut := DateToStr(TobPaie.Detail[i].GetValue('PPU_DATEDEBUT'));
                DateFin := DateToStr(TobPaie.Detail[i].GetValue('PPU_DATEFIN'));
                JJd := ReadTokenPipe(DateDebut,'/');
                MMd := ReadTokenPipe(DateDebut,'/');
                AAd := ReadTokenPipe(DateDebut,'/');
                JJf := ReadTokenPipe(DateFin,'/');
                MMf := ReadTokenPipe(DateFin,'/');
                AAf := ReadTokenPipe(DateFin,'/');
                TC.PutValue('PCS_DATESTRAVAIL',JJd+MMd+AAd+JJf+MMf+AAf);
                Q := OpenSQL('SELECT PCI_ISNBEFFECTUE FROM CONTRATTRAVAIL '+
                'WHERE PCI_SALARIE="'+TobPaie.Detail[i].GetValue('PPU_SALARIE')+'" '+
                'AND PCI_DEBUTCONTRAT="'+UsDateTime(TobPaie.Detail[i].GetValue('PPU_DATEDEBUT'))+'" '+
                'AND PCI_FINCONTRAT="'+UsDateTime(TobPaie.Detail[i].GetValue('PPU_DATEFIN'))+'" '+
                'AND PCI_ISCACHETS="X"',True);
                If Not Q.Eof then TC.PutValue('PCS_NBJOURS',FloatToStr(Arrondi(Q.FindField('PCI_ISNBEFFECTUE').AsFloat,0)))
                Else TC.PutValue('PCS_NBJOURS',FloatToStr(Arrondi(VerifDateFin - VerifDateDebut,0)));
                Ferme(Q);
                TCB := TobCumulBrut.FindFirst(['PHC_SALARIE','PHC_DATEDEBUT','PHC_DATEFIN'],[VerifSal,VerifDateDebut,VerifDateFin],False);
                if TCB <> Nil then Salaire := Arrondi(TCB.GetValue('BRUT'),0)
                else Salaire := Arrondi(TobPaie.Detail[i].GetValue('PPU_CBRUT'),0);
                TCS := TobCumulSS.FindFirst(['PHC_SALARIE','PHC_DATEDEBUT','PHC_DATEFIN'],[VerifSal,VerifDateDebut,VerifDateFin],False);
                if TCS <> Nil then BaseConge := Arrondi(TCS.GetValue('BASESS'),0)
                else BaseConge := Arrondi(TobPaie.Detail[i].GetValue('PPU_CBASESS'),0);
                If BaseConge < 0 then BaseConge := 0;
                If Salaire < 0 then Salaire := 0;
                TC.PutValue('PCS_BASECONGE',FloatToStr(BaseConge));
                TC.PutValue('PCS_SALBRUT',FloatToStr(Salaire));
                PayeLe := DateToStr(TobPaie.Detail[i].GetValue('PPU_PAYELE'));
                jj := ReadTokenPipe(PayeLe,'/');
                mm := ReadTokenPipe(PayeLe,'/');
                aa := ReadTokenPipe(PayeLe,'/');
                TC.PutValue('PCS_DATEPAIEM',jj+mm+aa);
                TC.PutValue('PCS_LIEU','');
                TC.PutValue('PCS_DATEDELIV','');
                TC.PutValue('PCS_NOMSIGN','');
                TC.PutValue('PCS_NUMCERTIFICAT','');
                TC.PutValue('PCS_NUMENTREP','');
                TC.PutValue('PCS_LETTRECLE','');
                TC.PutValue('PCS_LIBRE','');
                TC.InsertorUpdateDB(False);
                MoveCurProgressForm(TobPaie.Detail[i].GetValue('PSA_LIBELLE'));
        end;
        FiniMoveProgressForm;
        TobPaie.Free;
        TobCongesSpec.Free;
        TobCumulBrut.Free;
        TobCumulSS.Free;
        TFMul(Ecran).BCherche.Click;
end;

procedure TOF_PGMULCONGESSPEC.GenerationDuFichier(Sender:TObject);
var Q:TQuery;
    Etab,NumEntre,St : String;

begin
        Etab := GetControlText('ETAB');
        NumEntre := '';
        If Etab <> '' then
        begin
                Q := OpenSQL('SELECT ETB_ISNUMCPAY FROM ETABCOMPL WHERE ETB_ETABLISSEMENT="'+Etab+'"',True);
                If Not Q.Eof Then NumEntre := Q.FindField('ETB_ISNUMCPAY').AsString;
                Ferme(Q);
        end;
        St := GetControltext('PCS_DATEDEBUT')+';'+GetControltext('PCS_DATEFIN')+';'+NumEntre;
        AGLLanceFiche('PAY','CONGESSPECENTREP','','',St);
end;

procedure TOF_PGMULCONGESSPEC.DateElipsisClick(Sender : TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;


procedure TOF_PGMULCONGESSPEC.GrilleDblClick(Sender : TObject);
var St : String;
    Q_Mul : THQuery;
begin
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(Liste.Row-1) ;
        {$ENDIF}
        Q_Mul := THQuery(Ecran.FindComponent('Q'));
        St := Q_Mul.FindField('PCS_SALARIE').AsString+';'+
        DateToStr(Q_Mul.FindField('PCS_DATEDEBUT').AsDateTime)+';'+DateToStr(Q_Mul.FindField('PCS_DATEFIN').AsDatetime);
        AGLLanceFiche('PAY', 'CONGESSPEC', '',St, 'ACTION=MODIFICATION');
        TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

procedure TOF_PGMULCONGESSPEC.CreerPeriode(Sender : TObject);
begin
        AGLLanceFiche('PAY', 'CONGESSPEC', '','', 'ACTION=CREATION');
        TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

procedure TOF_PGMULCONGESSPEC.SalarieElipsisClick(Sender : TObject);
var St,StWhere,StOrder : String;
begin
        St := 'SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM FROM SALARIES LEFT JOIN DEPORTSAL ON PSE_SALARIE=PSA_SALARIE';
        StWhere := ' WHERE PSE_INTERMITTENT="X"';
        StOrder := 'PSA_SALARIE';
        StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT1
        {$IFNDEF EAGLCLIENT}
        LookupList(THEdit(Sender),'Liste des stages','STAGE','PSA_SALARIE','PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM','',StOrder, True,-1,St+StWhere);
        {$ELSE}
// PT1        LookupList(THEdit(Sender),'Liste des stages','STAGE','PSA_SALARIE','PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM',StWhere,StOrder, True,-1,St);
        LookupList(THEdit(Sender),'Liste des stages','STAGE','PSA_SALARIE','PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM','',StOrder, True,-1,St+StWhere);
        {$ENDIF}
end;

//PT2 AJout de la procedure ===========================>
Procedure TOF_PGMULCONGESSPEC.ExitEdit(Sender: TObject);
var edit : thedit;
Begin
edit:=THEdit(Sender);
If edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
End;

Initialization
  registerclasses ( [ TOF_PGMULCONGESSPEC ] ) ;
end.
