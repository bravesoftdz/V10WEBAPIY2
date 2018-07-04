{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 23/09/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : CONGESSPEC (CONGESSPEC)
Mots clefs ... : TOM;CONGESSPEC
*****************************************************************}
{
PT1 26/01/2007  V_80 FC Mise en place du filtage habilitation pour les lookuplist
                        pour les critères code salarié uniquement
}
Unit UTomCongesSpectacles;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,FE_Main,
{$ELSE}
     EFiche,EFichList,UtileAgl,MainEAGL,Utob,
{$ENDIF}
     sysutils,HCtrls,HEnt1,HMsgBox,UTOM,LookUp,HTB97,PGOutils;

Type
  TOM_CONGESSPEC = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    private
    InitCadre:Boolean;
    InitBaseConges,InitSalaire,InitDatePaiem,InitDateDeliv : String;
    procedure SalarieElipsisClick(Sender : TObject);
    procedure CadreClick(Sender : TObject);
    procedure ExitDatePaie(Sender : TObject);
    procedure ExitDateDelivrance(Sender : TObject);
    procedure ImprimerCertifEmploi(Sender : TObject);
    procedure ExitTHEdit(Sender : TObject);
    end ;

Implementation

procedure TOM_CONGESSPEC.OnNewRecord ;
var Jour,Mois,Annee : String;
    aa,mm,jj : Word;
begin
        SetField('PCS_SALBRUT','0');
        SetField('PCS_BASECONGE','0');
        DecodeDate(V_PGI.DateEntree,aa,mm,jj);
        Annee := IntToStr(aa);
        If mm<10 Then Mois := '0'+IntToStr(mm)
        Else Mois := IntToStr(mm);
        If jj<10 Then jour := '0'+IntToStr(jj)
        Else jour := IntToStr(jj);
        SetField('PCS_DATEDELIV',Jour+Mois+Annee);
        SetField('PCS_DATEPAIEM',Jour+Mois+Annee);
        SetField('PCS_DATEDEBUT',V_PGI.DateEntree);
        SetField('PCS_DATEFIN',V_PGI.DateEntree);
end;

procedure TOM_CONGESSPEC.OnUpdateRecord ;
var DatesTravail,DateDeliv,DatePaiem,Jour,Mois,Annee : String;
    aa,mm,jj : Word;
    SalaireBrut,BaseConges : Double;
begin
  Inherited ;
        If not Isnumeric(GetField('PCS_NBJOURS')) then
        begin
                PGIBox('Le nombre de jours ou cachets doit être numérique',Ecran.Caption);
                LastError := 1;
                Exit;
        end;
        DecodeDate(GetField('PCS_DATEDEBUT'),aa,mm,jj);
        Annee := IntToStr(aa);
        If mm<10 Then Mois := '0'+IntToStr(mm)
        Else Mois := IntToStr(mm);
        If jj<10 Then jour := '0'+IntToStr(jj)
        Else jour := IntToStr(jj);
        DatesTravail := Jour+Mois+Annee;
        DecodeDate(GetField('PCS_DATEFIN'),aa,mm,jj);
        Annee := IntToStr(aa);
        If mm<10 Then Mois := '0'+IntToStr(mm)
        Else Mois := IntToStr(mm);
        If jj<10 Then jour := '0'+IntToStr(jj)
        Else jour := IntToStr(jj);
        DatesTravail := DatesTravail+Jour+Mois+Annee;
        SetField('PCS_DATESTRAVAIL',DatesTravail);

        If Not isValidDate(GetControlText('DATEDELIV')) then SetField('PCS_DATEDELIV','')
        Else
        begin
                DecodeDate(StrToDate(GetControlText('DATEDELIV')),aa,mm,jj);
                Annee := IntToStr(aa);
                If mm<10 Then Mois := '0'+IntToStr(mm)
                Else Mois := IntToStr(mm);
                If jj<10 Then jour := '0'+IntToStr(jj)
                Else jour := IntToStr(jj);
                DateDeliv := Jour+Mois+Annee;
                SetField('PCS_DATEDELIV',DateDeliv);
        end;

        If Not isValidDate(GetControlText('DATEPAIE')) then SetField('PCS_DATEPAIEM','')
        Else
        begin
                DecodeDate(StrTodate(GetControlText('DATEPAIE')),aa,mm,jj);
                Annee := IntToStr(aa);
                If mm<10 Then Mois := '0'+IntToStr(mm)
                Else Mois := IntToStr(mm);
                If jj<10 Then jour := '0'+IntToStr(jj)
                Else jour := IntToStr(jj);
                DatePaiem := DatePaiem+Jour+Mois+Annee;
                SetField('PCS_DATEPAIEM',DatePaiem);
        end;

        If GetControlText('SALBRUT') = '' then SetControlText('SALBRUT','0')
        else
        begin
                SalaireBrut := StrToFloat(GetControlText('SALBRUT'));
                SetField('PCS_SALBRUT',Arrondi(SalaireBrut,0));
        end;
        If GetControlText('BASECONGE') = '' then SetControlText('BASECONGE','0')
        else
        begin
                BaseConges := StrToFloat(GetControlText('BASECONGE'));
                SetField('PCS_BASECONGE',Arrondi(BaseConges,0));
        end;
end ;

procedure TOM_CONGESSPEC.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_CONGESSPEC.OnLoadRecord ;
var jj,mm,aa : String;
begin
  Inherited ;
        If GetField('PCS_CADRE') = 'O' Then
        begin
                InitCadre := True;
                SetcontrolChecked('CBCADRE',True);
        end
        Else
        begin
                SetcontrolChecked('CBCADRE',False);
                InitCadre := False;
        end;
        If GetField('PCS_DATEPAIEM') <> '' then
        begin
                jj := Copy(GetField('PCS_DATEPAIEM'),1,2);
                mm := Copy(GetField('PCS_DATEPAIEM'),3,2);
                aa := Copy(GetField('PCS_DATEPAIEM'),5,4);
                SetControlText('DATEPAIE',jj+'/'+mm+'/'+aa);
        end
        Else SetControltext('DATEPAIE',DateToStr(IDate1900));
        If GetField('PCS_DATEDELIV') <> '' then
        begin
                jj := Copy(GetField('PCS_DATEDELIV'),1,2);
                mm := Copy(GetField('PCS_DATEDELIV'),3,2);
                aa := Copy(GetField('PCS_DATEDELIV'),5,4);
                SetControlText('DATEDELIV',jj+'/'+mm+'/'+aa);
        end
        Else SetControltext('DATEDELIV',DateToStr(IDate1900));
        SetControlText('SALBRUT',GetField('PCS_SALBRUT'));
        SetControlText('BASECONGE',getField('PCS_BASECONGE'));
        InitBaseConges := GetControlText('BASECONGE');
        InitSalaire := GetControlText('SALBRUT');
        InitDatePaiem := GetControlText('DATEPAIE');
        InitDateDeliv := GetControlText('DATEDELIV');
end ;

procedure TOM_CONGESSPEC.OnChangeField ( F : TField ) ;
var NumSS,NumCS,NumEtab,Nom,NomJF,Prenom,LibelleEmploi : String;
    Q : TQuery;
begin
  Inherited ;
        if DS.State in [dsInsert] then
        begin
                If F.FieldName = 'PCS_SALARIE' Then
                begin
                        If GetField('PCS_SALARIE') <> '' Then
                        begin
                                Q := OpenSQL('SELECT PSA_NOMJF,PSA_LIBELLE,PSA_PRENOM,PSA_NUMEROSS,PSA_LIBELLEEMPLOI,PSE_ISCONGSPEC,PSA_DADSCAT,'+
                                'ETB_ISNUMCPAY '+
                                'FROM SALARIES '+
                                'LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT '+
                                'LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+
                                'WHERE PSA_SALARIE="'+GetField('PCS_SALARIE')+'"',True);
                                NumSS := Q.findField('PSA_NUMEROSS').AsString;
                                If NumSS <> '' Then
                                begin
                                        SetField('PCS_NUMEROSS',Copy(NumSS,1,13));
                                        SetField('PCS_CLENUMSS',Copy(NumSS,14,2));
                                end;
                                NumCS := Q.findField('PSE_ISCONGSPEC').AsString;
                                If NumCS <> '' Then
                                begin
                                        SetField('PCS_NUMCONGES',Copy(NumCS,2,7));
                                        SetField('PCS_CLENUMCONG',Copy(NumCS,1,1));
                                end;
                                NumEtab := Q.findField('ETB_ISNUMCPAY').AsString;
                                If NumEtab <> '' Then
                                begin
                                        SetField('PCS_NUMENTREP',Copy(NumEtab,1,8));
                                        SetField('PCS_LETTRECLE',Copy(NumEtab,9,1));
                                end;
                                NomJF := Q.findField('PSA_NOMJF').AsString;
                                Nom := Q.findField('PSA_LIBELLE').AsString;
                                If NomJF <> '' Then
                                begin
                                        SetField('PCS_NOMNAISS',NomJF);
                                        SetField('PCS_NOMEPOUX',Nom);
                                end
                                Else
                                begin
                                        SetField('PCS_NOMNAISS',Nom);
                                        SetField('PCS_NOMEPOUX','');
                                end;
                                Prenom := Q.findField('PSA_PRENOM').AsString;
                                SetField('PCS_PRENOM',Prenom);
                                LibelleEmploi := Q.findField('PSA_LIBELLEEMPLOI').AsString;
                                SetField('PCS_EMPLOI',LibelleEmploi);
                                If (Q.FindField('PSA_DADSCAT').AsString = '01') or (Q.FindField('PSA_DADSCAT').AsString = '02') then
                                begin
                                        SetControlChecked('CBCADRE',True);
                                        SetField('PCS_CADRE','O');
                                end
                                else
                                begin
                                        SetControlChecked('CBCADRE',False);
                                        SetField('PCS_CADRE','N');
                                end;
                                Ferme(Q);
                        end;
                end;
        end;
end ;

procedure TOM_CONGESSPEC.OnArgument ( S : String ) ;
var
{$IFNDEF EAGLCLIENT}
Salarie : THDBEdit;
{$ELSE}
Salarie : THEdit;
{$ENDIF}
Cadre : TCheckBox;
Edit,Datepaie,Dateliv : THEdit;
BImprim : TToolBarButton97;
begin
  Inherited ;
        {$IFNDEF EAGLCLIENT}
        Salarie := THDBEdit(GetControl('PCS_SALARIE'));
        {$ELSE}
        Salarie := THEdit(GetControl('PCS_SALARIE'));
        {$ENDIF}
        If Salarie <> Nil Then Salarie.OnElipsisClick := SalarieElipsisClick;
        Cadre := TCheckBox(GetControl('CBCADRE'));
        If Cadre <> Nil Then cadre.OnClick := CadreClick;
        Datepaie := THEdit(GetControl('DATEPAIE'));
        If DatePaie <> Nil Then DatePaie.OnExit := ExitDatePaie;
        DateLiv := THEdit(GetControl('DATEDELIV'));
        If DateLiv <> Nil Then DateLiv.OnExit := ExitDateDelivrance;
        BImprim := TToolBarButton97(GetControl('BIMPCERTIF'));
        if BImprim <> Nil then BImprim.OnClick  :=  ImprimerCertifEmploi;
        Edit := THEdit(GetControl('BASECONGE'));
        If Edit <> Nil Then Edit.OnExit := ExitTHEdit;
        Edit := THEdit(GetControl('SALBRUT'));
        If Edit <> Nil Then Edit.OnExit := ExitTHEdit;
        Edit := THEdit(GetControl('DATEPAIE'));
        If Edit <> Nil Then Edit.OnExit := ExitTHEdit;
        Edit := THEdit(GetControl('DATEDELIV'));
        If Edit <> Nil Then Edit.OnExit := ExitTHEdit;
end ;

procedure TOM_CONGESSPEC.ExitDateDelivrance(Sender : TObject);
var aa,mm,jj : Word;
    Jour,Mois,Annee : String;
begin
        ForceUpdate;
        DecodeDate(StrToDate(GetControlText('DATEDELIV')),aa,mm,jj);
        Annee := IntToStr(aa);
        If mm<10 Then Mois := '0'+IntToStr(mm)
        Else Mois := IntToStr(mm);
        If jj<10 Then jour := '0'+IntToStr(jj)
        Else jour := IntToStr(jj);
        SetField('PCS_DATEDELIV',Jour+Mois+Annee);
end;

procedure TOM_CONGESSPEC.ExitDatePaie(Sender : TObject);
var aa,mm,jj : Word;
    Jour,Mois,Annee : String;
begin
        ForceUpdate;
        DecodeDate(StrToDate(GetControlText('DATEPAIE')),aa,mm,jj);
        Annee := IntToStr(aa);
        If mm<10 Then Mois := '0'+IntToStr(mm)
        Else Mois := IntToStr(mm);
        If jj<10 Then jour := '0'+IntToStr(jj)
        Else jour := IntToStr(jj);
        SetField('PCS_DATEPAIEM',Jour+Mois+Annee);
end;

procedure TOM_CONGESSPEC.CadreClick(Sender : TObject);
begin
        If GetControlText('CBCADRE') = 'X' Then
        begin
                If InitCadre=False Then ForceUpdate;
                SetField('PCS_CADRE','O');
        end
        Else
        begin
                If InitCadre=True Then ForceUpdate;
                SetField('PCS_CADRE','N');
        end;
end;

procedure TOM_CONGESSPEC.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_CONGESSPEC.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_CONGESSPEC.SalarieElipsisClick(Sender : TObject);
var St,Where : String;
begin
        St := 'SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE';
        Where :=' WHERE PSE_INTERMITTENT="X"';
        Where := RecupClauseHabilitationLookupList(Where);  //PT1
        {$IFNDEF EAGLCLIENT}
        LookupList(THDBEdit(Sender),'Liste des intermittents','SALARIES','PSA_SALARIE','PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM','','', True,-1,St+Where);
        {$ELSE}
//PT1        LookupList(THEdit(Sender),'Liste des intermittents','SALARIES','PSA_SALARIE','PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM',Where,'', True,-1,St);
        LookupList(THEdit(Sender),'Liste des intermittents','SALARIES','PSA_SALARIE','PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM','','', True,-1,St+Where);
        {$ENDIF}
end;

procedure TOM_CONGESSPEC.ImprimerCertifEmploi(Sender : TObject);
var St : String;
begin
//        pages := TPageControl(GetControl('PAGES'));
//        St := 'SELECT * FROM CONGESSPEC WHERE '+
//        'PCS_SALARIE="'+GetField('PCS_SALARIE')+'"';// +
//        'AND PCS_DATEDEBUT="'+UsDateTime(GetField('PCS_DATEDEBUT'))+'"' +
//        'AND PCS_DATEFIN="'+UsDateTime(GetField('PCS_DATEFIN'))+'"';     
//        LanceEtat('E','PCG','CSP',True,False,False,Pages,St,'',False);
        St := 'PCS_SALARIE='+GetField('PCS_SALARIE')+';'+
        'PCS_DATEDEBUT='+DateToStr(GetField('PCS_DATEDEBUT'))+';'+
        'PCS_DATEDEBUT_='+DateToStr(GetField('PCS_DATEDEBUT'))+';'+
        'TSIGNATAIRE='+GetField('PCS_NOMSIGN')+';'+
        'TLIEU='+GetField('PCS_LIEU')+';'+
        'TDATEDELIV='+GetControlText('DATEDELIV')+';'+
        'NUMCERTIF='+GetField('PCS_NUMCERTIFICAT')+';';
        AglLanceFiche ('PAY','EDIT_CONGESSPEC',St , '' , '' );
end;

procedure TOM_CONGESSPEC.ExitTHEdit(Sender : TObject);
{$IFDEF EAGLCLIENT}
var ModifDate,DatePaie,DateDEliv : String;
{$ENDIF}
begin
        If THEdit(Sender) = Nil then Exit;
        If THEdit(Sender).Name = 'BASECONGE' then
        begin
                if InitBaseConges <> getControlText('BASECONGE') then ForceUpdate;
                {$IFDEF EAGLCLIENT}
                if (InitBaseConges <> getControlText('BASECONGE')) and (IsNumeric(getControlText('BASECONGE'))) then
                begin
                        SetField('PCS_BASECONGE',StrToFloat(GetControlText('BASECONGE')));
                        InitBaseConges := GetControlText('BASECONGE');
                end;
                {$ENDIF}
        end;
        If THEdit(Sender).Name = 'SALBRUT' then
        begin
                if InitSalaire <> getControlText('SALBRUT') then ForceUpdate;
                {$IFDEF EAGLCLIENT}
                if (InitSalaire <> getControlText('SALBRUT')) and (IsNumeric(getControlText('SALBRUT'))) then
                begin
                        SetField('PCS_SALBRUT',StrToFloat(GetControlText('SALBRUT')));
                        InitSalaire := GetControlText('SALBRUT');
                end;
                {$ENDIF}
        end;
        If THEdit(Sender).Name = 'DATEPAIE' then
        begin
                if InitDatePaiem <> getControlText('DATEPAIE') then ForceUpdate;
                {$IFDEF EAGLCLIENT}
                if (InitDatePaiem <> getControlText('DATEPAIE')) and (IsValidDate(getControlText('DATEPAIE'))) then
                begin
                        DatePaie := getControlText('DATEPAIE');
                        ModifDate := ReadTokenPipe(DatePaie,'/');
                        ModifDate := ModifDate + ReadTokenPipe(DatePaie,'/');
                        ModifDate := ModifDate + ReadTokenPipe(DatePaie,'/');
                        SetField('PCS_DATEPAIEM',ModifDate);
                        InitDatePaiem := GetControlText('DATEPAIE');
                end;
                {$ENDIF}
        end;
        If THEdit(Sender).Name = 'DATEDELIV' then
        begin
                if InitDateDeliv <> getControlText('DATEDELIV') then ForceUpdate;
                {$IFDEF EAGLCLIENT}
                if (InitDateDeliv <> getControlText('DATEDELIV')) and (IsValidDate(getControlText('DATEPAIE'))) then
                begin
                        Datedeliv := getControlText('DATEDELIV');
                        ModifDate := ReadTokenPipe(Datedeliv,'/');
                        ModifDate := ModifDate + ReadTokenPipe(Datedeliv,'/');
                        ModifDate := ModifDate + ReadTokenPipe(Datedeliv,'/');
                        SetField('PCS_DATEDELIV',ModifDate);
                        InitDateDeliv := GetControlText('DATEDELIV');
                end;
                {$ENDIF}
        end;
end;

Initialization
  registerclasses ( [ TOM_CONGESSPEC ] ) ;
end.
