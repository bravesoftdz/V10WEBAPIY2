{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 23/06/2004
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : DECLARATIONS (DECLARATIONS)
Mots clefs ... : TOM;DECLARATIONS
*****************************************************************
PT1 04/01/2005 JL V_60 Requete non fermé
PT2 04/03/2005 JL V_60 FQ 12074 Bouton imprimer visible en CWAS
PT3 07/03/2005 JL V_60 Modif lance etat pour recuperer critere en CWAS
PT4 20/07/2005 JL V_60 FQ 12402 Correction qualité du déclarant
PT5 15/11/2005 JL V_650 FQ 12594 Gestion champs PDT_AUTREVICT
PT6 16/01/2006 JL V_650 FQ 12792 gestion de la zone Code risque
PT7 15/03/2006 JL V_650 FQ 10541 Date constaté = date accident par défaut
PT8 16/03/2007 V_72 RMA Déclaration accident du travail pour la MSA
PT9 07/06/2007 V_72 RMA FQ14334 Ergonomie
PT10 18/07/2008 V_80 JL Gestion accès depuis fiche salarié
}
Unit UTomDeclarations ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fiche,FichList,EdtREtat, DBCtrls,HDB,
{$else}
     eFiche,eFichList,UtilEAGl,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOM,UTob ,PgEditOutils,P5Util,HQRY,HTB97,paramsoc,PGOutils,EntPaie,PgOutils2,ParamDat;

Type
  TOM_DECLARATIONS = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    private
    EtabSalarie : String;
    Typedeclar  : String;  //PT8
    Lieu, Quand : String;  //PT8
    SalarieNew : String;  //PT10
    procedure InitSalarie;
    procedure CheckClick (Sender : TObject);
    procedure ImpressionDecl (Sender : TObject);
    procedure AffichDeclarant(Sender:TObject);
    procedure CheckRapportPolice (Sender : TObject);
    procedure CheckConstation (Sender : TObject);
    procedure DateElipsisclick (Sender : TObject);
    procedure CheckVictime (Sender : TObject);
    procedure AffectCodeRisque(Sender: TObject); { PT6 }
    procedure CheckMSA(Sender: TObject); { PT8 }
    Function  FImpr(St : string): string;  //PT9
    Procedure EditDate (Sender : TObject); //PT9
//    procedure CheckTiers (Sender : TObject);
    end ;

Implementation

procedure TOM_DECLARATIONS.OnCancelRecord ;
begin
  Inherited ;
        If Typedeclar = 'ACT' Then   //PT8
        Begin                        //PT8
           SetControlChecked('ARRETTRAVOUI',GetField('PDT_ARRETTRAVAIL') = 'X');
           SetControlChecked('ARRETTRAVNON',GetField('PDT_ARRETTRAVAIL') = '-');
        End;  //PT8
        SetControlChecked('ARRETTRAVNON1',GetField('PDT_ARRETTRAVAIL') = '-');
        SetControlChecked('ARRETTRAVOUI1',GetField('PDT_ARRETTRAVAIL') = 'X');
        SetControlChecked('POLICEOUI',GetField('PDT_RAPPORTPOLICE') = 'X');
        SetControlChecked('POLICENON',GetField('PDT_RAPPORTPOLICE') = '-');
        SetControlChecked('TIERSOUI',GetField('PDT_CAUSETIERS') = 'X');
        SetControlChecked('TIERSNON',GetField('PDT_CAUSETIERS') = '-');
        //PT8 Debut Ajout =====>
        If Typedeclar = 'MSA' Then
        Begin
           SetControlChecked('ACCAUTVICTOUI',GetField('PDT_AUTREVICT') = 'X');
           SetControlChecked('ACCAUTVICTNON',GetField('PDT_AUTREVICT') = '-');
           SetControlChecked('CKALLER',GetField('PDT_ALLERETOUR') = 'X');
           SetControlChecked('CKRETOUR',GetField('PDT_ALLERETOUR') = '-');
           SetControlChecked('CKLIEU1',False);
           SetControlChecked('CKLIEU2',False);
           SetControlChecked('CKLIEU3',False);
           SetControlChecked('CKCOUR1',False);
           SetControlChecked('CKCOUR2',False);
           SetControlChecked('CKCOUR3',False);
        End;
        //PT8 Fin Ajout <======
end ;

procedure TOM_DECLARATIONS.OnNewRecord ;
begin
  Inherited ;
  If SalarieNew <> '' then
  begin
    SetField('PDT_SALARIE',SalarieNew);
    InitSalarie;
    SalarieNew := '';
  end;
        //SetField('PDT_TYPEDECLAR','ACT');     //PT8
        SetField('PDT_TYPEDECLAR',Typedeclar);  //PT8
        SetField('PDT_ARRETTRAVAIL','-');
        SetField('PDT_RAPPORTPOLICE','-');
        SetField('PDT_CAUSETIERS','-');
        SetField('PDT_DATEACC',Date);
        //PT8 Debut Ajout =====>
        If Typedeclar = 'MSA' Then
        Begin
           SetField('PDT_AUTREVICT','-');
           SetField('PDT_ALLERETOUR','-');
           SetField('PDT_NATLIEUACC','00');
        End;
        //PT8 Fin Ajout <======
end ;

procedure TOM_DECLARATIONS.OnUpdateRecord ;
VAR IMax : Integer;
    QQ : TQuery;
    DA,Heure : TDateTime;
begin
  Inherited ;
    if DS.State in [dsInsert] then
    begin
      //PT8 QQ := OpenSQL('SELECT MAX(PDT_ORDRE) FROM DECLARATIONS WHERE PDT_SALARIE="' + GetField('PDT_SALARIE') + '"', TRUE);
      QQ := OpenSQL('SELECT MAX(PDT_ORDRE) FROM DECLARATIONS WHERE PDT_SALARIE="' + GetField('PDT_SALARIE') + '" And PDT_TYPEDECLAR="'+Typedeclar+'"', TRUE);
      if not QQ.EOF then imax := QQ.Fields[0].AsInteger + 1 else iMax := 1;
      Ferme(QQ);
      SetField('PDT_ORDRE', IMax);
    end;
    DA := GetField('PDT_DATEACC');
    Heure := GetField('PDT_HEUREACC');
    If Heure < IDate1900 then heure := DA + Heure;
    SetField('PDT_HEUREACC',Heure);
    Heure := GetField('PDT_HEUREDEBAM');
    If Heure < IDate1900 then heure := DA + Heure;
    SetField('PDT_HEUREDEBAM',Heure);
    Heure := GetField('PDT_HEUREFINAM');
    If Heure < IDate1900 then heure := DA + Heure;
    SetField('PDT_HEUREFINAM',Heure);
    Heure := GetField('PDT_HEUREDEBAM1');
    If Heure < IDate1900 then heure := DA + Heure;
    SetField('PDT_HEUREDEBAM1',Heure);
    Heure := GetField('PDT_HEUREFINPM');
    If Heure < IDate1900 then heure := DA + Heure;
    SetField('PDT_HEUREFINPM',Heure);
    Heure := GetField('PDT_HEURECONSTAT');
    If Heure < IDate1900 then heure := DA + Heure;
    SetField('PDT_HEURECONSTAT',Heure);

    If Typedeclar = 'MSA' Then                 //PT8
       SetField('PDT_NATLIEUACC',Lieu+Quand);  //PT8
end ;

procedure TOM_DECLARATIONS.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_DECLARATIONS.OnLoadRecord ;
begin
  Inherited ;
        If (DS.state <> dsInsert) then InitSalarie;
        If Typedeclar = 'ACT' Then   //PT8
        Begin                        //PT8
           SetControlChecked('ARRETTRAVOUI',GetField('PDT_ARRETTRAVAIL') = 'X');
           SetControlChecked('ARRETTRAVNON',GetField('PDT_ARRETTRAVAIL') = '-');
        End;  //PT8
        SetControlChecked('ARRETTRAVNON1',GetField('PDT_ARRETTRAVAIL') = '-');
        SetControlChecked('ARRETTRAVOUI1',GetField('PDT_ARRETTRAVAIL') = 'X');
        SetControlChecked('POLICEOUI',GetField('PDT_RAPPORTPOLICE') = 'X');
        SetControlChecked('POLICENON',GetField('PDT_RAPPORTPOLICE') = '-');
        SetControlChecked('TIERSOUI',GetField('PDT_CAUSETIERS') = 'X');
        SetControlChecked('TIERSNON',GetField('PDT_CAUSETIERS') = '-');
        SetControlChecked('ACCAUTVICTOUI',GetField('PDT_AUTREVICT') = 'X');
        SetControlChecked('ACCAUTVICTNON',GetField('PDT_AUTREVICT') = '-');
        //PT8 Debut Ajout =====>
        If Typedeclar = 'MSA' Then
        Begin
           SetControlChecked('CKALLER',GetField('PDT_ALLERETOUR') = 'X');
           SetControlChecked('CKRETOUR',GetField('PDT_ALLERETOUR') = '-');

           Lieu := Copy(GetField('PDT_NATLIEUACC'),1,1);
           Quand := Copy(GetField('PDT_NATLIEUACC'),2,1);

           SetControlChecked('CKLIEU1',False);
           SetControlChecked('CKLIEU2',False);
           SetControlChecked('CKLIEU3',False);
           SetControlChecked('CKCOUR1',False);
           SetControlChecked('CKCOUR2',False);
           SetControlChecked('CKCOUR3',False);
           If Lieu = '1' Then SetControlChecked('CKLIEU1',True);
           If Lieu = '2' Then SetControlChecked('CKLIEU2',True);
           If Lieu = '3' Then SetControlChecked('CKLIEU3',True);
           If Quand = '1' Then SetControlChecked('CKCOUR1',True);
           If Quand = '2' Then SetControlChecked('CKCOUR2',True);
           If Quand = '3' Then SetControlChecked('CKCOUR3',True);
        End;
        //PT8 Fin Ajout <======
End ;

procedure TOM_DECLARATIONS.OnChangeField ( F: TField ) ;
begin
  Inherited ;
        If F.FieldName = 'PDT_SALARIE' then
        begin
           if (DS.State in [dsInsert]) and (IsFieldModified('PDT_SALARIE')) then InitSalarie;
        end;
        //DEBUT PT7
        If F.FieldName = 'PDT_DATEACC' then
        begin
           If IsFieldModified('PDT_DATEACC') then //PT8 SetField('PDT_DATECONSTAT',GetField('PDT_DATEACC'));
           Begin //PT8 ==> Debut Ajout
             If (GetCheckBoxState('PDT_CONSTATE') <> CbChecked) And (GetCheckBoxState('PDT_CONNU') <> CbChecked) Then
                SetField('PDT_DATECONSTAT',IDate1900)
              Else
                SetField('PDT_DATECONSTAT',GetField('PDT_DATEACC'));
           End;  //PT8 Fin Ajout <===
        end;
        //FIN PT7
end ;

procedure TOM_DECLARATIONS.OnArgument ( S: String ) ;
var Check : TCheckBox;
    {$IFNDEF EAGLCLIENT}
    DbCheck : TDBCheckBox;
    THDate : THDBEdit;
    {$ELSE}
    DbCheck : TCheckBox;
    THDate : THEdit;
    {$ENDIF}
    BImp : TToolBarButton97;
    Lb : THLabel;
    Edit : THEdit;
    St : String;
    P : Integer; //PT8
    SalarieTemp : String;
begin
  Inherited ;
  St := S;
  SalarieTemp := ReadTokenPipe(St,';');
  SalarieTemp := ReadTokenPipe(St,';');
  SalarieNew := St;
  P := Pos(';',S);                   //PT8
  Typedeclar := Trim(copy(S,P+1,3)); //PT8
  Lieu := '0';                       //PT8
  Quand := '0';                      //PT8

        If Typedeclar = 'ACT' Then   //PT8
        Begin                        //PT8
           Check := TCheckBox(GetControl('ARRETTRAVOUI'));
           If Check <> Nil then Check.OnClick := CheckClick;
           Check := TCheckBox(GetControl('ARRETTRAVNON'));
           If Check <> Nil then Check.OnClick := CheckClick;
        End;                         //PT8
        Check := TCheckBox(GetControl('ARRETTRAVOUI1'));
        If Check <> Nil then Check.OnClick := CheckClick;
        Check := TCheckBox(GetControl('ARRETTRAVNON1'));
        If Check <> Nil then Check.OnClick := CheckClick;
        //PT8 Check := TCheckBox(GetControl('POLICEOUI'));
        //PT8 If Check <> Nil then Check.OnClick := CheckClick;
        //PT8 Check := TCheckBox(GetControl('POLICENON'));
        //PT8 If Check <> Nil then Check.OnClick := CheckClick;
        Check := TCheckBox(GetControl('TIERSOUI'));
        If Check <> Nil then Check.OnClick := CheckClick;
        Check := TCheckBox(GetControl('TIERSNON'));
        If Check <> Nil then Check.OnClick := CheckClick;
        Bimp:= TToolBarButton97(GetControl('BImprimer'));
        If BImp <> Nil then
        begin
             BImp.OnClick := ImpressionDecl;
             BImp.Visible := True; //PT2
        end;
        Lb := THLabel(GetControl('SOLIBELLE'));
        if Lb <> nil then Lb.Caption := GetParamSoc('SO_LIBELLE');
        If Typedeclar = 'ACT' Then   //PT8
        Begin                        //PT8
           Lb := THLabel(GetControl('SOADRESSE'));
           if Lb <> nil then Lb.Caption := GetParamSoc('SO_ADRESSE1') + '  ' +
              GetParamSoc('SO_ADRESSE2') + '   ' + GetParamSoc('SO_ADRESSE3');
           Lb := THLabel(GetControl('SOCPVILLE'));
           if Lb <> nil then Lb.Caption := GetParamSoc('SO_CODEPOSTAL') + '  ' + GetParamSoc('SO_VILLE');
           Lb := THLabel(GetControl('SOTELEPHONE'));
           if Lb <> nil then Lb.Caption := FormatCase(GetParamSoc('SO_TELEPHONE'), 'STR', 2);
        End;                         //PT8
        Edit := THEdit(GetControl('DECLARANT'));
        If Edit<>Nil Then Edit.OnExit := AffichDeclarant;
        Check := TCheckBox(GetControl('POLICEOUI'));
        If Check <> Nil then Check.OnClick := CheckRapportPolice;
        Check := TCheckBox(GetControl('POLICENON'));
        If Check <> Nil then Check.OnClick := CheckRapportPolice;
        Check := TCheckBox(GetControl('ACCAUTVICTOUI'));
        If Check <> Nil then Check.OnClick := CheckVictime;
        Check := TCheckBox(GetControl('ACCAUTVICTNON'));
        If Check <> Nil then Check.OnClick := CheckVictime;
        {$IFNDEF EAGLCLIENT}
        DbCheck := TDBCheckBox(GetControl('PDT_CONSTATE'));
        {$ELSE}
        DbCheck := TCheckBox(GetControl('PDT_CONSTATE'));
        {$ENDIF}
        If DBCheck <> Nil then DBCheck.OnClick := CheckConstation;
        {$IFNDEF EAGLCLIENT}
        DbCheck := TDBCheckBox(GetControl('PDT_CONNU'));
        THDate := THDBEdit(GetControl('PDT_DATEACC'));
        {$ELSE}
        DbCheck := TCheckBox(GetControl('PDT_CONNU'));
        THDate := THEdit(GetControl('PDT_DATEACC'));
        {$ENDIF}
        if THDate <> Nil then  //PT9
        Begin                  //PT9
           THDate.OnExit := AffectCodeRisque;
           THDate.OnChange := EditDate; //PT9
        End;                            //PT9
        If DBCheck <> Nil then DBCheck.OnClick := CheckConstation;
        EtabSalarie := '';
        St := 'PDA_ETABLISSEMENT = "" ' +
        ' AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%ACT%" )  ';
        SetControlProperty('DECLARANT', 'Plus', St);
        Edit := THEdit(GetControl('LE'));
        If Edit <> Nil then Edit.OnElipsisClick := DateElipsisclick;

        //PT8 Debut Ajout ========>
        If Typedeclar = 'MSA' Then
        Begin
           Check := TCheckBox(GetControl('CKLIEU1'));
           If Check <> Nil then Check.OnClick := CheckMSA;
           Check := TCheckBox(GetControl('CKLIEU2'));
           If Check <> Nil then Check.OnClick := CheckMSA;
           Check := TCheckBox(GetControl('CKLIEU3'));
           If Check <> Nil then Check.OnClick := CheckMSA;
           Check := TCheckBox(GetControl('CKCOUR1'));
           If Check <> Nil then Check.OnClick := CheckMSA;
           Check := TCheckBox(GetControl('CKCOUR2'));
           If Check <> Nil then Check.OnClick := CheckMSA;
           Check := TCheckBox(GetControl('CKCOUR3'));
           If Check <> Nil then Check.OnClick := CheckMSA;
           Check := TCheckBox(GetControl('CKALLER'));
           If Check <> Nil then Check.OnClick := CheckMSA;
           Check := TCheckBox(GetControl('CKRETOUR'));
           If Check <> Nil then Check.OnClick := CheckMSA;
           //PT9 Debut Ajout ===>
           {$IFNDEF EAGLCLIENT}
               THDate := THDBEdit(GetControl('PDT_HEUREACC'));
           {$ELSE}
               THDate := THEdit(GetControl('PDT_HEUREACC'));
           {$ENDIF}
           If THDate <> Nil then THDate.OnChange := EditDate;
           {$IFNDEF EAGLCLIENT}
               THDate := THDBEdit(GetControl('PDT_HEUREDEBAM'));
           {$ELSE}
               THDate := THEdit(GetControl('PDT_HEUREDEBAM'));
           {$ENDIF}
           If THDate <> Nil then THDate.OnChange := EditDate;
           {$IFNDEF EAGLCLIENT}
               THDate := THDBEdit(GetControl('PDT_HEUREFINAM'));
           {$ELSE}
               THDate := THEdit(GetControl('PDT_HEUREFINAM'));
           {$ENDIF}
           If THDate <> Nil then THDate.OnChange := EditDate;
           {$IFNDEF EAGLCLIENT}
               THDate := THDBEdit(GetControl('PDT_HEUREDEBAM1'));
           {$ELSE}
               THDate := THEdit(GetControl('PDT_HEUREDEBAM1'));
           {$ENDIF}
           If THDate <> Nil then THDate.OnChange := EditDate;
           {$IFNDEF EAGLCLIENT}
               THDate := THDBEdit(GetControl('PDT_HEUREFINPM'));
           {$ELSE}
               THDate := THEdit(GetControl('PDT_HEUREFINPM'));
           {$ENDIF}
           If THDate <> Nil then THDate.OnChange := EditDate;
           {$IFNDEF EAGLCLIENT}
               THDate := THDBEdit(GetControl('PDT_DATECONSTAT'));
           {$ELSE}
               THDate := THEdit(GetControl('PDT_DATECONSTAT'));
           {$ENDIF}
           If THDate <> Nil then THDate.OnChange := EditDate;
           {$IFNDEF EAGLCLIENT}
               THDate := THDBEdit(GetControl('PDT_HEURECONSTAT'));
           {$ELSE}
               THDate := THEdit(GetControl('PDT_HEURECONSTAT'));
           {$ENDIF}
           If THDate <> Nil then THDate.OnChange := EditDate;
           {$IFNDEF EAGLCLIENT}
               THDate := THDBEdit(GetControl('PDT_DATEREGINF'));
           {$ELSE}
               THDate := THEdit(GetControl('PDT_DATEREGINF'));
           {$ENDIF}
           If THDate <> Nil then THDate.OnChange := EditDate;
           //PT9 Fin Ajout <=======
        End;
        //PT8 Fin Ajout <=========
end ;

procedure TOM_DECLARATIONS.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_DECLARATIONS.InitSalarie;
var
  WDate: Word;
  QSal,QAttes: TQuery;
  Defaut: THLabel;
  CheckFr, CheckCee, CheckAutre: TCheckBox;
  St,LeSalarie : string;
  TXT : String ;  //PT9

begin
  LeSalarie := GetField('PDT_SALARIE');

  If Typedeclar = 'ACT' Then   //PT8
  Begin                        //PT8
     QSal := OpenSql('SELECT PSA_ETABLISSEMENT,PSA_SALARIE,PSA_LIBELLE,PSA_NOMJF,PSA_PRENOM,PSA_ADRESSE1,PSA_ADRESSE2,' +
       'PSA_ADRESSE3,PSA_CODEPOSTAL,PSA_VILLE,PSA_LIBELLEEMPLOI,PSA_NUMEROSS,' +
       'PSA_DATENAISSANCE,PSA_QUALIFICATION,PSA_DATEENTREE,PSA_DATEANCIENNETE,' +
       'PSA_NATIONALITE,PSA_SEXE, PSA_DATENAISSANCE,ET_ETABLISSEMENT,' +
       'ET_LIBELLE,ET_ADRESSE1,ET_ADRESSE2,ET_ADRESSE3,ET_CODEPOSTAL,ET_VILLE,' +
       'ET_TELEPHONE,ET_SIRET,ET_FAX,PAT_CODERISQUE '+
       'FROM SALARIES LEFT JOIN ETABLISS ON PSA_ETABLISSEMENT=ET_ETABLISSEMENT ' +
       'LEFT JOIN TAUXAT ON PSA_ETABLISSEMENT=PAT_ETABLISSEMENT AND PSA_ORDREAT=PAT_ORDREAT ' +  //PT6
       'WHERE PSA_SALARIE="' + LeSalarie + '" ORDER BY PAT_DATEVALIDITE DESC,PAT_ORDREAT DESC', TRUE);
  End  //PT8 + Debut Ajout ======>
  Else If Typedeclar = 'MSA' Then
  Begin
     QSal := OpenSql('SELECT PSA_ETABLISSEMENT,PSA_SALARIE,PSA_LIBELLE,PSA_NOMJF,PSA_PRENOM,PSA_ADRESSE1,PSA_ADRESSE2,' +
       'PSA_ADRESSE3,PSA_CODEPOSTAL,PSA_VILLE,PSA_LIBELLEEMPLOI,PSA_NUMEROSS,' +
       'PSA_DATENAISSANCE,PSA_QUALIFICATION,PSA_DATEENTREE,PSA_DATEANCIENNETE,' +
       'PSA_NATIONALITE,PSA_SEXE,ET_ETABLISSEMENT,' +
       'ET_LIBELLE,ET_ADRESSE1,ET_ADRESSE2,ET_ADRESSE3,ET_CODEPOSTAL,ET_VILLE,' +
       'ET_TELEPHONE,ET_SIRET,ETB_ETABLISSEMENT,ETB_NUMMSA,ETB_ACTIVITE '+
       'FROM SALARIES LEFT JOIN ETABLISS ON PSA_ETABLISSEMENT=ET_ETABLISSEMENT ' +
       'LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT ' +
       'WHERE PSA_SALARIE="' + LeSalarie + '"', TRUE);
  End; //PT8 + Fin Ajout <=======

  if QSal.eof then
  begin
    PGIBox('Salarié inexistant', Ecran.Caption);
    Ferme(QSal);
    exit;
  end;
  EtabSalarie := QSal.FindField('PSA_ETABLISSEMENT').asstring;
  St := '(PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%' + EtabSalarie + '%") ' +
        ' AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%ACT%" )  ';
        SetControlProperty('DECLARANT', 'Plus', St);
        QAttes := OpenSql('SELECT PDA_DECLARANTATTES FROM DECLARANTATTEST ' +
      'WHERE (PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%' + EtabSalarie + '%") ' +
      'AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%ACT%" )  ' +
      'ORDER BY PDA_ETABLISSEMENT DESC', True);
        if not QAttes.eof then
        begin
                SetControlText('DECLARANT', QAttes.FindField('PDA_DECLARANTATTES').AsString);
                AffichDeclarant(nil);
        end;
        Ferme(QAttes);

  Defaut := THLabel(GetControl('ETLIBELLE'));
  if defaut <> nil then Defaut.caption := QSal.FindField('ET_LIBELLE').asstring;
  Defaut := THLabel(GetControl('ETADRESSE123'));
  if defaut <> nil then Defaut.caption := QSal.FindField('ET_ADRESSE1').asstring + ' ' + QSal.FindField('ET_ADRESSE2').asstring + ' ' + QSal.FindField('ET_ADRESSE3').asstring;
  Defaut := THLabel(GetControl('ETCPVILLE'));
  If Typedeclar = 'ACT' Then   //PT8
  Begin                        //PT8
     if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('ET_CODEPOSTAL').asstring, 'STR', 3) + '   ' + QSal.FindField('ET_VILLE').asstring;
  End //PT8 + Debut Ajout ======>
  Else
  Begin
     if defaut <> nil then Defaut.caption := QSal.FindField('ET_VILLE').asstring;
     Defaut := THLabel(GetControl('ETCP'));
     if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('ET_CODEPOSTAL').asstring, 'STR', 4);
  End; //PT8 + Fin Ajout <=======

  Defaut := THLabel(GetControl('ETTELEPHONE'));
  if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('ET_TELEPHONE').asstring, 'STR', 2);
  If Typedeclar = 'ACT' Then   //PT8
  Begin                        //PT8
    Defaut := THLabel(GetControl('ETTELECOPIE'));
    if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('ET_FAX').asstring, 'STR', 2);
    Defaut := THLabel(GetControl('ETSIRET'));
    if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('ET_SIRET').asstring, 'STR', 3);
    Defaut := THLabel(GetControl('CODERISQUE')); //PT6 Recup du code risque
    if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('PAT_CODERISQUE').asstring, 'STR', 3);
  End //PT8 + Debut Ajout ======>
  Else
  Begin
    Defaut := THLabel(GetControl('ETSIRET'));
    if defaut <> nil then
       If Trim(QSal.FindField('ETB_NUMMSA').asstring) <> '' Then Defaut.caption := QSal.FindField('ETB_NUMMSA').asstring
          Else Defaut.caption := Copy(QSal.FindField('ET_SIRET').asstring,1,9);

    SetControlText('ETBACTIVITE',QSal.FindField('ETB_ACTIVITE').asstring);
  End; //PT8 + Fin Ajout <=======

  Defaut := THLabel(GetControl('PSALIBELLE'));
  if defaut <> nil then
    If Typedeclar = 'ACT' Then   //PT8
    Begin                        //PT8
      Defaut.caption := QSal.FindField('PSA_NOMJF').asstring + '  ' + QSal.FindField('PSA_LIBELLE').asstring + '   ' + QSal.FindField('PSA_PRENOM').asstring; { PT15 }
    End //PT8 + Debut Ajout ======>
    Else
    Begin
      If Trim(QSal.FindField('PSA_NOMJF').asstring) <> '' Then
      Begin
         Defaut.caption := QSal.FindField('PSA_NOMJF').asstring;
         Defaut := THLabel(GetControl('PSALIBELLE1'));
         If defaut <> nil Then Defaut.caption := QSal.FindField('PSA_LIBELLE').asstring;
      End
      Else Defaut.caption := QSal.FindField('PSA_LIBELLE').asstring;

     Defaut := THLabel(GetControl('PSAPRENOM'));
     If defaut <> nil Then Defaut.caption := QSal.FindField('PSA_PRENOM').asstring;
    End; //PT8 + Fin Ajout <=======

  Defaut := THLabel(GetControl('PSAADRESSE123'));
  if defaut <> nil then
    Defaut.caption := QSal.FindField('PSA_ADRESSE1').asstring + '  ' + QSal.FindField('PSA_ADRESSE2').asstring + '  ' + QSal.FindField('PSA_ADRESSE3').asstring;

  If Typedeclar = 'ACT' Then   //PT8
  Begin                        //PT8
    Defaut := THLabel(GetControl('PSASEXE'));
    if defaut <> nil then Defaut.caption := RechDom('PGSEXE', QSal.FindField('PSA_SEXE').asstring, False);
  End //PT8 + Debut Ajout ======>
  Else
  Begin
    If QSal.FindField('PSA_SEXE').asstring = 'M' Then
    Begin
       SetControlChecked('CKSEXM',True);
       SetControlChecked('CKSEXF',False);
    End
    Else
    Begin
       SetControlChecked('CKSEXM',False);
       SetControlChecked('CKSEXF',True);
    End;
  End; //PT8 + Fin Ajout <=======

  Defaut := THLabel(GetControl('PSADATENAISSANCE'));
  if defaut <> nil then Defaut.caption := QSal.FindField('PSA_DATENAISSANCE').asstring;

  Defaut := THLabel(GetControl('PSAEMPLOI'));
  if defaut <> nil then Defaut.caption := RechDom('PGLIBEMPLOI', QSal.FindField('PSA_LIBELLEEMPLOI').asstring, False);
  Defaut := THLabel(GetControl('PSAQUALIF'));
  if defaut <> nil then Defaut.caption := RechDom('PGLIBQUALIFICATION', QSal.FindField('PSA_QUALIFICATION').asstring, False);
  Defaut := THLabel(GetControl('PSADATEEMB'));
  if defaut <> nil then Defaut.caption := QSal.FindField('PSA_DATEENTREE').asstring;
  Defaut := THLabel(GetControl('PSAANC'));
  if defaut <> nil then
    if (QSal.FindField('PSA_DATEANCIENNETE').AsDateTime) > IDate1900 then
    begin
      WDate := AncienneteAnnee(QSal.FindField('PSA_DATEANCIENNETE').AsDateTime, date);
      if WDate > 1 then Defaut.caption := IntToStr(WDate) + ' ans'
      else Defaut.caption := IntToStr(WDate) + ' an';
    end
    else Defaut.caption := '';

  Defaut := THLabel(GetControl('PSACPVILLE'));
  If Typedeclar = 'ACT' Then   //PT8
  Begin                        //PT8
     if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('PSA_CODEPOSTAL').asstring, 'STR', 2) + '    ' + QSal.FindField('PSA_VILLE').asstring;
  End //PT8 + Debut Ajout ======>
  Else
  Begin
     if defaut <> nil then Defaut.caption := QSal.FindField('PSA_VILLE').asstring;
     Defaut := THLabel(GetControl('PSACP1'));
     if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('PSA_CODEPOSTAL').asstring, 'STR', 4);
  End; //PT8 + Fin Ajout <=======


  Defaut := THLabel(GetControl('PSASECU1'));
  if defaut <> nil then Defaut.caption := FormatCase(Copy(QSal.FindField('PSA_NUMEROSS').asstring, 1, 13), 'STR', 3);
  Defaut := THLabel(GetControl('PSASECU2'));
  if defaut <> nil then Defaut.caption := FormatCase(Copy(QSal.FindField('PSA_NUMEROSS').asstring, 14, 2), 'STR', 3); //PT1

  CheckFr := TCheckBox(GetControl('CKNATFRANC'));
  CheckCee := TCheckBox(GetControl('CKNATCEE'));
  CheckAutre := TCheckBox(GetControl('CKNATAUTRE'));
  if (QSal.FindField('PSA_NATIONALITE').asstring <> '') and (CheckFr <> nil) and (CheckCee <> nil) and (CheckAutre <> nil) then
  begin
    if (QSal.FindField('PSA_NATIONALITE').asstring) = 'FRA' then
    begin
      CheckFr.Checked := True;
      CheckCee.Enabled := False;
      CheckAutre.Enabled := False;
    end;
    if (QSal.FindField('PSA_NATIONALITE').asstring) <> 'FRA' then
      if (RechDom('PGCEE', (QSal.FindField('PSA_NATIONALITE').asstring), False) <> '') then
      begin
        CheckCee.Checked := True;
        CheckFr.Enabled := False;
        CheckAutre.Enabled := False;
      end;
    if (RechDom('PGCEE', (QSal.FindField('PSA_NATIONALITE').asstring), False) = '') then
    begin
      CheckAutre.Checked := True;
      CheckFr.Enabled := False;
      CheckCee.Enabled := False;
    end;
  end;

  //PT9 Debut Ajout =====>
  If Typedeclar = 'MSA' Then
  Begin
    TXT := FImpr(GetControlText('PSADATEEMB'));
    SetControlText('PSADATEEMB', TXT);
    TXT := FImpr(GetControlText('PSADATENAISSANCE'));
    SetControlText('PSADATENAISSANCE', TXT);
  End; //PT9 <====
  Ferme(QSal);
end;

procedure TOM_DECLARATIONS.CheckClick (Sender : TObject);
begin
        If TCheckBox(Sender) = Nil then Exit;
        If TCheckBox(Sender).Name = 'ARRETTRAVOUI' then
        begin
                If GetCheckBoxState('ARRETTRAVOUI') = CbChecked then
                begin
                        SetControlChecked('ARRETTRAVNON',False);
                        SetControlChecked('ARRETTRAVNON1',False);
                        SetControlChecked('ARRETTRAVOUI1',True);
                        SetField('PDT_ARRETTRAVAIL','X');
                end
                else
                begin
                        SetControlChecked('ARRETTRAVNON',True);
                        SetControlChecked('ARRETTRAVNON1',True);
                        SetControlChecked('ARRETTRAVOUI1',False);
                        SetField('PDT_ARRETTRAVAIL','-');
                end;
        end;
        If TCheckBox(Sender).Name = 'ARRETTRAVNON' then
        begin
                If GetCheckBoxState('ARRETTRAVNON') = CbChecked then
                begin
                        SetControlChecked('ARRETTRAVNON1',True);
                        SetControlChecked('ARRETTRAVOUI',False);
                        SetControlChecked('ARRETTRAVOUI1',False);
                        SetField('PDT_ARRETTRAVAIL','-');
                end
                else
                begin
                        SetControlChecked('ARRETTRAVNON1',False);
                        SetControlChecked('ARRETTRAVOUI',True);
                        SetControlChecked('ARRETTRAVOUI1',True);
                        SetField('PDT_ARRETTRAVAIL','X');
                end;
        end;
        If TCheckBox(Sender).Name = 'ARRETTRAVOUI1' then
        begin
                If GetCheckBoxState('ARRETTRAVOUI1') = CbChecked then
                begin
                        //SetControlChecked('ARRETTRAVNON',False);                           //PT8
                        If Typedeclar = 'ACT' Then SetControlChecked('ARRETTRAVNON',False);  //PT8
                        SetControlChecked('ARRETTRAVNON1',False);
                        //SetControlChecked('ARRETTRAVOUI',True);                            //PT8
                        If Typedeclar = 'ACT' Then SetControlChecked('ARRETTRAVOUI',True);   //PT8
                        SetField('PDT_ARRETTRAVAIL','X');
                end
                else
                begin
                        //SetControlChecked('ARRETTRAVNON',True);                           //PT8
                        If Typedeclar = 'ACT' Then SetControlChecked('ARRETTRAVNON',True);  //PT8
                        SetControlChecked('ARRETTRAVNON1',True);
                        //SetControlChecked('ARRETTRAVOUI',False);                          //PT8
                        If Typedeclar = 'ACT' Then SetControlChecked('ARRETTRAVOUI',False); //PT8
                        SetField('PDT_ARRETTRAVAIL','-');
                end;
        end;
        If TCheckBox(Sender).Name = 'ARRETTRAVNON1' then
        begin
                If GetCheckBoxState('ARRETTRAVNON1') = CbChecked then
                begin
                        //SetControlChecked('ARRETTRAVNON',True);                            PT8
                        //SetControlChecked('ARRETTRAVOUI',False);                           PT8
                        If Typedeclar = 'ACT' Then SetControlChecked('ARRETTRAVNON',True); //PT8
                        If Typedeclar = 'ACT' Then SetControlChecked('ARRETTRAVOUI',False);//PT8
                        SetControlChecked('ARRETTRAVOUI1',False);
                        SetField('PDT_ARRETTRAVAIL','-');
                end
                else
                begin
                        //SetControlChecked('ARRETTRAVNON',False);                          //PT8
                        //SetControlChecked('ARRETTRAVOUI',True);                           //PT8
                        If Typedeclar = 'ACT' Then SetControlChecked('ARRETTRAVNON',False); //PT8
                        If Typedeclar = 'ACT' Then SetControlChecked('ARRETTRAVOUI',True);  //PT8
                        SetControlChecked('ARRETTRAVOUI1',True);
                        SetField('PDT_ARRETTRAVAIL','X');
                end;
        end;
        {PT8=> If TCheckBox(Sender).Name = 'POLICEOUI' then
        begin
                If GetCheckBoxState('POLICEOUI') = CbChecked then
                begin
                        SetControlChecked('POLICENON',False);
                        SetField('PDT_RAPPORTPOLICE','X');
                end
                else
                begin
                        SetControlChecked('POLICENON',True);
                        SetField('PDT_RAPPORTPOLICE','-');
                end;
        end;
        If TCheckBox(Sender).Name = 'POLICENON' then
        begin
                If GetCheckBoxState('POLICENON') = CbChecked then
                begin
                        SetControlChecked('POLICEOUI',False);
                        SetField('PDT_RAPPORTPOLICE','-');
                end
                else
                begin
                        SetControlChecked('POLICEOUI',True);
                        SetField('PDT_RAPPORTPOLICE','X');
                end;
        end;     PT8<=}
        If TCheckBox(Sender).Name = 'TIERSOUI' then
        begin
                If GetCheckBoxState('TIERSOUI') = CbChecked then
                begin
                        SetControlChecked('TIERSNON',False);
                        SetField('PDT_CAUSETIERS','X');
                        SetControlEnabled('PDT_IDTIERS',True);
                        SetControlEnabled('PDT_ASSTIERS',True);
                end
                else
                begin
                        SetControlChecked('TIERSNON',True);
                        SetField('PDT_CAUSETIERS','-');
                        SetControlEnabled('PDT_IDTIERS',False);
                        SetControlEnabled('PDT_ASSTIERS',False);
                        SetField('PDT_IDTIERS','');
                        SetField('PDT_ASSTIERS','');
                end;
        end;
        If TCheckBox(Sender).Name = 'TIERSNON' then
        begin
                If GetCheckBoxState('TIERSNON') = CbChecked then
                begin
                        SetControlChecked('TIERSOUI',False);
                        SetField('PDT_CAUSETIERS','-');
                        SetControlEnabled('PDT_IDTIERS',False);
                        SetControlEnabled('PDT_ASSTIERS',False);
                        SetField('PDT_IDTIERS','');
                        SetField('PDT_ASSTIERS','');
                end
                else
                begin
                        SetControlChecked('TIERSOUI',True);
                        SetField('PDT_CAUSETIERS','X');
                        SetControlEnabled('PDT_IDTIERS',True);
                        SetControlEnabled('PDT_ASSTIERS',True);
                end;
        end;
end;

procedure TOM_DECLARATIONS.ImpressionDecl (Sender : TObject);
var Pages : TPageControl;
    StPages,requete : String;
begin
        Requete := 'SELECT * FROM DECLARATIONS '+
        'WHERE PDT_SALARIE="'+GetField('PDT_SALARIE')+'" AND PDT_TYPEDECLAR="'+GetField('PDT_TYPEDECLAR')+'" AND PDT_ORDRE='+IntToStr(GetField('PDT_ORDRE'));
        Pages := TPageControl(Getcontrol('PAGES'));
        StPages := AglGetCriteres (Pages, FALSE);
        If Typedeclar = 'ACT' Then   //PT8
           LanceEtat('E','PAT','PDA',True,False,False,Pages,Requete,'',False,0,StPages)
           Else LanceEtat('E','PAT','PMA',True,False,False,Pages,Requete,'',False,0,StPages); //PT8

end;

procedure TOM_DECLARATIONS.AffichDeclarant(Sender: TObject);
var St : String;
begin
if GetControlText('DECLARANT')='' then exit;
SetControlText('SIGNATAIRE',RechDom('PGDECLARANTATTEST',GetControlText('DECLARANT'),False));
St := RechDom('PGDECLARANTQUAL'  ,GetControlText('DECLARANT'),False);
if St = 'AUT' then SetControlText('QUALITE' ,RechDom('PGDECLARANTAUTRE' ,GetControlText('DECLARANT'),False))
else               SetControlText('QUALITE' ,RechDom('PGQUALDECLARANT2' ,St,False));
If GetField('PDT_FAITA') <> (RechDom('PGDECLARANTVILLE' ,GetControlText('DECLARANT'),False)) then
SetField('PDT_FAITA',RechDom('PGDECLARANTVILLE' ,GetControlText('DECLARANT'),False));
end;

procedure TOM_DECLARATIONS.CheckRapportPolice (Sender : TObject);
begin
        If TCheckBox(Sender).Name = 'POLICEOUI' then
        begin
                If GetCheckBoxState('POLICEOUI') = CbChecked then
                begin
                        SetControlChecked('POLICENON',False);
                        If GetField('PDT_RAPPORTPOLICE') <> 'X' then
                        begin
                                ForceUpdate;
                                SetField('PDT_RAPPORTPOLICE','X');
                                SetControlChecked('PDT_RAPPORTPOLICE',True);
                                UpdateChamp('PDT_RAPPORTPOLICE');
                        end;
                        SetControlEnabled('PDT_AUTRAPPOLICE',True);
                end;
        end;
        If TCheckBox(Sender).Name = 'POLICENON' then
        begin
                If GetCheckBoxState('POLICENON') = CbChecked then
                begin
                        SetControlChecked('POLICEOUI',False);
                        If GetField('PDT_RAPPORTPOLICE') <> '-' then
                        begin
                                ForceUpdate;
                                SetField('PDT_RAPPORTPOLICE','-');
                                SetControlChecked('PDT_RAPPORTPOLICE',False);
                                UpdateChamp('PDT_RAPPORTPOLICE');
                        end;
                        SetControlEnabled('PDT_AUTRAPPOLICE',False);
                end;
        end;
end;

procedure TOM_DECLARATIONS.CheckConstation (Sender : TObject);
begin
     If (DS.state = dsBrowse) then exit;  //PT8
     If TCheckBox(Sender).Name = 'PDT_CONSTATE' then
     begin
          If GetCheckBoxState('PDT_CONSTATE') = CbChecked then SetField('PDT_DATECONSTAT',GetField('PDT_DATEACC'))
          else
          begin
               If GetCheckBoxState('PDT_CONNU') <> CbChecked then SetField('PDT_DATECONSTAT',IDate1900);
          end;
     end;
     If TCheckBox(Sender).Name = 'PDT_CONNU' then
     begin
          If GetCheckBoxState('PDT_CONNU') = CbChecked then SetField('PDT_DATECONSTAT',GetField('PDT_DATEACC'))
          else
          begin
               If GetCheckBoxState('PDT_CONSTATE') <> CbChecked then SetField('PDT_DATECONSTAT',IDate1900);
          end;
     end;
end;

procedure TOM_DECLARATIONS.CheckVictime (Sender : TObject);
begin
        If TCheckBox(Sender).Name = 'ACCAUTVICTOUI' then
        begin
                If GetCheckBoxState('ACCAUTVICTOUI') = CbChecked then
                begin
                        SetControlChecked('ACCAUTVICTNON',False);
                        If GetField('PDT_AUTREVICT') <> 'X' then
                        begin
                                ForceUpdate;
                                SetField('PDT_AUTREVICT','X');
                                SetControlChecked('PDT_AUTREVICT',True);
                                UpdateChamp('PDT_AUTREVICT');
                        end;
                end;
        end;
        If TCheckBox(Sender).Name = 'ACCAUTVICTNON' then
        begin
                If GetCheckBoxState('ACCAUTVICTNON') = CbChecked then
                begin
                        SetControlChecked('ACCAUTVICTOUI',False);
                        If GetField('PDT_AUTREVICT') <> '-' then
                        begin
                                ForceUpdate;
                                SetField('PDT_AUTREVICT','-');
                                SetControlChecked('PDT_AUTREVICT',False);
                                UpdateChamp('PDT_AUTREVICT');
                        end;
                end;
        end;
end;

//PT8 Ajout de la procedure CheckMSA
Procedure TOM_DECLARATIONS.CheckMSA (Sender : TObject);
Begin
  If TCheckBox(Sender).Name = 'CKLIEU1' then
     If GetCheckBoxState('CKLIEU1') = CbChecked then
     Begin
        Lieu := '1';
        SetControlChecked('CKLIEU2',False);
        SetControlChecked('CKLIEU3',False);
     End;
  If TCheckBox(Sender).Name = 'CKLIEU2' then
     If GetCheckBoxState('CKLIEU2') = CbChecked then
     Begin
        Lieu := '2';
        SetControlChecked('CKLIEU1',False);
        SetControlChecked('CKLIEU3',False);
     End;
  If TCheckBox(Sender).Name = 'CKLIEU3' then
     If GetCheckBoxState('CKLIEU3') = CbChecked then
     Begin
        Lieu := '3';
        SetControlChecked('CKLIEU1',False);
        SetControlChecked('CKLIEU2',False);
     End;
  If TCheckBox(Sender).Name = 'CKCOUR1' then
     If GetCheckBoxState('CKCOUR1') = CbChecked then
     Begin
        Quand := '1';
        SetControlChecked('CKCOUR2',False);
        SetControlChecked('CKCOUR3',False);
     End;
  If TCheckBox(Sender).Name = 'CKCOUR2' then
     If GetCheckBoxState('CKCOUR2') = CbChecked then
     Begin
        Quand := '2';
        SetControlChecked('CKCOUR1',False);
        SetControlChecked('CKCOUR3',False);
     End;
  If TCheckBox(Sender).Name = 'CKCOUR3' then
     If GetCheckBoxState('CKCOUR3') = CbChecked then
     Begin
        Quand := '3';
        SetControlChecked('CKCOUR1',False);
        SetControlChecked('CKCOUR2',False);
     End;
  If TCheckBox(Sender).Name = 'CKALLER' then
  Begin
     If GetCheckBoxState('CKALLER') = CbChecked then
     Begin
        SetControlChecked('CKRETOUR',False);
        If GetField('PDT_ALLERETOUR') <> 'X' then
        Begin
           ForceUpdate;
           SetField('PDT_ALLERETOUR','X');
           SetControlChecked('PDT_ALLERETOUR',True);
           UpdateChamp('PDT_ALLERETOUR');
        End;
     End;
  End;
  If TCheckBox(Sender).Name = 'CKRETOUR' then
  Begin
     If GetCheckBoxState('CKRETOUR') = CbChecked then
     Begin
        SetControlChecked('CKALLER',False);
        If GetField('PDT_ALLERETOUR') <> '-' then
        Begin
           ForceUpdate;
           SetField('PDT_ALLERETOUR','-');
           SetControlChecked('PDT_ALLERETOUR',False);
           UpdateChamp('PDT_ALLERETOUR');
        End;
     End;
  End;
End;

{procedure TOM_DECLARATIONS.CheckTiers (Sender : TObject);
begin
        If TCheckBox(Sender).Name = 'TIERSOUI' then
        begin
                If GetCheckBoxState('TIERSOUI') = CbChecked then
                begin
                        SetControlChecked('TIERSNON',False);
                        If GetField('PDT_CAUSETIERS') <> 'X' then
                        begin
                                ForceUpdate;
                                SetField('PDT_CAUSETIERS','X');
                                SetControlChecked('PDT_CAUSETIERS',True);
                                UpdateChamp('PDT_CAUSETIERS');
                        end;
                        SetControlEnabled('PDT_AUTRAPTIERS',True);
                end;
        end;
        If TCheckBox(Sender).Name = 'TIERSNON' then
        begin
                If GetCheckBoxState('TIERSNON') = CbChecked then
                begin
                        SetControlChecked('TIERSOUI',False);
                        If GetField('PDT_CAUSETIERS') <> '-' then
                        begin
                                ForceUpdate;
                                SetField('PDT_CAUSETIERS','-');
                                SetControlChecked('PDT_CAUSETIERS',False);
                                UpdateChamp('PDT_CAUSETIERS');
                        end;
                        SetControlEnabled('PDT_AUTRAPTIERS',False);
                end;
        end;
end;       }

procedure TOM_DECLARATIONS.DateElipsisclick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

{ DEB PT6 }
procedure TOM_DECLARATIONS.AffectCodeRisque(Sender: TObject);
var
  DateAcc: TDateTime;
  Q: TQuery;
begin
  If Typedeclar = 'MSA' Then Exit;  //PT8
  DateAcc := GetField('PDT_DATEACC');

  Q := OpenSql('SELECT PAT_CODERISQUE,PAT_ORDREAT,PAT_DATEVALIDITE  ' +
    'FROM TAUXAT '+
    'LEFT JOIN SALARIES ON PSA_ETABLISSEMENT=PAT_ETABLISSEMENT AND PSA_ORDREAT=PAT_ORDREAT ' +  { PT18 }
    'WHERE PSA_SALARIE="' + GetField('PDT_SALARIE') + '" AND PAT_DATEVALIDITE<="' + USDateTime(DateAcc) + '" ' +
    'ORDER BY PAT_DATEVALIDITE DESC,PAT_ORDREAT DESC', True);
  if not Q.eof then
    SetControlText('CODERISQUE', FormatCase(Q.FindField('PAT_CODERISQUE').asstring, 'STR', 3));
  Ferme(Q);
end;
{ FIN PT6 }

//PT9 Ajout de la fonction FImpr ==>
Function TOM_DECLARATIONS.FImpr(St : string): string;
Var
  StFormat: string;
  i,j,Lg,Esp : integer;
Begin
  If St = '' Then Begin result:=''; Exit; End;
  If St = '01/01/1900' Then Begin result:=''; Exit; End;
  If St = '00:00' Then Begin result:=''; Exit; End;
  Lg := Length(St);
  StFormat := Stringofchar (' ',Lg+5*(Lg-1));
  j:=1;
  Esp := 4;
  For i:=1 to Lg do
    If (St[i]<>'/') And (St[i]<>':') then
       Begin
         StFormat[j] := St[i];
         j := j+Esp;
       End;
  result:=Trim(StFormat);
End; //PT9 <== FIN

//PT9 Ajout de la Procedure EditDate ==>
Procedure TOM_DECLARATIONS.EditDate(Sender: TObject);
var
  TXT : String;
Begin
  If THEdit(Sender).Name ='PDT_DATEACC' Then
  Begin
    TXT := FImpr(GetControlText('PDT_DATEACC'));
    SetControlText('PDTDATEACC', TXT);
  End;
  If THEdit(Sender).Name ='PDT_HEUREACC' Then
  Begin
    TXT := FImpr(GetControlText('PDT_HEUREACC'));
    SetControlText('PDTHEUREACC', TXT);
  End;
  If THEdit(Sender).Name ='PDT_HEUREDEBAM' Then
  Begin
    TXT := FImpr(GetControlText('PDT_HEUREDEBAM'));
    SetControlText('PDTHEUREDEBAM', TXT);
  End;
  If THEdit(Sender).Name ='PDT_HEUREFINAM' Then
  Begin
    TXT := FImpr(GetControlText('PDT_HEUREFINAM'));
    SetControlText('PDTHEUREFINAM', TXT);
  End;
  If THEdit(Sender).Name ='PDT_HEUREDEBAM1' Then
  Begin
    TXT := FImpr(GetControlText('PDT_HEUREDEBAM1'));
    SetControlText('PDTHEUREDEBAM1', TXT);
  End;
  If THEdit(Sender).Name ='PDT_HEUREFINPM' Then
  Begin
    TXT := FImpr(GetControlText('PDT_HEUREFINPM'));
    SetControlText('PDTHEUREFINPM', TXT);
  End;
  If THEdit(Sender).Name ='PDT_DATECONSTAT' Then
  Begin
    TXT := FImpr(GetControlText('PDT_DATECONSTAT'));
    SetControlText('PDTDATECONSTAT', TXT);
  End;
  If THEdit(Sender).Name ='PDT_HEURECONSTAT' Then
  Begin
    TXT := FImpr(GetControlText('PDT_HEURECONSTAT'));
    SetControlText('PDTHEURECONSTAT', TXT);
  End;
  If THEdit(Sender).Name ='PDT_DATEREGINF' Then
  Begin
    TXT := FImpr(GetControlText('PDT_DATEREGINF'));
    SetControlText('PDTDATEREGINF', TXT);
  End;
End; //PT9 <== FIN

Initialization
  registerclasses ( [ TOM_DECLARATIONS ] ) ;
end.

