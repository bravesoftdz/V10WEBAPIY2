{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 15/03/2002
Modifié le ... :   /  /
Description .. : TOM Gestion de la liste des animateurs d'une session
Mots clefs ... : PAIE
*****************************************************************
PT1 16/12/2003 V_50 JL - gestion liste salariés sur elipsisclick
                       - Frais inclus dans ct pédagogique par défaut
                       - Message si salarié anime une session avecmêmes dates = non bloquants
PT2 26/01/2007 V_80 FC Mise en place du filtage habilitation pour les lookuplist
                        pour les critères code salarié uniquement
PT3 30/01/2007 V_80 JL Supp test sur exercices, et prise en compte de l'exercice de la date de fin pour calcul
                        du salarie catégoreil
PT4 18/04/2007 V_720 FL FQ 14071 Prise en compte de l'exercice de formation pour le calcul du coût animateur
PT5 02/04/2008 V_80  FL Prise en compte du multidossier formation
 }
unit UTOMSESSIONANIMAT;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,Spin,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBCtrls,Fe_Main,Fiche,
{$ELSE}
       MaineAgl,eFiche,
{$ENDIF}
      HCtrls,HEnt1,HMsgBox,UTOM,UTOB,HTB97,PgOutils,PgoutilsFormation,EntPaie,LookUp ;

Type
     TOM_SESSIONANIMAT = Class(TOM)
       procedure OnArgument (stArgument : String ) ; override ;
       procedure OnNewRecord  ; override ;
       procedure OnUpdateRecord ; override ;
       procedure OnChangeField(F: TField)  ;  override ;
       procedure OnLoadRecord ; override ;
       procedure OnAfterUpdateRecord; override;
       private
       LeSalarie,LeStage,Millesime : String;
       Ordre   : Integer;
       DD,DF   : TDateTime;
       AvecClient:Boolean;
       InitNbHeure : Double;
       procedure SalarieElipsisClick (Sender : TObject);
       Procedure MAJCouts () ; //PT4
     END ;

implementation

Uses GalOutil;

{ TOM_SESSIONANIMAT }

procedure TOM_SESSIONANIMAT.OnArgument(stArgument: String);
var       St : String;
          {$IFNDEF EAGLCLIENT}
          Edit : THDBEdit;
          {$ELSE}
          Edit : THEdit;
          {$ENDIF}
begin
  inherited;
        LeStage := '';
        Ordre   := 0;
        st:=Trim (StArgument);
        LeStage :=ReadTokenSt(St);// recup du code du stage
        if LeStage <> '' then
        begin
                SetControlEnabled ('PAN_CODESTAGE',FALSE);
                Ordre := StrToInt (ReadTokenSt(St));
                DD    := StrToDate (ReadTokenSt(St));
                DF    := StrToDate (ReadTokenSt(St));
                Millesime:=ReadTokenST(St);
        end;
        AvecClient:=False;
        {$IFNDEF EAGLCLIENT}
        Edit := THDBEdit(GetControl('PAN_SALARIE'));
        {$ELSE}
        Edit := THEdit(GetControl('PAN_SALARIE'));
        {$ENDIF}
        If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;          //PT1
end;

procedure TOM_SESSIONANIMAT.OnChangeField(F: TField);
var Q     : TQuery;
    St    : String;
    DDSession,DFSession,DDAnim,DFAnim : Variant;
    NbJoursSession,NbJoursAnim : Integer;
    NbHeures : Double;
begin
  inherited;
        if (F.FieldName ='PAN_SALARIE') AND (GetField ('PAN_SALARIE') <> '')then
        begin
                If Not PGBundleHierarchie then //PT5
                begin
                  St := 'SELECT PSA_LIBELLE,PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE="'+GetField ('PAN_SALARIE')+'"';
                  Q := OpenSQL(st,TRUE) ;
                  if Not Q.EOF then St := Q.FindField('PSA_LIBELLE').AsString+' '+Q.FindField('PSA_PRENOM').AsString
                  else st := '';
                  Ferme (Q);
                end
                else
                begin
                  St := 'SELECT PSI_LIBELLE,PSI_PRENOM,PSI_PREDEFINI,PSI_NODOSSIER FROM INTERIMAIRES WHERE PSI_INTERIMAIRE="'+GetField ('PAN_SALARIE')+'"';
                  Q := OpenSQL(st,TRUE) ;
                  if Not Q.EOF then
                  begin
                    St := Q.FindField('PSI_LIBELLE').AsString+' '+Q.FindField('PSI_PRENOM').AsString;
                   SetField ('PAN_PREDEFINI','DOS');
                   SetField (' PAN_NODOSSIER',Q.FindField('PSI_NODOSSIER').AsString);
                  end
                  else st := '';
                  Ferme (Q);
                end;
                SetField ('PAN_LIBELLE',St);
                If (AvecClient = False) and (LeSalarie <> GetField('PAN_SALARIE')) Then
                begin
                        MajCouts(); //PT4
                end;
        end;
        if (F.FieldName ='PAN_SALARIE') AND (GetField ('PAN_SALARIE') = '')then
        begin
                If (AvecClient = False) and (LeSalarie <> GetField('PAN_SALARIE')) Then
                begin
                        LeSalarie := GetField('PAN_SALARIE');
                        SetField('PAN_SALAIREANIM',0);
                end;
        end;
        if F.FieldName='PAN_NBREHEURE' Then
        begin
                If (AvecClient = False) and (InitNbHeure <> GetField('PAN_NBREHEURE')) and (GetField('PAN_SALARIE') <> '') Then
                begin
                        InitNbHeure := GetField('PAN_NBREHEURE');
                        MajCouts(); //PT4
                end;
        end;
        If F.FieldName=('PAN_DATEFIN') Then
        begin
                Q:=OpenSQL('SELECT PSS_DATEDEBUT,PSS_DATEFIN,PSS_DUREESTAGE FROM SESSIONSTAGE'+
                ' WHERE PSS_CODESTAGE="'+GetField('PAN_CODESTAGE')+'"'+
                ' AND PSS_MILLESIME="'+GetField('PAN_MILLESIME')+'" AND PSS_ORDRE='+IntToStr(GetField('PAN_ORDRE')),True);  //DB2
                DDSession := IDate1900;
                DFSession := IDate1900;
                NbHeures := 0;
                if not Q.eof then
                begin
                        DDSession := Q.FindField('PSS_DATEDEBUT').AsDateTime;
                        DFSession := Q.FindField('PSS_DATEFIN').AsDateTime;
                        NbHeures := Q.FindField('PSS_DUREESTAGE').AsFloat;
                end;
                Ferme(Q);
                DDAnim := Getfield('PAN_DATEDEBUT');
                DFAnim := GetField('PAN_DATEFIN');
                If (DDSession <> IDate1900) and (DFSession <> IDate1900) then
                begin
                        If (DDSession <> DDAnim) or (DFSession <> DFAnim) Then
                        begin
                                NbJoursAnim := DFAnim - DDAnim;
                                NbJoursAnim := NbJoursAnim + 1;
                                NbJoursSession := DFSession - DDSession;
                                NbJoursSession := NbjoursSession + 1;
                                NbHeures:=NbHeures / NbJoursSession;
                                SetField('PAN_NBREHEURE',NbHeures * NbJoursAnim);
                        end;
                end;
        end;
end;

procedure TOM_SESSIONANIMAT.OnLoadRecord;
var Q : TQuery;
begin
  inherited;
        if LeStage = '' then
        begin
                LeStage := GetField ('PAN_CODESTAGE');
                Ordre   := GetField ('PAN_ORDRE');
                DD      := GetField ('PAN_DATEDEBUT');
                DF      := GetField ('PAN_DATEFIN');
                Millesime := GetField('PAN_MILLESIME');
        end;
        LeSalarie := Getfield('PAN_SALARIE');
        InitNbHeure := GetField('PAN_NBREHEURE');
        If ExisteSQL('SELECT PSS_ORDRE FROM SESSIONSTAGE WHERE PSS_ORDRE=' + IntToStr(Ordre) + ' '+  //DB2
             'AND PSS_MILLESIME="'+Millesime+'" AND PSS_CODESTAGE="' + LeStage + '" AND PSS_NATUREFORM="002"') then SetControlEnabled('PAN_SALARIE',False)
        Else SetControlEnabled('PAN_SALARIE',True);
        Q := OpenSQL('SELECT PSS_AVECCLIENT FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+LeStage+'" AND PSS_MILLESIME="' + Millesime + '" AND PSS_ORDRE=' + IntToStr(Ordre) + '',True);  //DB2
        if not Q.eof then
        begin
                If Q.FindField('PSS_AVECCLIENT').AsString = 'X' Then AvecClient := True
                Else AvecClient := False;
        end;
        Ferme(Q);
end;

procedure TOM_SESSIONANIMAT.OnNewRecord;
var Q : TQuery;
begin
  inherited;
        SetControlEnabled ('POS_ORDRE',FALSE);
        SetField ('PAN_CODESTAGE', LeStage);
        If LeStage<>'' Then
        begin
                SetControlEnabled('PAN_CODESTAGE',False);
                Q:=OpenSQL('SELECT PSS_DUREESTAGE FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+LeStage+'" AND PSS_ORDRE='+IntToStr(Ordre)+' AND PSS_MILLESIME="'+Millesime+'"',True);  //DB2
                if not Q.eof then SetField('PAN_NBREHEURE',Q.FindField('PSS_DUREESTAGE').AsFloat);
                Ferme(Q);
        end;
        SetField ('PAN_ORDRE', Ordre);
        SetField ('PAN_DATEDEBUT', DD);
        SetField ('PAN_DATEFIN', DF);
        SetField('PAN_MILLESIME',Millesime);        
        If (GetField('PAN_CODESTAGE') = '') then
        begin
                PgiBox ('Vous devez saisir un stage', Ecran.caption);
                SetFocusControl ('PAN_CODESTAGE');
                exit;
        end;
        SetField ('PAN_DEBUTDJ','MAT');
        SetField ('PAN_FINDJ','PAM');
        SetField('PAN_FRASPEDAG','X'); //PT1
        If PGBundleInscFormation then //PT5
        begin
          SetField('PAN_PREDEFINI','DOS');
          SetField('PAN_NODOSSIER',V_PGI.NoDossier);
        end
        else
        begin
          SetField('PAN_PREDEFINI','DOS');
          SetField('PAN_NODOSSIER','000000');
        end;
end;

procedure TOM_SESSIONANIMAT.OnUpdateRecord;
Var QQ,Q      : TQuery ;
    IMax    :integer ;
    St,MessError      : String;
    DebSession,FinSession : TDateTime;
begin
        MessError := '';
        If (GetField('PAN_ORDRE')=0) OR (GetField('POS_CODESTAGE') = '') then
        begin
                MessError:='#13#10- Vous devez renseigner le stage et le numéro de session';
        end;
        If GetField('PAN_SALARIE')<>'' then
        begin
                If ExisteSQL('SELECT PAN_SALARIE FROM SESSIONANIMAT WHERE PAN_SALARIE="'+Getfield('PAN_SALARIE')+'" '+
                'AND (PAN_CODESTAGE<>"'+GetField('PAN_CODESTAGE')+'" OR PAN_ORDRE<>'+IntToStr(GetField('PAN_ORDRE'))+' OR PAN_MILLESIME<>"'+Getfield('PAN_MILLESIME')+'") '+  //DB2
                'AND ((PAN_DATEDEBUT>="'+UsDateTime(GetField('PAN_DATEDEBUT'))+'" AND PAN_DATEDEBUT<="'+UsDateTime(GetField('PAN_DATEFIN'))+'") '+
                'OR (PAN_DATEFIN>="'+UsDateTime(GetField('PAN_DATEDEBUT'))+'" AND PAN_DATEFIN<="'+UsDateTime(GetField('PAN_DATEFIN'))+'") OR '+
                '(PAN_DATEDEBUT<="'+UsDateTime(GetField('PAN_DATEDEBUT'))+'" AND PAN_DATEFIN>="'+UsDateTime(GetField('PAN_DATEFIN'))+'"))') then
                begin
                        PGIInfo('Le salarié anime déja une session dont les dates coïncides',Ecran.Caption); //PT1
                end;
        end;
        DebSession := IDate1900;
        FinSession := IDate1900;
        Q:=OpenSQL('SELECT PSS_DATEDEBUT,PSS_DATEFIN FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+GetField('PAN_CODESTAGE')+'" '+
           'AND PSS_ORDRE='+IntToStr(GetField('PAN_ORDRE'))+' AND PSS_MILLESIME="'+GetField('PAN_MILLESIME')+'"',True);  //DB2
        if not Q.eof then
        begin
                DebSession := Q.FindField('PSS_DATEDEBUT').AsDateTime;
                FinSession := Q.FindField('PSS_DATEFIN').AsDateTime;
        end;
        Ferme(Q);
        If (GetField('PAN_DATEDEBUT')<DebSession) or (GetField('PAN_DATEDEBUT')>FinSession) then
        MessError := MessError+'#13#10- La date de début doit être comprise entre le '+DateToStr(DebSession)+' et le '+DateToStr(FinSession);
        If (GetField('PAN_DATEFIN')<DebSession) or (GetField('PAN_DATEFIN')>FinSession) then
        MessError := MessError+'#13#10- La date de fin doit être comprise entre le '+DateToStr(DebSession)+' et le '+DateToStr(FinSession);
        If MessError<>'' then
        begin
                PGIBox('Attention : '+MessError,Ecran.Caption);
                LastError := 1;
        end;
        If (DS.State in [dsInsert]) then
        begin     // increment automatique du numero d'ordre au moment de la creation
                st := 'SELECT MAX(PAN_RANG) FROM SESSIONANIMAT WHERE PAN_ORDRE='+IntToStr(GetField('PAN_ORDRE'))+  //DB2
                ' AND PAN_CODESTAGE="'+GetField('PAN_CODESTAGE')+'"';
                QQ:=OpenSQL(st,TRUE) ;
                if Not QQ.EOF then
                begin
                        IMax := QQ.Fields[0].AsInteger;
                        if IMax <> 0 then
                        IMax := IMax + 1
                        else
                        IMax := 1;
                end
                else IMax := 1 ;
                Ferme(QQ) ;
                SetField ('PAN_RANG', IMax);
        end;
end;

procedure TOM_SESSIONANIMAT.OnAfterUpdateRecord;
var QS,QF : TQuery;
    TobSessions,TS : Tob;
    Montant : Double;
begin
    QS:=OpenSQL('SELECT * FROM SESSIONSTAGE WHERE '+
        'PSS_CODESTAGE="'+GetField('PAN_CODESTAGE')+'" AND PSS_MILLESIME="'+GetField('PAN_MILLESIME')+'"'+
        ' AND PSS_ORDRE='+IntToStr(GetField('PAN_ORDRE'))+'',True);  //DB2
    TobSessions := Tob.create('SESSIONSTAGE',Nil,-1);
    TobSessions.LoadDetailDB('SESSIONSTAGE','','',QS,False);
    Ferme(QS);
    QF := OpenSQL('SELECT SUM(PAN_SALAIREANIM) AS MONTANT FROM SESSIONANIMAT WHERE '+
        'PAN_CODESTAGE="'+GetField('PAN_CODESTAGE')+'" AND PAN_MILLESIME="'+GetField('PAN_MILLESIME')+'"'+
        ' AND PAN_ORDRE='+IntToStr(GetField('PAN_ORDRE'))+'',True);  //DB2
    Montant := 0;
    if not QF.eof then Montant := QF.FindField('MONTANT').AsFloat;
    Ferme(QF);
    TS := TobSessions.FindFirst([''],[''],False);
    If TS <> Nil then
    begin
            TS.PutValue('PSS_COUTSALAIR',Montant);
            TS.UpdateDB(False);
    end;
    TobSessions.Free;
 	If isFieldModified( 'PAN_SALAIREANIM') then CalcCtInvestSession('PEDAGOGIQUE',GetField('PAN_CODESTAGE'),GetField('PAN_MILLESIME'),GetField('PAN_ORDRE'));
end;

procedure TOM_SESSIONANIMAT.SalarieElipsisClick (Sender : TObject);        //PT1
var StDate:String;
    StWhere,StFrom : String;
begin
	If (PGBundleHierarchie) and (PGDroitMultiForm) then //PT5
	begin
		If IsValidDate(GetField('PAN_DATEDEBUT')) then StWhere :=  '(PSI_DATESORTIE IS NULL OR PSI_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR PSI_DATESORTIE >="'+UsDateTime(GetField('PAN_DATEDEBUT'))+'")'
		Else StWhere := '';
		StFrom := 'INTERIMAIRES LEFT JOIN DOSSIER ON PSI_NODOSSIER=DOS_NODOSSIER';
		StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT16
		StWhere := StWhere + ' AND PSI_NODOSSIER IN (SELECT DOS_NODOSSIER FROM DOSSIER WHERE '+GererCritereGroupeConfTous+')'; //PT5
		LookupList(THEdit(Sender), 'Liste des salariés',StFrom,'PSI_INTERIMAIRE', 'PSI_LIBELLE,PSI_PRENOM,PSI_ETABLISSEMENT,DOS_LIBELLE', StWhere, 'PSI_INTERIMAIRE', TRUE, -1);
	end
	else
	begin
		If IsValidDate(GetField('PAN_DATEDEBUT')) then StDate :=  '(PSA_DATESORTIE="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>="'+UsDateTime(GetField('PAN_DATEDEBUT'))+'")'
		Else StDate := '';
		StDate := RecupClauseHabilitationLookupList(StDate);  //PT2
		{$IFNDEF EAGLCLIENT}                                                    //PT4
		LookupList(THEdit(Sender),'Liste des salariés','SALARIES LEFT JOIN ETABLISS ON PSA_ETABLISSEMENT=ET_ETABLISSEMENT','PSA_SALARIE','PSA_LIBELLE,PSA_PRENOM,ET_LIBELLE',StDate,'PSA_LIBELLE', True,-1);
		{$ELSE}
		LookupList(THEdit(Sender),'Liste des salariés','SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE','PSA_SALARIE','PSA_LIBELLE,PSA_PRENOM',StDate,'PSA_LIBELLE', True,-1);
		{$ENDIF}
	end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 18/04/2007 / PT4
Modifié le ... :   /  /    
Description .. : Met à jour les coûts et le nombre d'heures de l'animateur
Mots clefs ... : ANIMATEUR;COUT;
*****************************************************************}
Procedure TOM_SESSIONANIMAT.MAJCouts () ;
Var
    Q : TQuery;
    ExerForm,LibEmploi    : String;
    Salaire : Double;
    Mess : String;
    DateExercice : TDateTime;
    TauxChargeNC,TauxChargeC : Double;
Begin
    Mess      := '';
    Salaire   := 0;
    LeSalarie := GetField('PAN_SALARIE');
    LibEmploi := '';
    ExerForm  := '';
    TauxChargeNC := 1;
    TauxChargeC := 1;
    DateExercice := 0;
    // Récupération du millesime
    Q := OpenSQL('SELECT PFE_MILLESIME,PFE_DATEDEBUT,PFE_TAUXCHARGENC,PFE_TAUXCHARGEC FROM EXERFORMATION WHERE PFE_DATEDEBUT<="'+UsDateTime(GetField('PAN_DATEFIN'))+'" '+ //PT3
                 'AND PFE_DATEFIN>="'+UsDateTime(GetField('PAN_DATEFIN'))+'"',True);
    If Not Q.eof then
    begin
      ExerForm:=Q.FindField('PFE_MILLESIME').AsString;
      TauxChargeC := Q.FindField('PFE_TAUXCHARGEC').AsFloat;
      // Récupération de la date de début de l'exercice afin de récupérer le taux horaire en vigueur
      DateExercice := Q.FindField('PFE_DATEDEBUT').AsDateTime;
      TauxChargeNC := Q.FindField('PFE_TAUXCHARGENC').AsFloat;
    end
    Else Mess := Mess+'#13#10- Aucun exercice de formation n''existe pour les dates saisies';
    Ferme(Q);

    If (ExerForm <> '') Then
    Begin
      if VH_Paie.PGForValoSalaire='VCR' then
      Begin
            Salaire := ForTauxHoraireReel(GetField('PAN_SALARIE'),0,0,'',ValReel,DateExercice);
            If GetField('PAN_SALARIE') <> '' then
            begin
              Q := OpenSQL('SELECT PSA_DADSCAT FROM SALARIES WHERE PSA_SALARIE="'+GetField('PAN_SALARIE')+'"',True);  //PT13
              If (Q.FindField('PSA_DADSCAT').AsString = '01') or (Q.FindField('PSA_DADSCAT').AsString = '02') then Salaire := Arrondi(Salaire * TauxChargeC,2)
              else Salaire := Arrondi(Salaire * TauxChargeNC,2);
              Ferme(Q);
            end;
     End
     Else If VH_Paie.PGForValoSalaire='VCC' then
     begin
            // Récupération du libellé emploi
            Q := OpenSQL('SELECT PSA_LIBELLEEMPLOI FROM SALARIES WHERE PSA_SALARIE="'+GetField('PAN_SALARIE')+'"',True);
            If ((not Q.eof) And (Q.FindField('PSA_LIBELLEEMPLOI').AsString <> '')) then
                LibEmploi:=Q.FindField('PSA_LIBELLEEMPLOI').AsString
            Else
                Mess := Mess+'#13#10- Le libellé emploi du salarié formateur n''est pas renseigné dans la fiche salarié';
            Ferme(Q);

            If (LibEmploi <> '') then Salaire := ForTauxHoraireCategoriel(LibEmploi,ExerForm);
     end;
     SetField('PAN_SALAIREANIM',Salaire*GetField('PAN_NBREHEURE'));
    End;

    If (Mess <> '') Then PGIBox('Calcul du salaire impossible car ' + Mess,Ecran.Caption);
End;

Initialization
registerclasses([TOM_SESSIONANIMAT]) ;
end.
