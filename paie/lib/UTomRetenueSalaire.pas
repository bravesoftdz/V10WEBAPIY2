{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 03/06/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : RETENUESALAIRE (RETENUESALAIRE)
Mots clefs ... : TOM;RETENUESALAIRE
*****************************************************************}
{
PT1 17/12/2002 PH V591 Toutes les dates sont initialisées à idate1900 ou 2099 au lieu de null
PT2 29/07/2004 JL V_50 Modification du recalcul de l'échéancier, on garde les paiements effectués
PT3 01/03/2005 JL V_60 FQ 12042 Griser bouton echenacier et historique en création
PT4 20/07/2005 JL V_60 FQ 12448 Contrôle pour salarié confidentiel
---- JL 20/03/2006 modification clé annuaire ----
PT5 11/04/2006 JL V_65 FQ 13044 Modif accès annuaire pour bénéficiaire
PT6 14/06/2006 PH V_65 FQ 13276 Erreur CWAS - Nom des champs incorrects
PT7 22/06/2006 PH V_65 requête tenant tenant du contenu de la clause pour la confidentialité
PT8 17/10/2006 JL V_70 : FQ 13044 Rafraichissement tablette annuaire pour bénéficiaire saisie arrêt
PT9 21/05/2007 V_720 JL Gestion accès depuis fiche salarié
PT10 16/07/2007 V_720 JL FQ 14541 Gestion elipsis salarié
PT11 26/12/2007 V_810 FC Concept accessibilité fiche salarié
}
Unit UTomRetenueSalaire;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fiche,FichList,FE_Main,HDB,
{$ELSE}
     eFiche,eFichList,MaineAgl,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOM,UTob,HTB97,Entpaie,PgOutils,PgOutils2,LookUp ;

Type
  TOM_RETENUESALAIRE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    private
    InitDateDeb,InitDateFin : TDateTime;
    InitMois : Integer;
    InitMtMens,InitMtTot : Double;
    EcheancierOK,InitEcheancier : Boolean;
    LeSalarie : String;
    TypeAction : String;
    procedure AccesHistorique (Sender : Tobject);
    procedure BtnAnnuClick (Sender : TObject);
    procedure BEcheancierClick (Sender : TObject);
    procedure GenererEcheancier;
    procedure ExitEdit(Sender: TObject);
    procedure SalarieElipsisClick(Sender : TObject);
    procedure ClickMainLevee(Sender : TObject);
    end ;

Implementation

procedure TOM_RETENUESALAIRE.OnNewRecord ;
begin
Inherited ;
        If LeSalarie <> '' then
        begin
          SetField('PRE_SALARIE',LeSalarie); //PT9
          SetControlEnabled('PRE_SALARIE',false); //PT11
        end;
        // PT1 17/12/2002 PH V591 Toutes les dates sont initialisées à idate1900 ou 2099 au lieu de null
        SetField('PRE_DATEDEBUT',DebutDeMois(Date));
        SetField('PRE_DATEFIN',FinDeMois(Date));
        SetField('PRE_NIVEAURS',0);
        SetField('PRE_ACTIF','X');
        Setfield('PRE_BENEFICIAIRERS',-1);
        setField('PRE_NBMOIS',1);
        SetControlEnabled('BECHEANCIER',False);
        SetControlEnabled('BHISTO',False); //PT3
end ;

procedure TOM_RETENUESALAIRE.OnDeleteRecord ;
begin
  Inherited ;
        If existeSQL('SELECT PHR_SALARIE FROM HISTORETENUE WHERE PHR_SALARIE="'+GetField('PRE_SALARIE')+'"'+
        ' AND PHR_ORDRE='+IntToStr(GetField('PRE_ORDRE'))+'') Then      // DB2
        begin
                PgiBox('Attention, il existe des mouvements dans l''historique pour cette retenue ils vont être suppirmés',TFFiche(Ecran).Caption);
                ExecuteSQL('DELETE FROM HISTORETENUE WHERE PHR_SALARIE="'+GetField('PRE_SALARIE')+'"'+
                ' AND PHR_ORDRE='+IntToStr(GetField('PRE_ORDRE'))+''); // DB2
        end;
end ;

procedure TOM_RETENUESALAIRE.OnUpdateRecord ;
var MessOblig,StSQL,MessModif : String;
begin
  Inherited ;
        MessOblig := '';
        If GetField('PRE_RETENUESAL') = '' then
        begin
                MessOblig := MessOblig + '#13#10- le type de retenue';
        end;
        If GetField('PRE_DATEDEBUT') = IDate1900 then
        begin
                MessOblig := MessOblig + '#13#10- la date de début';
        end;
        If GetField('PRE_REMBMAX') <> 'X' then
        begin
                If GetField('PRE_DATEFIN') = IDate1900 then
                begin
                        MessOblig := MessOblig + '#13#10- la date de fin';
                end;
                If GetField('PRE_MONTANTMENS') = 0 then
                begin
                        MessOblig := MessOblig + '#13#10- le montant mensuel';
                end;
                If GetField('PRE_DATEFIN') < GetField('PRE_DATEDEBUT') then
                begin
                        PGIBox ('La date de fin ne peut être antérieure à la date de début',Ecran.Caption);
                        LastError := 1;
                        Exit;
                end;
        end;
        If GetField('PRE_MONTANTTOT') = 0 then
        begin
                MessOblig := MessOblig + '#13#10- le montant total';
        end;
        If MessOblig <> '' then
        begin
                PGIBox ('Vous devez renseigné :'+MessOblig,Ecran.Caption);
                LastError := 1;
                Exit;
        end;
        StSQL := 'SELECT PTR_RETENUESAL FROM TYPERETENUE WHERE PTR_RETENUESAL="'+GetField('PRE_RETENUESAL')+'"';
        If Not ExisteSQL (StSQL) then
        begin
                PGIBox ('Ce type de retenue n''est pas paramétré',Ecran.Caption);
                LastError := 1;
                Exit;
        end;
        If (GetField('PRE_ECHEANCIER') = 'X') and (DS.State = DsInsert) then GenererEcheancier
        else if (GetField('PRE_ECHEANCIER') = 'X') and (InitEcheancier = False) then GenererEcheancier
        else If (GetField('PRE_ECHEANCIER') = 'X') and Not (DS.State = DsInsert) then
        begin
                MessModif := '';
                If IsFieldModified('PRE_DATEDEBUT') then MessModif := MessModif + '#13#10- la date de début';
                If IsFieldModified('PRE_DATEDEBUT') then MessModif := MessModif + '#13#10- la date de fin';
                If IsFieldModified('PRE_NBMOIS') then MessModif := MessModif + '#13#10- le nombre de mois';
                If IsFieldModified('PRE_MONTANTMENS') then MessModif := MessModif + '#13#10- le montant mensuel';
                If MessModif <> '' then
                begin
                        PGIBox('Vous avez modifié :'+MessModif+'#13#10 l''échéancier va être mis à jour',Ecran.Caption);
                        GenererEcheancier;
                end;
        end;
        If GetField('PRE_ECHEANCIER') = 'X' then SetControlEnabled('BECHEANCIER',True);
        SetControlEnabled('BHISTO',True);      //PT3
end ;

procedure TOM_RETENUESALAIRE.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_RETENUESALAIRE.OnLoadRecord ;
begin
Inherited ;
        InitDateDeb := GetField('PRE_DATEDEBUT');
        InitDateFin := GetField('PRE_DATEFIN');
        InitMois := GetField('PRE_NBMOIS');
        InitMtMens := GetField('PRE_MONTANTMENS');
        InitMtTot := GetField('PRE_MONTANTTOT');
        If GetField('PRE_ECHEANCIER') = 'X' then InitEcheancier := True
        else InitEcheancier := False;

  //DEB PT11
  if TypeAction='ACTION=CONSULTATION' then
  begin
    SetControlEnabled('PRE_REMBMAX',false);
    SetControlEnabled('PRE_ECHEANCIER',false);
    SetControlEnabled('BHISTO',false);
    SetControlEnabled('BECHEANCIER',false);
    SetControlEnabled('BMAINLEVEE',false);
    SetControlEnabled('BANNU',false);
  end;
  //FIN PT11
end ;

procedure TOM_RETENUESALAIRE.OnChangeField ( F: TField ) ;
var Q : TQuery;
    NbMoisPaye,MaxOrdre : Integer;
    PremMois,PremAnnee,NMois : Word;
    Montant : Double;
    st : String;
begin
  Inherited ;
  if TypeAction<>'ACTION=CONSULTATION' then //PT11
  begin
        If F.FieldName = 'PRE_SALARIE' Then
        begin
                If (TFFiche(Ecran).FTypeAction = TaCreat) and (GetField('PRE_SALARIE') <> '') Then
                begin
                       ExitEdit(THEdit(GetControl('PRE_SALARIE')));
                        // DEBUT PT4
                        // DEB PT7
                        St := SQLConf('SALARIES');
                        if st <> '' then st := ' AND '+ st;
                        If not ExisteSQL('SELECT PSA_SALARIE FROM SALARIES WHERE PSA_SALARIE="'+GetField('PRE_SALARIE')+'"'+st) then
                        begin  // FIN PT7
                             iF GetField('PRE_SALARIE') <> '' then
                             begin
                                  PGIBox('Le matricule salarié que vous avez saisi n''existe pas',Ecran.Caption);
                                  SetFocusControl('PRE_SALARIE');
                                  SetField('PRE_SALARIE','');
                             end;
                        end;
                        // FIN PT4
                        Q:=OpenSQL('SELECT MAX (PRE_ORDRE) AS MAXORDRE FROM RETENUESALAIRE WHERE PRE_SALARIE="'+GetField('PRE_SALARIE')+'"',True);
                        if NOT Q.EOF then MaxOrdre := Q.FindField('MAXORDRE').AsInteger else MaxOrdre := 0; // PORTAGECWAS
                        Ferme(Q);
                        MaxOrdre := MaxOrdre+1;
                        SetField('PRE_ORDRE',MaxOrdre);
                end;
        end;
        If F.FieldName = 'PRE_DATEDEBUT' Then
        begin
                If GetField('PRE_REMBMAX') <> 'X' then
                begin
                        If (GetField('PRE_DATEFIN') > IDate1900) and (GetField('PRE_DATEDEBUT') <> InitdateDeb ) Then
                        begin
                                NombreMois (GetField('PRE_DATEDEBUT'), GetField('PRE_DATEFIN'), PremMois, PremAnnee, NMois);
                                InitMois := NMois;
                                SetField('PRE_NBMOIS',InitMois);
                                InitMtTot := InitMois * InitMtMens;
                                SetField('PRE_MONTANTTOT',initMtTot);
                        end;
                end;
        end;
        If F.FieldName = 'PRE_DATEFIN' Then
        begin
                If GetField('PRE_REMBMAX') <> 'X' then
                begin
                        If (GetField('PRE_DATEFIN') > IDate1900) and (GetField('PRE_DATEFIN') <> InitDateFin ) Then
                        begin
                                NombreMois (GetField('PRE_DATEDEBUT'), GetField('PRE_DATEFIN'), PremMois, PremAnnee, NMois);
                                InitMois := NMois;
                                SetField('PRE_NBMOIS',InitMois);
                                InitMtTot := InitMois * InitMtMens;
                                SetField('PRE_MONTANTTOT',initMtTot);
                        end;
                end;
        end;
        If F.FieldName = 'PRE_NBMOIS' Then
        begin
                Q := OpenSQL('SELECT COUNT(PHR_SALARIE) NBVERSEMENT FROM HISTORETENUE WHERE '+
                'PHR_SALARIE="'+GetField('PRE_SALARIE')+
                '" AND PHR_ORDRE='+IntToStr(GetField('PRE_ORDRE'))+' AND PHR_EFFECTUE="X"',True);
                If Not Q.Eof then NbMoisPaye := Q.FindField('NBVERSEMENT').AsInteger
                else NbMoisPaye := 0;
                If (GetField('PRE_NBMOIS') < NbMoisPaye) and (NbMoisPaye > 0) then
                begin
                        PGIBox('Le nombre de mois doit être supérieur à '+IntToStr(NbMoisPaye)+' qui est le nombre de versement déja effectué',Ecran.Caption);
                        SetFocusControl('PRE_NBMOIS');
                        SetField('PRE_NBMOIS',NbMoisPaye + 1);
                        Exit;
                end;
                If GetField('PRE_REMBMAX') <> 'X' then
                begin
                        If GetField('PRE_NBMOIS') <> InitMois then
                        begin
                                NMois := GetField('PRE_NBMOIS');
                                InitDateFin := FinDeMois(PlusMois(GetField('PRE_DATEDEBUT'),NMois-1));
                                SetField('PRE_DATEFIN',InitDateFin);
                                InitMois := NMois;
                                Montant := GetField('PRE_MONTANTMENS');
                                InitMtTot := Montant * NMois;
                                SetField('PRE_MONTANTTOT',InitMtTot);
                        end;
                end;
        end;
        If F.FieldName = 'PRE_MONTANTMENS' Then
        begin
                If GetField('PRE_REMBMAX') <> 'X' then
                begin
                        if getField('PRE_MONTANTMENS') <> InitMtMens then
                        begin
                                InitMtMens := GetField('PRE_MONTANTMENS');
                                NMois := getField('PRE_NBMOIS');
                                InitMtTot := InitMtMens * NMois;
                                SetField('PRE_MONTANTTOT',InitMtTot);
                        end;
                end;
        end;
        If F.FieldName = 'PRE_MONTANTTOT' Then
        begin
                If GetField('PRE_REMBMAX') <> 'X' then
                begin
                        if GetField('PRE_MONTANTTOT') <> InitMtTot then
                        begin
                                InitMtTot := GetField('PRE_MONTANTTOT');
                                NMois := getField('PRE_NBMOIS');
                                InitMtMens := Arrondi(InitMtTot / NMois,2);
                                SetField('PRE_MONTANTMENS',InitMtMens);
                        end;
                end;
        end;
        If F.FieldName = 'PRE_REMBMAX' Then
        begin
                If GetField('PRE_REMBMAX') = 'X' then
                begin
                        If isFieldModified('PRE_REMBMAX') then
                        begin
                                SetField('PRE_NBMOIS',0); // PT6
                                SetField('PRE_DATEFIN',IDate1900);
                                SetField('PRE_MONTANTMENS',0);// PT6
                        end;
                        SetControlEnabled('PRE_DATEFIN',False);
                        SetControlEnabled('PRE_MONTANTMENS',False);
                        SetControlEnabled('PRE_NBMOIS',False);
                        If GetField('PRE_ECHEANCIER') = 'X' then SetField('PRE_ECHEANCIER','-');
                        SetControlEnabled('PRE_ECHEANCIER',False);
                end
                else
                begin
                        SetControlEnabled('PRE_DATEFIN',True);
                        SetControlEnabled('PRE_MONTANTMENS',True);
                        SetControlEnabled('PRE_NBMOIS',True);
                        SetControlEnabled('PRE_ECHEANCIER',True);
                end;
        end;
        If F.FieldName = 'PRE_ECHEANCIER' then
        begin
                If GetField('PRE_ECHEANCIER') = 'X' then
                begin
                        If DS.State <> DsInsert then SetControlEnabled('BECHEANCIER',True); //PT3
                        If GetField ('PRE_REMBMAX') = 'X' then SetField('PRE_REMBMAX','-');
                        SetControlEnabled('PRE_REMBMAX',False);
                end
                Else
                begin
                        SetControlEnabled('BECHEANCIER',False);   //False
                        SetControlEnabled('PRE_REMBMAX',True);
                end;
        end;
  end; //PT11
end ;

procedure TOM_RETENUESALAIRE.OnArgument ( S: String ) ;
var BHisto,BAnnu,BEcheancier,Bt : TToolBarButton97;
    {$IFNDEF EAGLCLIENT}
    Sal : THDBEdit;
    {$ELSE}
    Sal : THEdit;
    {$ENDIF}
begin
Inherited ;
        TypeAction := ReadTokenPipe(S,';');
        LeSalarie := ReadTokenPipe(S,';');           //PT9
        BHisto := TToolBarButton97(GetControl('BHISTO'));
        If BHisto <> Nil then BHisto.OnClick := AccesHistorique;
        BAnnu := TToolBarButton97(GetControl('BANNU'));
        If BAnnu <> Nil then BAnnu.OnClick := BtnAnnuClick;
        BEcheancier := TToolBarButton97(GetControl('BECHEANCIER'));
        If BEcheancier <> Nil then BEcheancier.OnClick := BEcheancierClick;
        {$IFNDEF EAGLCLIENT}
        sal:=THDBEdit(GetControl('PRE_SALARIE')) ;
        {$ELSE}
        sal:=THEdit(GetControl('PRE_SALARIE')) ;
        {$ENDIF}
         If sal <> Nil then sal.OnElipsisClick := SalarieElipsisClick; // PT10
//        if sal<>nil then sal.OnExit:=ExitEdit;
        SetControlVisible('BGENEREECHEANCIER',False);
        Bt := TToolBarButton97(GetControl('BMAINLEVEE'));
        If Bt <> Nil then Bt.OnClick := ClickMainLevee;
end ;


procedure TOM_RETENUESALAIRE.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_RETENUESALAIRE.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_RETENUESALAIRE.AccesHistorique (Sender : TObject);
var St : String;
begin
        St := GetField('PRE_SALARIE') + ';' + IntToStr(GetField('PRE_ORDRE'));
        AglLanceFiche('PAY','HISTORETENUE_MUL','','',St);
end;

procedure TOM_RETENUESALAIRE.BtnAnnuClick(Sender:TObject);
var Ret : String;
    {$IFNDEF EAGLCLIENT}
    Combo : THDBValComboBox;
    {$ELSE}
    Combo : THValComboBox;
    {$ENDIF}
begin
     //DEBUT PT5
     Ret := AGLLanceFiche('PAY','PGANNUAIRE','','','COMBO;RET');   //PT8
     If Ret <> 'VIDE' then
     begin
          {$IFNDEF EAGLCLIENT}
          Combo := THDBValComboBox(GetControl('PRE_BENEFRSGU'));
          {$ELSE}
          Combo := THValComboBox(GetControl('PRE_BENEFRSGU'));
          {$ENDIF}
          Combo.ReLoad;
          ForceUpdate;
          SetField('PRE_BENEFRSGU',Ret);
          SetControlText('PRE_BENEFRSGU',Ret);
     end;
     //FIN PT5
end;

procedure TOM_RETENUESALAIRE.BEcheancierClick (Sender : TObject);
var St : String;
    Q : TQuery;
    MtTotH,MtTotR : Double;
    Ok : Boolean;
    rep : Integer;
begin
        St := GetField('PRE_SALARIE') + ';' + IntToStr(GetField('PRE_ORDRE'));
        Ok := False;
        While Ok = False do
        begin
                MtTotR := GetField('PRE_MONTANTTOT');
                AglLanceFiche('PAY','ECHEANCIERRETENUE',St,'',St);
                Q := OpenSQL('SELECT SUM(PHR_MONTANTMENS) TOTAL '+
                'FROM HISTORETENUE '+
                'WHERE PHR_ORDRE='+IntToStr(GetField('PRE_ORDRE'))+' '+  // DB2
                'AND PHR_SALARIE="'+GetField('PRE_SALARIE')+'"',True);
                If Not Q.Eof then MtTotH := Q.FindField('TOTAL').AsFloat
                Else MtTotH := 0;
                If (MtTotH - MtTotR > 1) and (MtTotH - MtTotR <1) then
                begin
                        Rep := PGIAskCancel('Attention, la somme des mensualités de l''échéancier est égale à '+FloatToStr(MtTotH)+
                        '#13#10 Vous devez le modifier pour que la somme soit égale au montant total de la retenue : '+FloatToStr(MtTotR)+'.'+
                        '#13#10 Voulez-vous retourner modifier l''échéancier (si non il sera recalculé en fonction du montant total) ?',Ecran.Caption);
                        If Rep = MrNo then
                        begin
                                GenererEcheancier;
                                OK := True;
                        end;
                end
                else OK := True;
        end;
end;

procedure TOM_RETENUESALAIRE.GenererEcheancier;
var DateDebut,DateFinPaye : TDateTime;
    NbMois,NbMoisPaye,i : Integer;
    MtMensuel,MtTotal,MtPaye : Double;
    TobHisto : Tob;
    Q : TQuery;
begin
//        If ExisteSQL('SELECT PHR_SALARIE FROM HISTORETENUE WHERE '+
//        'PHR_SALARIE="'+GetField('PRE_SALARIE')+'" AND PHR_ORDRE="'+IntToStr(GetField('PRE_ORDRE'))+'"') then
//        begin
//                If PGIAsk('Attention il existe déja des enregistrements, ils seront effacer si l''échéancier est recalculé, #13#10voulez vous continuer',Ecran.Caption) <> MrYes then
//                begin
//                Exit;
//                end;
//        end;


        Q := OpenSQL('SELECT SUM(PHR_MONTANTMENS) SOMME FROM HISTORETENUE WHERE '+
        'PHR_SALARIE="'+GetField('PRE_SALARIE')+
        '" AND PHR_ORDRE='+IntToStr(GetField('PRE_ORDRE'))+' AND PHR_EFFECTUE="X"',True);
        If Not Q.Eof then MtPaye := Q.FindField('SOMME').AsFloat
        else MtPaye := 0;
        Q := OpenSQL('SELECT MAX(PHR_DATEFIN) DATEFIN FROM HISTORETENUE WHERE '+
        'PHR_SALARIE="'+GetField('PRE_SALARIE')+
        '" AND PHR_ORDRE='+IntToStr(GetField('PRE_ORDRE'))+' AND PHR_EFFECTUE="X"',True);
        If Not Q.Eof then DateFinPaye := Q.FindField('DATEFIN').AsDateTime
        else DateFinPaye := IDate1900;
        Q := OpenSQL('SELECT COUNT(PHR_SALARIE) NBVERSEMENT FROM HISTORETENUE WHERE '+
        'PHR_SALARIE="'+GetField('PRE_SALARIE')+
        '" AND PHR_ORDRE='+IntToStr(GetField('PRE_ORDRE'))+' AND PHR_EFFECTUE="X"',True);
        If Not Q.Eof then NbMoisPaye := Q.FindField('NBVERSEMENT').AsInteger
        else NbMoisPaye := 0;

        ExecuteSQL('DELETE FROM HISTORETENUE WHERE '+
        'PHR_SALARIE="'+GetField('PRE_SALARIE')+
        '" AND PHR_ORDRE='+IntToStr(GetField('PRE_ORDRE'))+' AND PHR_EFFECTUE="-"');  //PT2
        DateDebut := DebutdeMois(GetField('PRE_DATEDEBUT'));
        NbMois := GetField('PRE_NBMOIS');
        MtMensuel := GetField('PRE_MONTANTMENS');
        If NbMoisPaye > 0 then
        begin
                MtTotal := GetField('PRE_MONTANTTOT');
                MtTotal := MtTotal - MtPaye;
                NbMois := NbMois - NbMoisPaye;
                MtMensuel := MtTotal / NbMois;
                DateDebut := PlusMois(DateFinPaye,1);
                DateDebut := DebutDeMois(DateDebut);
        end;
        For i := 1 to NbMois do
        begin
                TobHisto := Tob.Create('HISTORETENUE',Nil,-1);
                TobHisto.PutValue('PHR_SALARIE',GetField('PRE_SALARIE'));
                TobHisto.PutValue('PHR_ORDRE',GetField('PRE_ORDRE'));
                TobHisto.PutValue('PHR_DATEDEBUT',DateDebut);
                TobHisto.PutValue('PHR_DATEFIN',FinDeMois(DateDebut));
                TobHisto.PutValue('PHR_MONTANT',0);
                TobHisto.PutValue('PHR_MONTANTMENS',MtMensuel);
                TobHisto.PutValue('PHR_REPRISEARR',0);
                TobHisto.PutValue('PHR_ARRIERE',0);
                TobHisto.PutValue('PHR_DATEPAIEMENT',IDate1900);
                TobHisto.PutValue('PHR_CUMULARRIERE',0);
                TobHisto.PutValue('PHR_CUMULVERSE',0);
                TobHisto.PutValue('PHR_NBJOURS',0);
                TobHisto.PutValue('PHR_RETENUESAL',GetField('PRE_RETENUESAL'));
                TobHisto.PutValue('PHR_EFFECTUE','-');
                TobHisto.InsertDB(Nil);
                TobHisto.Free;
                DateDebut := PlusMois(DateDebut,1);
        end;
end;

procedure TOM_RETENUESALAIRE.ExitEdit(Sender: TObject);
var edit : thedit;
    Salarie:String;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<10) and (isnumeric(edit.text)) then
       begin
       Salarie:=AffectDefautCode(edit,10);
       edit.text:=Salarie;
       SetField('PRE_SALARIE',Salarie);
       end;
end;

procedure TOM_RETENUESALAIRE.SalarieElipsisClick(Sender : TObject); //PT10
begin
  {$IFNDEF EAGLCLIENT}
  LookupList(THDBEdit(Sender), 'Liste des salariés','SALARIES','PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', '', 'PSA_SALARIE', TRUE, -1);
  {$ELSE}
  LookupList(THEdit(Sender), 'Liste des salariés','SALARIES','PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', '', 'PSA_SALARIE', TRUE, -1);
  {$ENDIF}
end;

procedure TOM_RETENUESALAIRE.ClickMainLevee(Sender : TObject);
var Sal,NumOrdre : String;
begin
  Sal := GetField('PRE_SALARIE');
  NumOrdre := IntToStr(GetField('PRE_ORDRE'));
  AGLLanceFiche('PAY','RETENUEMAINLEVEE','','',Sal+';'+NumOrdre+';ACTION=CREATION');
end;


Initialization
  registerclasses ( [ TOM_RETENUESALAIRE ] ) ;
end.




