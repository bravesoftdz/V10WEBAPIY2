{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 02/09/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGSAISIEFRAISSALFOR ()
Mots clefs ... : TOF;PGSAISIEFRAISSALFOR
*****************************************************************
PT1  | 08/04/2003 | V_42  | JL | Correction pour saisie montant avec décimales
PT2  | 03/11/2003 | V_50  | JL | Sasie des frais par salarié au lieu de la saisie par session pour CEGID
PT3  | 10/12/2004 | V_60  | JL | Mise à jour des plafonds sur changement quantité
PT4  | 01/03/2005 | V_60  | JL | FQ 12038 correction message + erreur lors de navigation
---  | 20/03/2006 |       | JL | Modification clé annuaire ----
PT5  | 17/05/2007 | V_720 | FL | FQ 11532 Gestion des plafonds de frais par population
PT6  | 07/06/2007 | V8720 | FL | Gestion des multi-sessions + types de plans
PT7  | 19/02/2008 | V_803 | FL | Correction des critères Predefini/NoDossier pour la gestion multidossier
PT8  | 02/04/2008 | V_803 | FL | Adaptation partage formation
}
Unit UTofPgSaisieFraisSalFor;

Interface

uses     UTOF,
{$IFNDEF EAGLCLIENT}
         DBCtrls,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
         MaineAGL,UtileAgl,
{$ENDIF}
         Hctrls,HEnt1,HMsgBox,Classes,UTOB,
         sysutils,HTB97,Vierge,EntPaie,
         Controls,Spin,PgOutilsFormation,PGPopulOutils,PGOutils2;

//Uses StdCtrls,Controls,Classes,  FE_Main,
{$IFNDEF EAGLCLIENT}
//     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
//     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,Spin,UTob,EntPaie,Vierge,HTB97,FE_Main,LookUp ;

Type
  TOF_PGSAISIEFRAISSALFOR = Class (TOF)
    procedure OnClose                  ; override ;
    procedure OnArgument (S : String ) ; override ;
    Private
    QMul:TQuery;
    LeMillesime,LaSession,LeStage,LeSalarie,MillesimeEC,OrgCollect,LieuFormation,TypeSaisie,LePlan,NbHeuresStag,NbHeuresAnim : String;
    SalPrem, SalPrec, SalSuiv, SalDern : TToolBarButton97;
    TobLibFrais : TOB;
    AvecClient : Boolean;
    procedure AfficheFrais;
    procedure AfficheEntete;
    procedure SalPremClick (Sender : TObject);
    procedure SalPrecClick (Sender : TObject);
    procedure SalSuivClick (Sender : TObject);
    procedure SalDernClick (Sender : TObject);
    Function BougeSal(Button : TNavigateBtn) : boolean ;
    procedure GereQuerySal();
    Procedure MetABlanc;
    Procedure ExitMontant(Sender : TObject);
    Procedure ExitQuantite(Sender : TObject);  //PT3
    Procedure ExitFrais(Sender : TObject);
    procedure UpdateTable(Sender : TObject);
    Procedure RefreshSelonPlan (Sender : TObject); //PT6

  end ;

Implementation

Uses UtilPGI;

procedure TOF_PGSAISIEFRAISSALFOR.OnClose;
begin
Inherited ;
        FreeAndNil(TobLibFrais); //PT6
        Ferme(QMul);
end;

procedure TOF_PGSAISIEFRAISSALFOR.OnArgument (S : String ) ;
var i,NbFrais : Integer;
    Combo : THValComboBox;
    Edit : THEdit;
    Spin : THSpinEdit;
    Q : TQuery;
    Requete : String; //PT8
    Champ : String; //PT8
begin
  Inherited ;
        TFVierge(Ecran).BValider.OnClick := UpdateTable;
        TypeSaisie := ReadTokenPipe(S,';');
        LeSalarie := Trim(ReadTokenPipe(S,';'));
        LeStage := Trim(ReadTokenPipe(S,';'));
        LaSession := Trim(ReadTokenPipe(S,';'));
        LeMillesime := Trim(ReadTokenPipe(S,';'));
        MillesimeEC := Trim(ReadTokenPipe(S,';'));

        //PT6 - Début
        // Récupération des libellés de frais
        If PGBundleInscFormation Then
        	Requete := 'SELECT CC_CODE FROM '+GetBase(GetBasePartage(BUNDLE_FORMATION),'CHOIXCOD')+' WHERE CC_TYPE="PFA" ORDER BY CC_CODE' //PT8
        Else
        	Requete := 'SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="PFA" ORDER BY CC_CODE';
        Q := OpenSQL(Requete, True);
        TobLibFrais := Tob.Create('Les libellés',Nil,-1);
        TobLibFrais.LoadDetailDB('CHOIXCOD','','',Q,False);
        Ferme(Q);

        NbFrais := VH_Paie.PGFNbFraisLibre; //TobLibFrais.Detail.Count; //PT8
        //PT6 - Fin

        AfficheEntete; //PT6
        AfficheFrais;

        If TypeSaisie <> 'FRAISSALARIE' then      //DEBUT PT2
        begin
        	If PGBundleHierarchie Then //PT8
        	Begin
        		Requete := 'SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE '+
                'PSI_INTERIMAIRE IN (SELECT PAN_SALARIE FROM SESSIONANIMAT WHERE PAN_CODESTAGE="'+LeStage+'" '+
                'AND PAN_MILLESIME="'+LeMillesime+'" AND PAN_ORDRE='+LaSession+') OR '+  
                'PSI_INTERIMAIRE IN (SELECT PFO_SALARIE FROM FORMATIONS WHERE PFO_CODESTAGE="'+LeStage+'" '+
                'AND PFO_MILLESIME="'+LeMillesime+'" AND PFO_ORDRE='+LaSession+')';
                Champ := 'PSI_INTERIMAIRE';
            End
        	Else
        	Begin
        		Requete := 'SELECT PSA_SALARIE FROM SALARIES WHERE '+
                'PSA_SALARIE IN (SELECT PAN_SALARIE FROM SESSIONANIMAT WHERE PAN_CODESTAGE="'+LeStage+'" '+
                'AND PAN_MILLESIME="'+LeMillesime+'" AND PAN_ORDRE='+LaSession+') OR '+  //DB2
                'PSA_SALARIE IN (SELECT PFO_SALARIE FROM FORMATIONS WHERE PFO_CODESTAGE="'+LeStage+'" '+
                'AND PFO_MILLESIME="'+LeMillesime+'" AND PFO_ORDRE='+LaSession+')';  //DB2
                Champ := 'PSA_SALARIE';
            End;
                
                QMul := OpenSQL(Requete, True);
                If QMul.eof Then
                begin
                        PGIBox('Il n''y a pas de salarié',Ecran.Caption);
                        Exit;
                end;
                QMul.First;
                
                While QMul.FindField(Champ).AsString <> LeSalarie do //PT8
                begin
                        IF QMul.Eof Then Break;
                        QMul.Next;
                end;


                // Gestion du navigateur
                SalPrem  :=  TToolbarButton97(GetControl('BSALPREM'));
                if SalPrem  <>  NIL then
                begin
                        SalPrem.Enabled  :=  False;
                        SalPrem.Visible  :=  True;
                        SalPrem.OnClick  :=  SalPremClick;
                end;
                SalPrec  :=  TToolbarButton97(GetControl('BSALPREC'));
                if SalPrec  <>  NIL then
                begin
                        SalPrec.Enabled  :=  False;
                        SalPrec.Visible  :=  True;
                        SalPrec.OnClick  :=  SalPrecClick;
                end;
                SalSuiv  :=  TToolbarButton97(GetControl('BSALSUIV'));
                if SalSuiv  <>  NIL then
                begin
                        SalSuiv.Enabled  :=  True;
                        SalSuiv.Visible  :=  True;
                        SalSuiv.OnClick  :=  SalSuivClick;
                end;
                SalDern  :=  TToolbarButton97(GetControl('BSALDERN'));
                if SalDern  <>  NIL then
                begin
                        SalDern.Enabled  :=  True;
                        SalDern.Visible  :=  True;
                        SalDern.OnClick  :=  SalDernClick;
                end;
        end
        else
        begin
                setControlVisible('BSALPREM',False);
                setControlVisible('BSALPREC',False);
                setControlVisible('BSALSUIV',False);
                setControlVisible('BSALDERN',False);
        end;                                                //FIN PT2

        For i := 1 to NbFrais do //PT6
        begin
                Edit := THEdit(GetControl('MONTANT'+IntToStr(i)));
                Combo := THValComboBox(GetControl('FRAIS'+IntToStr(i)));
                Spin := THSpinEdit(GetControl('QTEFRAIS'+IntToStr(i)));
                If Edit <> Nil Then Edit.OnChange := ExitMontant;
                If Combo <> Nil Then Combo.onExit := ExitFrais;
                If Spin <> Nil Then Spin.OnChange := ExitQuantite; //PT3
                SetControlVisible('FRAIS'+IntToStr(i),True);
                SetControlVisible('QTEFRAIS'+IntToStr(i),True);
                SetControlVisible('MONTANT'+IntToStr(i),True);
                SetControlVisible('PLAFOND'+IntToStr(i),True);
        end;

        //PT6 - Début
        Combo := THValComboBox(GetControl('TYPEPLANPREV'));
        If (Combo <> Nil) Then Combo.OnChange := RefreshSelonPlan;
        //PT6 - Fin
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 08/06/2007 / PT6
Modifié le ... :   /  /    
Description .. : Affichage de l'en-tête de la fiche
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGSAISIEFRAISSALFOR.AfficheEntete;
var QF : TQuery;
begin
        MetABlanc;

        // Cas du stagiaire : Récupération des caractéristiques de la session
        QF := OpenSQL('SELECT PFO_NBREHEURE,PSS_DATEDEBUT,PSS_DATEFIN,PST_LIBELLE,PSS_LIBELLE,PSS_LIEUFORM,PSS_ORGCOLLECTSGU FROM FORMATIONS'+
        ' LEFT JOIN SESSIONSTAGE ON PFO_ORDRE=PSS_ORDRE AND PFO_MILLESIME=PSS_MILLESIME AND PFO_CODESTAGE=PSS_CODESTAGE'+
        ' LEFT JOIN STAGE ON PFO_MILLESIME=PST_MILLESIME AND PFO_CODESTAGE=PST_CODESTAGE'+
        ' WHERE PFO_ORDRE='+LaSession+' AND PFO_MILLESIME="'+LeMillesime+'" AND PFO_CODESTAGE="'+LeStage+'"',True);  //DB2
        if not QF.eof then
        begin
                NbHeuresStag := QF.FindField('PFO_NBREHEURE').AsString;
                SetControltext('STAGE',QF.FindField('PST_LIBELLE').AsString);
                SetControltext('SESSION',QF.FindField('PSS_LIBELLE').AsString);
                SetControltext('LIEUFORM',RechDom('PGLIEUFORMATION',QF.FindField('PSS_LIEUFORM').AsString,False));
                LieuFormation := QF.FindField('PSS_LIEUFORM').AsString;
                SetControltext('ORGCOLLECTEUR',RechDom('PGORGCOLLECTEUR',QF.FindField('PSS_ORGCOLLECTSGU').AsString,False));
                OrgCollect := QF.FindField('PSS_ORGCOLLECTSGU').AsString;
                SetControltext('DATEDEBUT',DateToStr(QF.FindField('PSS_DATEDEBUT').AsDateTime));
                SetControltext('DATEFIN',DateToStr(QF.FindField('PSS_DATEFIN').AsDateTime));
        end;
        Ferme(QF);

        // Cas de l'animateur : Récupération des caractéristiques de la session
        QF := OpenSQL('SELECT PAN_NBREHEURE,PSS_AVECCLIENT FROM SESSIONANIMAT'+
        ' LEFT JOIN SESSIONSTAGE ON PAN_ORDRE=PSS_ORDRE AND PAN_MILLESIME=PSS_MILLESIME AND PAN_CODESTAGE=PSS_CODESTAGE'+
        ' WHERE PAN_ORDRE='+LaSession+' AND PAN_MILLESIME="'+LeMillesime+'" AND PAN_CODESTAGE="'+LeStage+'"',True);  //DB2
        If Not QF.eof Then
        begin
                NbHeuresAnim := QF.FindField('PAN_NBREHEURE').AsString;
                If QF.FindField('PSS_AVECCLIENT').AsString='X' then AvecClient := True
                else Avecclient := False;
        end;
        Ferme(QF);
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : ??/??/???? / PT6
Modifié le ... :   /  /    
Description .. : Détermination du cas stagiaire/animateur, mise à jour du 
Suite ........ : titre de la fenêtre
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGSAISIEFRAISSALFOR.AfficheFrais;
var Q{,QS} : TQuery;
    Stagiaire : Boolean;
    Plus, Valeur: String;
begin
        // Mise à jour du titre de la fenêtre
        //PT8 - Début
        (*QS := OpenSQL('SELECT PSA_LIBELLE,PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE="'+LeSalarie+'"',True);
        If not QS.eof Then
        begin
                TFVierge(Ecran).Caption := 'Salarié : '+LeSalarie+' '+QS.FindField('PSA_LIBELLE').AsString+' '+QS.FindField('PSA_PRENOM').AsString;
                UpdateCaption(TFVierge(Ecran));
        end;
        Ferme(QS);*)
        TFVierge(Ecran).Caption := 'Salarié : '+LeSalarie+' ';
        If PGBundleHierarchie Then
			TFVierge(Ecran).Caption := TFVierge(Ecran).Caption + RechDom('PGSALARIEINT',LeSalarie,False)
		Else
			TFVierge(Ecran).Caption := TFVierge(Ecran).Caption + RechDom('PGSALARIE',LeSalarie,False);
        UpdateCaption(TFVierge(Ecran));
        //PT8 - Fin

        // Détermination s'il s'agit d'un stagiaire ou d'un animateur et recherche des types de plans associés
        Q := OpenSQL('SELECT PFO_SALARIE,PFO_TYPEPLANPREV FROM FORMATIONS '+
        'WHERE PFO_CODESTAGE="'+LeStage+'" AND PFO_ORDRE='+LaSession+' AND PFO_MILLESIME="'+LeMillesime+'" AND PFO_SALARIE="'+LeSalarie+'"',True);  //DB2
        Stagiaire := Not Q.EOF;
        If Not Q.EOF Then
        Begin
             Plus := 'AND CO_CODE IN (';
             Valeur := Q.FindField('PFO_TYPEPLANPREV').AsString;
             While Not Q.EOF Do
             Begin
                    Plus := Plus + '"' + Q.FindField('PFO_TYPEPLANPREV').AsString + '"';
                    Q.Next;
                    If Not Q.EOF Then Plus := Plus + ',';
             End;
             Plus := Plus + ')';
        End
        Else
             Plus := '';
        Ferme(Q);
        SetControlProperty ('TYPEPLANPREV','Plus',Plus);

        // Sélection par défaut du premier type de plan
        THValComboBox(GetControl('TYPEPLANPREV')).Value := Valeur;
        LePlan := Valeur;

        // Si une seule valeur de plan : on bloque la combo
        If (Valeur = '') Or (THValComboBox(GetControl('TYPEPLANPREV')).Values.Count = 1) Then
               THValComboBox(GetControl('TYPEPLANPREV')).Enabled := False
        Else
               THValComboBox(GetControl('TYPEPLANPREV')).Enabled := True;

        // Cas du stagiaire : Récupération des caractéristiques de la session
        If Stagiaire=True Then
        begin
                SetControlProperty('THQUALITE','Value','STA');
                SetControltext('NBHEURE',NbHeuresStag);
        end
        // Cas de l'animateur : Récupération des caractéristiques de la session
        Else
        begin
                SetControlProperty('THQUALITE','Value','FOR');
                SetControltext('NBHEURE',NbHeuresAnim);
                If AvecClient = True then PGIBox('Attention ce salarié est formateur et c''est une session avec client',Ecran.Caption);
        end;

        RefreshSelonPlan(Nil);
end;

procedure TOF_PGSAISIEFRAISSALFOR.SalPremClick(Sender : TObject);
begin
        BougeSal(nbFirst) ;
        SalPrem.Enabled  :=  FALSE;
        SalPrec.Enabled  :=  FALSE;
        SalSuiv.Enabled  :=  TRUE;
        SalDern.Enabled  :=  TRUE;
end;

procedure TOF_PGSAISIEFRAISSALFOR.SalPrecClick(Sender : TObject);
begin
        BougeSal(nbPrior) ;
        if QMul.BOF then
        begin
                SalPrem.Enabled  :=  FALSE;
                SalPrec.Enabled  :=  FALSE;
        end;
        SalSuiv.Enabled  :=  TRUE;
        SalDern.Enabled  :=  TRUE;
end;

procedure TOF_PGSAISIEFRAISSALFOR.SalSuivClick(Sender : TObject);
begin
        BougeSal(nbNext) ;
        SalPrem.Enabled  :=  TRUE;
        SalPrec.Enabled  :=  TRUE;
        if QMul.EOF then
        begin
                SalSuiv.Enabled  :=  FALSE;
                SalDern.Enabled  :=  FALSE;
        end;
end;

procedure TOF_PGSAISIEFRAISSALFOR.SalDernClick(Sender : TObject);
begin
        BougeSal(nbLast) ;
        SalPrem.Enabled  :=  TRUE;
        SalPrec.Enabled  :=  TRUE;
        SalSuiv.Enabled  :=  FALSE;
        SalDern.Enabled  :=  FALSE;
end;

Function TOF_PGSAISIEFRAISSALFOR.BougeSal(Button : TNavigateBtn) : boolean ;
BEGIN
        result := TRUE ;
        if Button=nbDelete then
        begin
                if QMul.EOF = FALSE then
                begin
                        QMul.Next;
                        if QMul.EOF = TRUE then
                        begin
                                QMul.prior ;
                                if QMul.BOF then Close;
                        end
                end
                else
                begin
                        if QMul.BOF = FALSE then
                        begin
                                QMul.prior;
                                if QMul.BOF = TRUE then Close;
                        end;
                end;
        end;
        if QMul  <>  NIL then
        begin
                Case Button of
                        nblast : QMul.Last;
                        nbfirst : QMul.First;
                        nbnext : QMul.Next;
                        nbprior : QMul.prior;
                end;
        end;
        GereQuerySal;
end ;

procedure TOF_PGSAISIEFRAISSALFOR.GereQuerySal();
begin
        if QMul = NIL then exit;
        If PGBundleHierarchie Then
        	LeSalarie  :=  QMul.FindField('PSI_INTERIMAIRE').AsString
        Else
        	LeSalarie  :=  QMul.FindField('PSA_SALARIE').AsString;
        AfficheFrais;
end;

Procedure TOF_PGSAISIEFRAISSALFOR.MetABlanc;
var i : Integer;
    Combo : THValComboBox;
begin
        For i := 1 to 15 do
        begin
                Combo := THValComboBox(GetControl('FRAIS'+IntToStr(i)));
                Combo.Value := '';
//                Quantite := TSpinEdit(GetControl('QTEFRAIS'+IntToStr(i)));    PT4 mis en commentaire
//                Quantite.Value := 1;  //PT4
//                SetControlText('MONTANT'+IntToStr(i),'');
//                SetControlText('PLAFOND'+IntToStr(i),'');
        end;
end;

Procedure TOF_PGSAISIEFRAISSALFOR.ExitMontant(Sender : TObject);
var Long,i : Integer;
    Plafond,Montant : Double;
    Combo : THValComboBox;
    Qte : TSpinEdit;
begin
        If GetControlText(THEdit(Sender).Name) <> '' Then
        begin
                i := 0;
                Long := length(THEdit(Sender).Name);
                If Long=8 Then i := StrToInt(THEdit(Sender).Name[8]);
                If Long=9 Then i := StrToInt(THEdit(Sender).Name[8]+THEdit(Sender).Name[9]);
                Combo := THValComboBox(GetControl('FRAIS'+IntToStr(i)));
                Qte := TSpinEdit(GetControl('QTEFRAIS'+IntToStr(i)));
                Montant := StrToFloat(GetControlText('MONTANT'+IntToStr(i)));
                Plafond := RendPlafondFraisForm(Combo.Value,LieuFormation,OrgCollect,MillesimeEC,DonnePopulation(LeSalarie)); //PT5
                Plafond := Plafond*Qte.value;
                If (Montant>Plafond) and (plafond>0) Then SetControlText('PLAFOND'+IntToStr(i),FloatToStr(Plafond))
                Else SetControlText('PLAFOND'+IntToStr(i),FloatToStr(Montant));
        end;
end;

Procedure TOF_PGSAISIEFRAISSALFOR.ExitQuantite(Sender : TObject); //PT3
var Long,i : Integer;
    Plafond,Montant : Double;
    Combo : THValComboBox;
    Qte : TSpinEdit;
begin
        If GetControlText(THEdit(Sender).Name) <> '' Then
        begin
                i := 0;
                Long := length(THSpinEdit(Sender).Name);
                If Long=9 Then i := StrToInt(THSpinEdit(Sender).Name[9]);
                If Long=10 Then i := StrToInt(THSpinEdit(Sender).Name[9]+THEdit(Sender).Name[10]);
                Combo := THValComboBox(GetControl('FRAIS'+IntToStr(i)));
                Qte := TSpinEdit(GetControl('QTEFRAIS'+IntToStr(i)));
                Montant := StrToFloat(GetControlText('MONTANT'+IntToStr(i)));
                Plafond := RendPlafondFraisForm(Combo.Value,LieuFormation,OrgCollect,MillesimeEC,DonnePopulation(LeSalarie)); //PT5
                Plafond := Plafond*Qte.value;
                If (Montant>Plafond) and (plafond>0) Then SetControlText('PLAFOND'+IntToStr(i),FloatToStr(Plafond))
                Else SetControlText('PLAFOND'+IntToStr(i),FloatToStr(Montant));
        end;
end;


Procedure TOF_PGSAISIEFRAISSALFOR.ExitFrais(Sender : TObject);
var Long,i,j : Integer;
    Qte : TSpinEdit;
    ValeurFrais,TestValeur : String;
    Combo : THValComboBox;
begin
        Long := length(THValComboBox(Sender).Name);
        i := 0;
        If Long=6 Then i := StrToInt(THValComboBox(Sender).Name[6]);
        If Long=7 Then i := StrToInt(THValComboBox(Sender).Name[6]+THEdit(Sender).Name[7]);
        Qte := TSpinEdit(GetControl('QTEFRAIS'+IntToStr(i)));
        If Qte.Value=0 then Qte.Value := 1;
        ValeurFrais := THValComboBox(Sender).Value;
        If ValeurFrais <> '' Then
        begin
                For j := 1 to 15 do
                begin
                        If j <> i Then
                        begin
                                Combo := THValComboBox(GetControl('FRAIS'+IntToStr(j)));
                                TestValeur := Combo.Value;
                                If TestValeur=ValeurFrais Then
                                begin
                                        PgiBox('Ce type de frais a déja été saisi, veuillez en saisir un autre','Saisie des frais'); //PT4
                                        SetFocusControl(THValcomboBox(Sender).Name);
                                        SetControlProperty(THValcomboBox(Sender).Name,'Value','');
                                        Break;
                                end;
                        end;
                end;
        end;
end;

procedure TOF_PGSAISIEFRAISSALFOR.UpdateTable(Sender : TObject);
var Combo : THValComboBox;
    Quantite : TSpinEdit;
    i : Integer;
    Montant,Plafond : THEdit;
    Total,TotalPlaf,SalaireAnim,FraisAnim : Double;
    Q : TQuery;
    IdSession,NumSession,TypeSession : String; //PT6
begin
        // Suppression des enregistrements actuels
        ExecuteSQL('DELETE FROM FRAISSALFORM'+
        ' WHERE PFS_SALARIE="'+LeSalarie+'" AND PFS_CODESTAGE="'+LeStage+'"'+
        ' AND PFS_ORDRE='+LaSession+' AND PFS_MILLESIME="'+LeMillesime+'" AND PFS_TYPEPLANPREV="'+LePlan+'"');  //DB2

        //PT6 - Début
        // Informations sur la formation rattachée
        Q := OpenSQL ('SELECT PFO_IDSESSIONFOR,PFO_NOSESSIONMULTI,PFO_PGTYPESESSION ' +
                      'FROM FORMATIONS ' +
                      'WHERE PFO_CODESTAGE="'+LeStage+'" AND PFO_ORDRE="'+LaSession+'" AND PFO_MILLESIME="'+LeMillesime+'" AND PFO_SALARIE="'+LeSalarie+'" AND PFO_TYPEPLANPREV="'+LePlan+'"',True);
        If Not Q.EOF Then
        Begin
               IdSession   := Q.FindField('PFO_IDSESSIONFOR').AsString;
               NumSession  := Q.FindField('PFO_NOSESSIONMULTI').AsString;
               TypeSession := Q.FindField('PFO_PGTYPESESSION').AsString;
        End
        Else
        Begin
               IdSession := '000000';
               NumSession  := '-1';
               TypeSession := 'AUC';
        End;
        Ferme(Q);
        //PT6 - Fin

        // Création de tous les paramètres
        For i := 1 to 15 do
        begin
                Combo := THValComboBox(GetControl('FRAIS'+IntToStr(i)));
                Quantite := TSpinEdit(GetControl('QTEFRAIS'+IntToStr(i)));
                Montant := THEdit(GetControl('MONTANT'+IntToStr(i)));
                Plafond := THEdit(GetControl('PLAFOND'+IntToStr(i)));
                If Combo.Value <> '' Then
                begin
                        If IsNumeric(Montant.Text) Then
                        begin
                                If StrToFloat(GetControltext('MONTANT'+IntToStr(i)))>0 then
                                begin
                                        If Plafond.Text = '' Then  ExitMontant(Montant);
                                        ExecuteSQL('INSERT INTO FRAISSALFORM (PFS_SALARIE,PFS_CODESTAGE,PFS_ORDRE,PFS_MILLESIME,PFS_FRAISSALFOR'+
                                        ',PFS_QUANTITE,PFS_MONTANT,PFS_MONTANTPLAF,PFS_PREDEFINI,PFS_NODOSSIER,PFS_PGTYPESESSION,PFS_NOSESSIONMULTI,'+
                                        'PFS_IDSESSIONFOR,PFS_POPULATION,PFS_CODEPOP,PFS_TYPEPLANPREV)'+  //PT5 + PT6
                                        ' VALUES '+
                                        ' ("'+LeSalarie+'","'+LeStage+'","'+LaSession+'","'+LeMillesime+'","'+Combo.Value+'"'+
                                        ',"'+StrfPOINT(Quantite.value)+'","'+StrfPOINT(StrToFloat(GetControltext('MONTANT'+IntToStr(i))))+'","'+
                                        StrfPOINT(StrToFloat(GetControlText('PLAFOND'+intToStr(i))))+'","'+
                                        GetPredefiniPopulation(TYP_POPUL_FORM_PREV)+'","'+
                                        PgrendNodossier()+'","'+TypeSession+'","'+NumSession+'","'+IdSession+'","'+
                                        DonnePopulation(LeSalarie)+'","","'+LePlan+'")'); //PT1 + PT5 + PT6
                                end;
                        end;
                end;
        end;

        // Mise à jour des montants totaux
        Q := OpenSQL('SELECT SUM(PFS_MONTANT) AS MONTANT, SUM(PFS_MONTANTPLAF) AS PLAFOND FROM FRAISSALFORM WHERE '+
        'PFS_SALARIE="'+LeSalarie+'" AND PFS_MILLESIME="'+LeMillesime+'" AND PFS_CODESTAGE="'+LeStage+'" AND PFS_ORDRE='+LaSession+'',True);  //DB2
        Total := 0;
        TotalPlaf := 0;
        if not Q.eof then
        begin
                Total := Q.FindField('MONTANT').AsFloat;
                TotalPlaf := Q.FindField('PLAFOND').AsFloat;
        end;
        Ferme(Q);

        If GetControlText('THQUALITE') = 'FOR' Then
        begin
                Q := OpenSQL('SELECT SUM(PAN_SALAIREANIM) SALAIRE FROM SESSIONANIMAT WHERE '+
                'PAN_CODESTAGE="'+LeStage+'" AND PAN_ORDRE='+LaSession+' AND PAN_MILLESIME="'+LeMillesime+'"',True);  //DB2
                If Not Q.Eof then SalaireAnim := Q.FindField('SALAIRE').AsFloat
                Else SalaireAnim := 0;
                Ferme(Q);
                Q := OpenSQL('SELECT SUM(PFS_MONTANT) FRAIS FROM SESSIONANIMAT '+
                'LEFT JOIN FRAISSALFORM ON PAN_SALARIE=PFS_SALARIE AND PAN_CODESTAGE=PFS_CODESTAGE AND PAN_ORDRE=PFS_ORDRE AND PAN_MILLESIME=PFS_MILLESIME '+
                'WHERE PAN_FRASPEDAG="X" AND PAN_CODESTAGE="'+LeStage+'" AND PAN_ORDRE='+LaSession+' AND PAN_MILLESIME="'+LeMillesime+'"',True);  //DB2
                If Not Q.Eof then FraisAnim := Q.FindField('FRAIS').AsFloat
                Else FraisAnim := 0;
                Ferme(Q);
                ExecuteSQL('UPDATE SESSIONSTAGE SET PSS_COUTSALAIR='+StrfPOINT(SalaireAnim + FraisAnim)+' '+  //DB2
                'WHERE PSS_CODESTAGE="'+LeStage+'" AND PSS_ORDRE="'+LaSession+'" AND PSS_MILLESIME="'+LeMillesime+'"');
                MAJCoutsFormation(LeMillesime,LeStage,LaSession);
        end;

        If GetControlText('THQUALITE') = 'STA' Then
        begin
                ExecuteSQL('UPDATE FORMATIONS SET PFO_FRAISREEL="'+StrfPOINT(Total)+'",PFO_FRAISPLAF="'+StrfPOINT(TotalPlaf)+'"'+ //PT1
                ' WHERE PFO_SALARIE="'+LeSalarie+'" AND PFO_MILLESIME="'+LeMillesime+'" AND PFO_CODESTAGE="'+LeStage+'" AND PFO_ORDRE='+LaSession+'');  //DB2
        end;
        CalcCtInvestSession('FRAIS',LeStage,LeMillesime,StrToInt(LaSession));

        //PT6 - Début
        // Gestion des multi-sessions : Mise à jour des frais totaux de la session en-tête
        If (TypeSession = TYP_SOUSSESSION) Then MajSousSessionSal(LeSalarie,LeStage,NumSession);
        //PT6 - Fin
end;

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 07/06/2007 / PT6
Modifié le ... :   /  /    
Description .. : Mise à jour de l'écran suite à la sélection du plan de 
Suite ........ : formation
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGSAISIEFRAISSALFOR.RefreshSelonPlan(Sender: TObject);
Var
     i : Integer;
     TobFrais,TF : TOB;
     Q : TQuery;
     Quantite : TSpinEdit;
begin
     LePlan := THValComboBox(GetControl('TYPEPLANPREV')).Value;

     // Récupération des frais
     Q := OpenSQL('SELECT PFS_QUANTITE,PFS_MONTANT,PFS_MONTANTPLAF,PFS_FRAISSALFOR FROM FRAISSALFORM '+
     ' WHERE PFS_SALARIE="'+LeSalarie+'" AND PFS_CODESTAGE="'+LeStage+'"'+
     ' AND PFS_ORDRE='+LaSession+' AND PFS_MILLESIME="'+LeMillesime+'" AND PFS_TYPEPLANPREV="'+LePlan+'"',True);  //DB2
     TobFrais := Tob.Create('Les frais',Nil,-1);
     TobFrais.LoadDetailDB('FRAISSALFORM','','',Q,False);
     Ferme(Q);

     // Mise à jour de l'écran avec les frais
     For i := 0 to TobLibFrais.Detail.Count-1 do
     begin
             If i>15 then break;

             SetControlProperty('FRAIS'+IntToStr(i+1),'Value',TobLibFrais.Detail[i].GetValue('CC_CODE'));
             TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'],[TobLibFrais.Detail[i].GetValue('CC_CODE')],False);
             If TF <> Nil then
             begin
                     Quantite := TSpinEdit(GetControl('QTEFRAIS'+IntToStr(i+1)));
                     Quantite.Value := TF.GetValue('PFS_QUANTITE');
                     SetControlText('MONTANT'+IntToStr(i+1),FloatToStr(TF.GetValue('PFS_MONTANT')));
                     SetControlText('PLAFOND'+IntToStr(i+1),FloatToStr(TF.GetValue('PFS_MONTANTPLAF')));
             end
             Else
             begin
                     Quantite := TSpinEdit(GetControl('QTEFRAIS'+IntToStr(i+1)));
                     Quantite.Value := 1;
                     SetControlText('MONTANT'+IntToStr(i+1),'0');
                     SetControlText('PLAFOND'+IntToStr(i+1),'0');
             end;
     end;

     // Libérations
     TobFrais.Free;
end;

Initialization
  registerclasses ( [ TOF_PGSAISIEFRAISSALFOR ] ) ;
end.
