{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :
Description .. : Unité de gestion de la saisie des périodes d'inactivité de la
Suite ........ : DADS-U
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
{
PT1   : 21/10/2002 VG V585 Gestion du journal des evenements
PT2   : 26/02/2003 VG V_42 Libération de TOB non faite mise en évidence avec
                           memcheck
PT3   : 13/03/2003 VG V_42 Gestion du journal des evenements
PT4   : 22/10/2003 VG V_42 Violation d'accès en saisie des périodes d'inactivité
                           FQ N°10922
PT5-1 : 01/03/2004 VG V_50 On ne permet plus de saisir une période avec des
                           dates en dehors de l'exercice en cours - FQ N°11090
PT5-2 : 01/03/2004 VG V_50 Suppression des TControl
PT6   : 12/05/2004 VG V_50 Ajout du traitement lié aux raccourcis clavier
PT7   : 05/07/2004 VG V_50 Adaptation cahier des charges V8R00
PT8   : 15/09/2004 VG V_50 Les périodes d'inactivité ne doivent plus se
                           chevaucher - FQ N°11522
PT9   : 24/09/2004 VG V_50 Adaptation cahier des charges V8R01
PT10  : 06/01/2005 VG V_60 Le segment S46.G01.02.002 Temps d'arrêt des périodes
                           d'inactivité doit être * 100 - FQ N°11894
PT11-1: 02/02/2005 VG V_60 Le segment S46.G01.00.004 doit être alimenté lorsque
                           le début de la période d'inactivité est antérieur OU
                           EGAL au début d'exercice - FQ N°11959
PT11-2: 02/02/2005 VG V_60 Contrôles sur périodes d'inactivité - FQ N°11960
PT12  : 07/10/2005 VG V_60 Adaptation cahier des charges DADS-U V8R02
PT13-1: 16/10/2006 VG V_70 Traitement des DADS-U IRCANTEC
PT13-2: 16/10/2006 VG V_70 Utilisation d'un type pour la cle DADS-U
PT13-3: 16/10/2006 VG V_70 Adaptation cahier des charges DADS-U V8R04
PT13-4: 16/10/2006 VG V_70 Suppression du fichier de contrôle - mise en table
                           des erreurs
PT14  : 21/11/2006 VG V_70 Permettre de déclarer des honoraires en DADS-U
                           complémentaire - FQ N°13613
PT15-1: 11/01/2007 VG V_72 En saisie complémentaire, suppression des erreurs
                           lors de la validation d'une période d'inactivité
                           FQ N°13834
PT15-2: 11/01/2007 VG V_72 Suppression des enregistrements DADSCONTROLE lors de
                           la suppression d'une période - FQ N°13833
PT16  : 12/03/2007 VG V_72 Mauvaise alimentation du code "Début période
                           d'inactivité" en cas de décalage de paie - FQ N°13871
PT17  : 05/11/2007 VG V_80 Adaptation cahier des charges V8R06
PT18  : 19/11/2007 NA V_80 FQ 14877 Affichage des motifs d'inactivité en fonction du régime
PT19  : 19/11/2007 NA V_80 FQ 14447 ne plus afficher les codes unités expression tps arrêt 04,06,08
}
unit UTOFPG_DADSINACT;

interface
uses     UTOF,
{$IFNDEF EAGLCLIENT}
         DBCtrls,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
         UtileAGL,
{$ENDIF}
         Hctrls,
         ComCtrls,
         HEnt1,
         HMsgBox,
         Classes,
         sysutils,
         UTob,
         HTB97,
         Vierge,
         PgDADSOutils,
         PgDADSCommun,
         Pgoutils2,
         Controls,
{$IFNDEF DADSUSEULE}
         P5Def,
{$ENDIF}      
         windows,
         ExtCtrls,
         StdCtrls;

Type
      TOF_PG_DADSINACT = Class (TOF)
      public
        procedure OnArgument (stArgument : String ) ; override ;
        procedure OnLoad; override ;
        procedure OnClose ; override ;

      private

        Nbre, NbreTot, Sal, State : string;

        QMul, QPer : TQUERY;     // Query recuperee du mul

        ControleOK, SalChange : Boolean;

        T_Periode : TOB;

        Daccord, PerPrec, PerSuiv, SalPrem, SalPrec, SalSuiv : TToolBarButton97;
        SalDern : TToolBarButton97;

        Regime : TRadioGroup;

        Anterieur, Ouvre : TCheckBox;

        procedure New(Sender: TObject);
        procedure Del(Sender: TObject);
        procedure AfficheCaption();
        procedure ChargeZones ();
        procedure SauveZones();
        procedure MetABlanc();
        function ControleConform() : boolean;
        function UpdateTable() : boolean;
        procedure PerPrecClick (Sender: TObject);
        procedure PerSuivClick (Sender: TObject);
{$IFNDEF EAGLCLIENT}
        procedure SalPremClick (Sender: TObject);
        procedure SalPrecClick (Sender: TObject);
        procedure SalSuivClick (Sender: TObject);
        procedure SalDernClick (Sender: TObject);
{$ENDIF}
        Function BougeSal(Button: TNavigateBtn) : boolean ;
        procedure GereQuerySal();
        Function BougePer(Button: TNavigateBtn) : boolean ;
        procedure GereQueryPer();
        procedure MAJQuery();
        procedure Impression(Sender: TObject);
        procedure Validation(Sender: TObject);
        procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
        procedure RegimeClick(Sender: TObject);
        procedure choixmotif; // pt18
     END ;

var PGDADSEtab : string;

implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSINACT.OnArgument(stArgument: String);
var
Pages : TPageControl;
Arg, StSQL, listeplus : string;  // PT19
Ircantec : boolean;
begin
Inherited ;
//Récupération des arguments
Arg := stArgument;
State := Trim(ReadTokenPipe(Arg,';'));
Sal := Trim(ReadTokenPipe(Arg,';')) ;
TypeD := Trim(ReadTokenPipe(Arg,';')) ;

TFVierge(Ecran).OnKeyDown:=FormKeyDown;

Anterieur:= TCheckBox (GetControl ('CHANTERIEUR'));
Ouvre:= TCheckBox (GetControl ('CHOUVRE'));

SetControlText('LSALARIE_', Sal);

// Positionnement sur le premier onglet
Pages := TPageControl(GetControl('PAGES'));
if Pages<>nil then
   Pages.ActivePageIndex:=0;

// recuperation de la query du multicritere
QMul := TFVierge(Ecran).FMULQ ;

// Gestion du navigateur
SalPrem := TToolbarButton97(GetControl('BSALPREM'));
SalPrec := TToolbarButton97(GetControl('BSALPREC'));
SalSuiv := TToolbarButton97(GetControl('BSALSUIV'));
SalDern := TToolbarButton97(GetControl('BSALDERN'));
{$IFNDEF EAGLCLIENT}
if SalPrem <> NIL then
   begin
   SalPrem.Enabled := True;
   SalPrem.Visible := True;
   SalPrem.OnClick := SalPremClick;
   end;

if SalPrec <> NIL then
   begin
   SalPrec.Enabled := True;
   SalPrec.Visible := True;
   SalPrec.OnClick := SalPrecClick;
   end;

if SalSuiv <> NIL then
   begin
   SalSuiv.Enabled := True;
   SalSuiv.Visible := True;
   SalSuiv.OnClick := SalSuivClick;
   end;

if SalDern <> NIL then
   begin
   SalDern.Enabled := True;
   SalDern.Visible := True;
   SalDern.OnClick := SalDernClick;
   end;
{$ELSE}
if SalPrem <> NIL then
   SalPrem.Visible := False;

if SalPrec <> NIL then
   SalPrec.Visible := False;

if SalSuiv <> NIL then
   SalSuiv.Visible := False;

if SalDern <> NIL then
   SalDern.Visible := False;
{$ENDIF}

SalChange := True;

PerPrec := TToolbarButton97(GetControl('BPERPREC'));
if PerPrec <> NIL then
   begin
   PerPrec.Enabled := True;
   PerPrec.Visible := True;
   PerPrec.OnClick := PerPrecClick;
   end;

PerSuiv := TToolbarButton97(GetControl('BPERSUIV'));
if PerSuiv <> NIL then
   begin
   PerSuiv.Enabled := True;
   PerSuiv.Visible := True;
   PerSuiv.OnClick := PerSuivClick;
   end;

TFVierge(Ecran).Binsert.OnClick := New;
TFVierge(Ecran).BDelete.OnClick := Del;

TFVierge(Ecran).BImprimer.OnClick := Impression;

Daccord := TToolbarButton97(GetControl('BDACCORD'));
if Daccord <> NIL then
   begin
   Daccord.Enabled := True;
   Daccord.Visible := True;
   Daccord.OnClick := Validation;
   end;

// Deb PT19 : Affichage code unité expression temps d'arrêt 01,03 et 07
listeplus := '01,03,07';
setcontrolproperty('CTEMPSARRET', 'plus' , 'AND CO_CODE in ('+listeplus+')');
// Fin PT19

//PT13-1
Regime:= TRadioGroup (GetControl ('RGIRCANTEC'));
StSQL:= 'SELECT PDS_SALARIE'+
        ' FROM DADSDETAIL WHERE'+
        ' PDS_SALARIE="'+Sal+'" AND'+
        ' PDS_SEGMENT="S41.G01.01.001" AND'+
        ' (PDS_DONNEEAFFICH="I001" OR PDS_DONNEEAFFICH="I002")';
Ircantec:= ExisteSql (StSQL);
if (Ircantec=True) then
   begin
   Regime.ItemIndex:= 0;
   SetControlVisible ('GBELEMCOMPL', False);
   SetControlVisible ('GBIRCANTECSITPARTIC', True);
   end
else
   begin
   Regime.ItemIndex:= 1;
   SetControlVisible ('GBELEMCOMPL', True);
   SetControlVisible ('GBIRCANTECSITPARTIC', False);
   end;
choixmotif; // PT18
Regime.OnClick:= RegimeClick;
//FIN PT13-1
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :   /  /    
Description .. : procédure exécutée sur le Binsert.onclick
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSINACT.New(Sender: TObject);
var
IMax : Integer;
QRechNum :TQuery;
begin
{PT5-2
if ((NomSal  <> NIL) and (PrenomSal <> NIL)) then
   Ecran.Caption:= 'Inactivité : Salarié '+Sal+' '+
                   NomSal.Text+' '+PrenomSal.Text+' Période nouvelle';
}
Ecran.Caption:= 'Inactivité : Salarié '+Sal+' '+GetControlText('ENOM')+' '+
                GetControlText('EPRENOM')+' Période nouvelle';

MetABlanc;

State := 'CREATION';
IMax:= -1;
QRechNum:=OpenSQL('SELECT MIN(PDE_ORDRE)'+
                  ' FROM DADSPERIODES WHERE'+
                  ' PDE_SALARIE ="'+Sal+'" AND'+
                  ' PDE_ORDRE < 0 AND'+
                  ' PDE_EXERCICEDADS = "'+PGExercice+'"',TRUE);

if Not QRechNum.EOF then
   try
   IMax:= QRechNum.Fields[0].AsInteger-1;
   except
         on E: EConvertError do
            IMax:= -1;
   end;
Ferme(QRechNum);
{PT5-2
if OrdreDADS <> NIL then
   OrdreDADS.Text := IntToStr(IMax);
if TypeDADS <> NIL then
   TypeDADS.Text := TypeD;
if DateDeb <> NIL then
   DateDeb.Text := DateToStr(IDate1900);
if DateFin <> NIL then
   DateFin.Text := DateToStr(IDate1900);
}
SetControlText('ORDREDADS', IntToStr(IMax));
SetControlText('TYPEDADS', TypeD);
SetControlText('EDEBUT', DateToStr(IDate1900));
SetControlText('EFIN', DateToStr(IDate1900));
//FIN PT5-2

// Gestion du navigateur
if SalPrem <> NIL then
   SalPrem.Enabled := False;

if SalPrec <> NIL then
   SalPrec.Enabled := False;

if SalSuiv <> NIL then
   SalSuiv.Enabled := False;

if SalDern <> NIL then
   SalDern.Enabled := False;

if PerPrec <> NIL then
   PerPrec.Enabled := False;

if PerSuiv <> NIL then
   PerSuiv.Enabled := False;

TFVierge(Ecran).Binsert.Enabled := False;
TFVierge(Ecran).BDelete.Enabled := False;

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :   /  /
Description .. : procédure exécutée sur le bdelete.onclick
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSINACT.Del(Sender: TObject);
var
Rep : integer;
begin
Rep:=PGIAsk ('Voulez vous supprimer cette fiche ?', 'Suppression DADS-U') ;
if Rep=mrNo then
   exit
else
   begin
   try
      begintrans;
{PT5-2
      DeletePeriode(Sal, TypeD, StrToInt(OrdreDADS.Text));
      DeleteDetail(Sal, TypeD, StrToInt(OrdreDADS.Text));
}
{PT14
      DeletePeriode(Sal, TypeD, StrToInt(GetControlText('ORDREDADS')));
      DeleteDetail(Sal, TypeD, StrToInt(GetControlText('ORDREDADS')));
}
      DeletePeriode(Sal, StrToInt(GetControlText('ORDREDADS')));
      DeleteDetail(Sal, StrToInt(GetControlText('ORDREDADS')));
      DeleteErreur (Sal, 'S4', StrToInt(GetControlText('ORDREDADS'))); //PT15-2
//PT3
      Trace := TStringList.Create;
      if Trace <> nil then
{
         Trace.Add (Sal+' : Suppression de la période d''inactivité'+
                   IntToStr(-StrToInt(OrdreDADS.Text)));
}
         Trace.Add (Sal+' : Suppression de la période d''inactivité'+
                    IntToStr (-StrToInt (GetControlText ('ORDREDADS'))));
{$IFNDEF DADSUSEULE}
      CreeJnalEvt('001', '043', 'OK', NIL, NIL, Trace);
{$ENDIF}      
      if Trace <> nil then
         FreeAndNil (Trace);
//FIN PT3
      MAJQuery;
      CommitTrans;
   Except
      Rollback;
      PGIBox ('Une erreur est survenue lors de la mise à jour de la base', 'Suppression DADS-U');
      END;
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :   /  /    
Description .. : Procédure exécutée sur le chargement de la fiche
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSINACT.AfficheCaption();
var
PeriodeCaption : string;
begin
if ((Nbre<>'0') and (NbreTot<>'0')) then
   begin
   if (Nbre <> '-1') then
      PerPrec.Enabled := TRUE
   else
      PerPrec.Enabled := FALSE;


   if (Nbre <> NbreTot) then
      PerSuiv.Enabled := TRUE
   else
      PerSuiv.Enabled := FALSE;


{PT5-2
   if ((NomSal  <> NIL) and (PrenomSal <> NIL)) then
      begin
      PeriodeCaption := 'Salarié '+Sal+' '+NomSal.Text+' '+PrenomSal.Text+
                             '   Période '+IntToStr(-StrToInt(Nbre))+'/'+
                             IntToStr(-StrToInt(NbreTot));
      Ecran.Caption:= 'Inactivité : '+PeriodeCaption;
      end;
}
   PeriodeCaption:= 'Salarié '+Sal+' '+GetControlText('ENOM')+' '+
                    GetControlText('EPRENOM')+
                    '   Période '+IntToStr(-StrToInt(Nbre))+'/'+
                    IntToStr(-StrToInt(NbreTot));
   Ecran.Caption:= 'Inactivité : '+PeriodeCaption;
   end
else
   begin
{PT5-2
   if ((NomSal  <> NIL) and (PrenomSal <> NIL)) then
      Ecran.Caption:= 'Inactivité : Salarié '+Sal+' '+
                      NomSal.Text+' '+PrenomSal.Text+'   Période Nouvelle';
}
   Ecran.Caption:= 'Inactivité : Salarié '+Sal+' '+GetControlText('ENOM')+' '+
                   GetControlText('EPRENOM')+'   Période Nouvelle';
   TFVierge(Ecran).Binsert.Enabled := False;
   TFVierge(Ecran).BDelete.Enabled := False;
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :   /  /    
Description .. : Procédure exécutée sur le chargement de la fiche
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSINACT.OnLoad;
begin
Inherited;
ChargeZones;
AfficheCaption;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSINACT.OnClose;
begin
Ferme(QPer);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :
Description .. : Chargement des éléments de la fiche
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSINACT.ChargeZones();
var
StDADSD, StPer, StSal : String;
QCount, QRechDADSD, QRechSal : TQuery;
TDetail, TDetailD : Tob;
DateEntree, DateSortie : TDateTime;
begin
MetABlanc;
SetControlText('ENOM', '');
SetControlText('EPRENOM', '');

if (SalChange = True) then
   begin
   StPer:= 'SELECT COUNT(PDE_SALARIE) AS NBRE'+
           ' FROM DADSPERIODES WHERE'+
           ' PDE_SALARIE="'+Sal+'" AND'+
           ' PDE_TYPE="'+TypeD+'" AND'+
           ' PDE_ORDRE < 0 AND'+
           ' PDE_EXERCICEDADS = "'+PGExercice+'"';

   QCount:= OpenSql(StPer, True);
   if (not QCount.EOF) then
      NbreTot:= IntToStr(-QCount.FindField('NBRE').AsInteger);
   Ferme(QCount);

   if (QPer <> nil) then
      Ferme(QPer);
   StPer:= 'SELECT *'+
           ' FROM DADSPERIODES WHERE'+
           ' PDE_SALARIE = "'+Sal+'" AND'+
           ' PDE_TYPE="'+TypeD+'" AND'+
           ' PDE_ORDRE < 0 AND'+
           ' PDE_EXERCICEDADS = "'+PGExercice+'"'+
           ' ORDER BY PDE_DATEDEBUT';

   QPer:= OpenSql(StPer, True);
   QPer.First;
   PerPrec.Enabled:= FALSE;
   if QPer.EOF then
      begin
      PerSuiv.Enabled:= FALSE;
      Nbre:= '0';
      end
   else
      begin
      PerSuiv.Enabled:= TRUE;
      Nbre:= '-1';
      end;
   SalChange:= False;
   end;

if (not QPer.eof) then
   begin
   if (QPer.FindField ('PDE_ORDRE').Asinteger = 0) then
      begin
      SetControlText('ORDREDADS', '-1');
      State:= 'CREATION';
      end
   else
      begin
      SetControlText('ORDREDADS', IntToStr(QPer.FindField ('PDE_ORDRE').Asinteger));
      State:= 'MODIFICATION';
      end;
   SetControlText('TYPEDADS', QPer.FindField ('PDE_TYPE').Asstring);
   SetControlText('CMOTIF', QPer.FindField ('PDE_MOTIFDEB').Asstring);
   SetControlText('EDEBUT', DateToStr(QPer.FindField ('PDE_DATEDEBUT').AsDateTime));
   SetControlText('EFIN', DateToStr(QPer.FindField ('PDE_DATEFIN').AsDateTime));
   end
else
   begin
   State:= 'CREATION';
   SetControlText('ORDREDADS', '-1');
   SetControlText('TYPEDADS', '');
   SetControlText('CMOTIF', '');
   SetControlText('EDEBUT', DateToStr(IDate1900));
   SetControlText('EFIN', DateToStr(IDate1900));
   end;


StSal:= 'SELECT PSA_DATEENTREE, PSA_DATESORTIE, PSA_LIBELLE, PSA_PRENOM'+
        ' FROM SALARIES WHERE'+
        ' PSA_SALARIE="'+Sal+'"';

QRechSal:= OpenSql(StSal,TRUE);
if (not QRechSal.EOF) then
   begin
   SetControlText ('ENOM', QRechSal.FindField('PSA_LIBELLE').Asstring);
   SetControlText ('EPRENOM', QRechSal.FindField ('PSA_PRENOM').Asstring);

   if (State = 'CREATION') then // chargement uniquement en mode creation
      begin
      DateEntree:= QRechSal.FindField('PSA_DATEENTREE').AsDateTime;
      if (DateEntree >= DebExer) then
         SetControlText('EDEBUT', DateToStr(DateEntree))
      else
         SetControlText('EDEBUT', DateToStr(DebExer));

      DateSortie:= QRechSal.FindField('PSA_DATESORTIE').AsDateTime;
      if ((DateSortie <= FinExer) and (DateSortie > IDate1900)) then
         SetControlText('EFIN', DateToStr(DateSortie))
      else
         SetControlText('EFIN', DateToStr(FinExer));
      end
   else
      if (State = 'MODIFICATION') then // chargement uniquement en mode modification
         begin
         StDADSD:= 'SELECT PDS_SALARIE, PDS_TYPE, PDS_ORDRE, PDS_DATEDEBUT,'+
                   ' PDS_DATEFIN, PDS_ORDRESEG, PDS_SEGMENT, PDS_DONNEE,'+
                   ' PDS_DONNEEAFFICH'+
                   ' FROM DADSDETAIL WHERE'+
                   ' PDS_SALARIE="'+Sal+'" AND'+
                   ' PDS_TYPE="'+TypeD+'" AND'+
                   ' PDS_ORDRE='+GetControlText('ORDREDADS')+' AND'+
                   ' PDS_EXERCICEDADS = "'+PGExercice+'"'+
                   ' ORDER BY PDS_ORDRESEG,PDS_SEGMENT,PDS_DATEDEBUT,'+
                   ' PDS_DATEFIN';
         QRechDADSD:= OpenSql(StDADSD,TRUE);
         TDetail:= TOB.Create('Les détails', NIL, -1);
         TDetail.LoadDetailDB('DADSDETAIL','','',QRechDADSD,False);
         Ferme(QRechDADSD);

         TDetailD:= TDetail.FindFirst(['PDS_SEGMENT'],
                                      ['S46.G01.00.001'], TRUE);
         if (TDetailD <> NIL) then
            SetControlText('CMOTIF', TDetailD.GetValue('PDS_DONNEEAFFICH'));

         TDetailD:= TDetail.FindFirst(['PDS_SEGMENT'],
                                      ['S46.G01.00.002'], TRUE);
         if (TDetailD <> NIL) then
            SetControlText('EDEBUT', TDetailD.GetValue('PDS_DONNEEAFFICH'));

         TDetailD:= TDetail.FindFirst(['PDS_SEGMENT'],
                                      ['S46.G01.00.003'], TRUE);
         if (TDetailD <> NIL) then
            SetControlText('EFIN', TDetailD.GetValue('PDS_DONNEEAFFICH'));

//PT17
         TDetailD:= TDetail.FindFirst(['PDS_SEGMENT'], ['S46.G01.00.004'], TRUE);
         if ((Anterieur<>NIL) and (TDetailD<>NIL)) then
            Anterieur.Checked:= TDetailD.GetValue ('PDS_DONNEEAFFICH')='01';
//FIN PT17

//PT13-1
         TDetailD:= TDetail.FindFirst (['PDS_SEGMENT'],
                                       ['S46.G01.01.002.001'], TRUE);
         if (TDetailD <> NIL) then
            SetControlText ('EIRCANTECMONTANTPART',
                            TDetailD.GetValue ('PDS_DONNEEAFFICH'));

         TDetailD:= TDetail.FindFirst (['PDS_SEGMENT'],
                                       ['S46.G01.01.004'], TRUE);
         if (TDetailD <> NIL) then
            SetControlText ('EIRCANTECTAUX',
                            TDetailD.GetValue ('PDS_DONNEEAFFICH'));
//FIN PT13-1

{PT17
         TDetailD:= TDetail.FindFirst(['PDS_SEGMENT'],
                                       ['S46.G01.02.001'], TRUE);
}
         TDetailD:= TDetail.FindFirst(['PDS_SEGMENT'],
                                       ['S46.G01.02.001.001'], TRUE);
//FIN PT17
         if (TDetailD <> NIL) then
            SetControlText('CTEMPSARRET', TDetailD.GetValue('PDS_DONNEEAFFICH'));

//PT17
         TDetailD:= TDetail.FindFirst(['PDS_SEGMENT'],
                                       ['S46.G01.02.001.002'], TRUE);
         if ((Ouvre<>NIL) and (TDetailD<>NIL)) then
            Ouvre.Checked:= TDetailD.GetValue ('PDS_DONNEEAFFICH')='02';
//FIN PT17

         TDetailD:= TDetail.FindFirst(['PDS_SEGMENT'],
                                      ['S46.G01.02.002'], TRUE);
         if (TDetailD <> NIL) then
            SetControlText('ETEMPSARRET', TDetailD.GetValue('PDS_DONNEEAFFICH'));

         TDetailD:= TDetail.FindFirst(['PDS_SEGMENT'],
                                      ['S46.G01.02.003.001'], TRUE);
         if (TDetailD <> NIL) then
            SetControlText('EMONTANTEMPL', TDetailD.GetValue('PDS_DONNEEAFFICH'));
         end;
   end;
Ferme(QRechSal);
if (TDetail<>nil) then
   FreeAndNil(TDetail);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :
Description .. : Enregistrement des éléments de la fiche
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSINACT.SauveZones();
var
BufDest, BufOrig : String;
DateBuf, PerDeb, PerFin : TDateTime;
Ordre : integer;
CleDADS : TCleDADS;
begin
Trace:= TStringList.Create;

Ordre:= StrToInt (GetControlText ('ORDREDADS'));

{PT17
if (StrToDate(GetControlText('EDEBUT'))<DebExer) then
   PerDeb:= DebExer
else
}
PerDeb:= StrToDate(GetControlText('EDEBUT'));
if (StrToDate(GetControlText('EFIN'))>FinExer) then
   PerFin:= FinExer
else
   PerFin:= StrToDate(GetControlText('EFIN'));

{PT13-2
CreeEntete(Sal, TypeD, Ordre, PerDeb, PerFin, GetControlText('CMOTIF'), '');
}
CleDADS.Salarie:= Sal;
CleDADS.TypeD:= TypeD;
CleDADS.Num:= Ordre;
CleDADS.DateDeb:= PerDeb;
CleDADS.DateFin:= PerFin;
CleDADS.Exercice:= PGExercice;
CreeEntete (CleDADS, GetControlText('CMOTIF'), '', Decalage);
//FIN PT13-2
CreeDetail (CleDADS, 1, 'S46.G01.00.001',
            RechDom('PGTYPEABS', GetControlText('CMOTIF'), TRUE),
            GetControlText('CMOTIF'));

ForceNumerique(GetControlText('EDEBUT'), BufDest);
{PT17
CreeDetail (CleDADS, 2, 'S46.G01.00.002', Copy(BufDest, 1, 4),
            GetControlText('EDEBUT'));
}
CreeDetail (CleDADS, 2, 'S46.G01.00.002', BufDest, GetControlText('EDEBUT'));
//FIN PT17

ForceNumerique(GetControlText('EFIN'), BufDest);
{PT17
CreeDetail (CleDADS, 3, 'S46.G01.00.003', Copy(BufDest, 1, 4),
            GetControlText('EFIN'));
}
CreeDetail (CleDADS, 3, 'S46.G01.00.003', BufDest, GetControlText('EFIN'));
//FIN PT17

{PT13-3
if (StrToDate(GetControlText('EDEBUT'))<DebExer) then
   CreeDetail (CleDADS, 4, 'S46.G01.00.004', '02', '02');
}
{PT16
if ((StrToDate(GetControlText('EDEBUT'))<DebExer) and
   (StrToDate(GetControlText('EFIN'))<DebExer)) then
   CreeDetail (CleDADS, 4, 'S46.G01.00.004', '02', '02')
else
if ((StrToDate(GetControlText('EDEBUT'))<DebExer) and
   (StrToDate(GetControlText('EFIN'))>DebExer)) then
   CreeDetail (CleDADS, 4, 'S46.G01.00.004', '03', '03');
}
{PT17
BufOrig:= DateToStr (DebExer);
if ((Copy (BufOrig, 1, 5)='01/12') and ((Copy (TypeD, 2, 2)='01') or
   (Copy (TypeD, 2, 2)='02'))) then
   BufOrig:= '01/01/'+ IntToStr (StrToInt (Copy (BufOrig, 7, 4))+1);
DateBuf:= DebutDeMois (StrToDate (BufOrig));

if ((StrToDate(GetControlText('EDEBUT'))<DateBuf) and
   (StrToDate(GetControlText('EFIN'))<DateBuf)) then
   CreeDetail (CleDADS, 4, 'S46.G01.00.004', '02', '02')
else
if ((StrToDate(GetControlText('EDEBUT'))<DateBuf) and
   (StrToDate(GetControlText('EFIN'))>DateBuf)) then
   CreeDetail (CleDADS, 4, 'S46.G01.00.004', '03', '03');
}
if (StrToDate (GetControlText ('EDEBUT'))<DebExer) then
   CreeDetail (CleDADS, 4, 'S46.G01.00.004', '01', '01')
else
   if (Anterieur.Checked=True) then
      CreeDetail (CleDADS, 4, 'S46.G01.00.004', '01', '01');
//FIN PT17
//FIN PT16


//PT13-1
if ((Regime <> nil) and (Regime.ItemIndex = 0)) then
   begin
   CreeDetail (CleDADS, 5, 'S46.G01.01.001', '01', '01');

   if (GetControlText('EIRCANTECMONTANTPART')<>'') then
      begin
      CreeDetail (CleDADS, 6, 'S46.G01.01.002.001',
                  FloatToStr(Abs(Arrondi(StrToFloat(GetControlText('EIRCANTECMONTANTPART')), 0))),
                  FloatToStr(Arrondi(StrToFloat(GetControlText('EIRCANTECMONTANTPART')), 0)));
      if (StrToFloat(GetControlText('EIRCANTECMONTANTPART')) < 0) then
         CreeDetail (CleDADS, 7, 'S46.G01.01.002.002', 'N', 'N');
      end
   else
      CreeDetail (CleDADS, 6, 'S46.G01.01.002.001', '0', '0');

   BufOrig:= GetControlText ('EIRCANTECTAUX');
   if ((BufOrig <> '0') and (BufOrig <> '')) then
      begin
      BufOrig:= FloatToStr (Arrondi (StrToFloat (BufOrig), 2));
      BufDest:= FloatToStr (StrToFloat (BufOrig)*100);
      CreeDetail (CleDADS, 8, 'S46.G01.01.004', BufDest, BufOrig);
      end;
   end
else
   begin
{PT17
   CreeDetail (CleDADS, 9, 'S46.G01.02.001', GetControlText('CTEMPSARRET'),
               GetControlText('CTEMPSARRET'));
}
   CreeDetail (CleDADS, 9, 'S46.G01.02.001.001', GetControlText('CTEMPSARRET'),
               GetControlText('CTEMPSARRET'));

   if (Ouvre.Checked=True) then
      CreeDetail (CleDADS, 10, 'S46.G01.02.001.002', '02', '02')
   else
      CreeDetail (CleDADS, 10, 'S46.G01.02.001.002', '01', '01');
//FIN PT17

   CreeDetail (CleDADS, 11, 'S46.G01.02.002',
               FloatToStr(StrToFloat(GetControlText('ETEMPSARRET'))*100),
               GetControlText('ETEMPSARRET'));

   if (GetControlText('EMONTANTEMPL')<>'') then
      begin
      CreeDetail (CleDADS, 12, 'S46.G01.02.003.001',
                  FloatToStr(Abs(Arrondi(StrToFloat(GetControlText('EMONTANTEMPL')), 0))),
                  FloatToStr(Arrondi(StrToFloat(GetControlText('EMONTANTEMPL')), 0)));
      if (StrToFloat(GetControlText('EMONTANTEMPL')) < 0) then
         CreeDetail (CleDADS, 13, 'S46.G01.02.003.002', 'N', 'N');
      end;
   end;
//FIN PT13-1

{$IFNDEF DADSUSEULE}
CreeJnalEvt('001', '042', 'OK', NIL, NIL, Trace);
{$ENDIF}
if Trace <> nil then
   FreeAndNil (Trace);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :   /  /
Description .. : Effacement des zones de la fiche de périodes d'inactivité
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSINACT.MetABlanc();
begin
SetControlText('CMOTIF', '');
SetControlText('EDEBUT', '');
SetControlText('EFIN', '');
//PT17
if (Anterieur<>NIL) then
   Anterieur.Checked:= False;
if (Ouvre<>nil) then
   Ouvre.Checked:= True;
//FIN PT17
//PT13-1
SetControlText('EIRCANTECMONTANTPART', '');
SetControlText('EIRCANTECTAUX', '');
//FIN PT13-1
SetControlText('ETEMPSARRET', '');
SetControlText('CTEMPSARRET', '');
SetControlText('EMONTANTEMPL', '');
SetControlText('LSALARIE_', Sal);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :   /  /
Description .. : Contrôle des données
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
function TOF_PG_DADSINACT.ControleConform() : boolean;
var
DateDebBuf, DateFinBuf : TDateTime;
Mess, StControlDate, StDate : string;
QRechControlDate : TQuery;
TControlDate, TControlDateD : TOB;
Ordre, OrdreBuf : integer;
ErreurDADSU : TControle;
begin
result := TRUE;
Mess:= '';
Ordre := StrToInt(GetControlText('ORDREDADS'));

//PT13-2
ErreurDADSU.Salarie:= Sal;
ErreurDADSU.TypeD:= TypeD;
ErreurDADSU.Num:= Ordre;
ErreurDADSU.DateDeb:= StrToDate(GetControlText('EDEBUT'));
ErreurDADSU.DateFin:= StrToDate(GetControlText('EFIN'));
ErreurDADSU.Exercice:= PGExercice;
//FIN PT13-2

if (GetControlText('CMOTIF')='') then
   begin
   result := FALSE;
   Mess:= Mess+ '#13#10 - Le motif de période d''inactivité';
   SetFocusControl('CMOTIF');
//PT13-4
   ErreurDADSU.Segment:= 'S46.G01.00.001';
   ErreurDADSU.Explication:= 'Le motif de période d''inactivité n''est pas'+
                             ' renseigné';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur(ErreurDADSU);
//FIN PT13-4
   end;

if (RechDom('PGTYPEABS', GetControlText('CMOTIF'), TRUE)='') then
   begin
   result:= FALSE;
   Mess:= Mess+ '#13#10 - Le motif de période d''inactivité est invalide';
   SetFocusControl('CMOTIF');
//PT13-4
   ErreurDADSU.Segment:= 'S46.G01.00.001';
   ErreurDADSU.Explication:= 'Le motif de période d''inactivité n''est pas'+
                             ' renseigné';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur(ErreurDADSU);
//FIN PT13-4
   end;

if ((StrToDate(GetControlText('EDEBUT'))>StrToDate(GetControlText('EFIN'))) or
   (StrToDate(GetControlText('EDEBUT'))<PlusMois(DebExer, -2)) or
   (StrToDate(GetControlText('EFIN'))>FinExer)) then
   begin
   result := FALSE;
   Mess:= Mess+ '#13#10 - Les dates sont incorrectes';
   SetControlText('EDEBUT', DateToStr(IDate1900));
   SetControlText('EFIN', DateToStr(IDate1900));
   SetFocusControl('EDEBUT');
//PT13-4
   ErreurDADSU.Segment:= 'S46.G01.00.002';
   ErreurDADSU.Explication:= 'Les dates de début et de fin sont incorrectes';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur(ErreurDADSU);
//FIN PT13-4
   end;

//PT13-1+PT13-4
if ((Regime <> nil) and (Regime.ItemIndex = 0)) then
   begin
   if (GetControlText ('EIRCANTECMONTANTPART')='') then
      begin
      SetControlText ('EIRCANTECMONTANTPART', '0');
      ErreurDADSU.Segment:= 'S46.G01.01.002.001';
      ErreurDADSU.Explication:= 'Le montant de la situation particulière a été'+
                                ' forcé à 0';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
      end;

   if (GetControlText ('EIRCANTECTAUX')='') then
      begin
      SetControlText ('EIRCANTECTAUX', '0');
      ErreurDADSU.Segment:= 'S46.G01.01.004';
      ErreurDADSU.Explication:= 'Le taux de la rémunération soumise à'+
                                ' cotisation a été forcé à 0';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
      end;
   end
else
   begin
   if ((GetControlText('CMOTIF')='CHO') and
      (GetControlText ('CTEMPSARRET') <> '01')) then
      begin
      SetControlText ('CTEMPSARRET', '01');
      ErreurDADSU.Segment:= 'S46.G01.02.001.001';
      ErreurDADSU.Explication:= 'Le code unité d''expression du temps d''arrêt'+
                                ' a été forcé à "heures et centièmes"';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
      end;

   if (GetControlText ('CTEMPSARRET')='') then
      begin
      result:= FALSE;
      Mess:= Mess+ '#13#10 - L''unité du temps d''arrêt';
      SetFocusControl ('CTEMPSARRET');
      ErreurDADSU.Segment:= 'S46.G01.02.001.001';
      ErreurDADSU.Explication:= 'Le code unité d''expression du temps d''arrêt'+
                                ' n''est pas renseigné';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end;

   if (GetControlText ('ETEMPSARRET')='') then
      begin
      result:= FALSE;
      Mess:= Mess+ '#13#10 - Le temps d''arrêt';
      SetFocusControl ('ETEMPSARRET');
      ErreurDADSU.Segment:= 'S46.G01.02.002';
      ErreurDADSU.Explication:= 'Le temps d''arrêt n''est pas renseigné';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end;

   if ((GetControlText('CMOTIF')='CHO') and
      ((GetControlText('EMONTANTEMPL')= '') or
      (GetControlText('EMONTANTEMPL')= '0'))) then
      begin
      result:= FALSE;
      Mess:= Mess+ '#13#10 - Le montant versé par l''employeur pour une#13#10'+
                   'inactivité de type "Chômage intempéries"';
      SetFocusControl('EMONTANTEMPL');
      ErreurDADSU.Segment:= 'S46.G01.02.003.001';
      ErreurDADSU.Explication:= 'Le montant versé par l''employeur n''est pas'+
                                ' renseigné';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end;
   end;
//FIN PT13-1+PT13-4

if result = FALSE then
   begin
   PGIBox('Vous devez renseigner :'+Mess,'Contrôle DADS-U');
   end;

StControlDate:= 'SELECT PDE_SALARIE, PDE_TYPE, PDE_ORDRE, PDE_DATEDEBUT,'+
                ' PDE_DATEFIN'+
                ' FROM DADSPERIODES WHERE'+
                ' PDE_SALARIE = "'+Sal+'" AND'+
                ' PDE_TYPE = "'+TypeD+'" AND'+
                ' PDE_ORDRE < 0 AND'+
                ' PDE_EXERCICEDADS = "'+PGExercice+'"'+
                ' ORDER BY PDE_DATEDEBUT';

QRechControlDate:=OpenSql(StControlDate,TRUE);
TControlDate := TOB.Create('Les periodes', NIL, -1);
TControlDate.LoadDetailDB('ControlDate','','',QRechControlDate,False);
Ferme(QRechControlDate);

TControlDateD := TControlDate.FindFirst([''],[''],FALSE);
While TControlDateD <> NIL do
      begin
      DateDebBuf := TControlDateD.GetValue('PDE_DATEDEBUT');
      DateFinBuf := TControlDateD.GetValue('PDE_DATEFIN');
      OrdreBuf := TControlDateD.GetValue('PDE_ORDRE');
      if (StrToInt(GetControlText('ORDREDADS')) <> OrdreBuf) then
         begin
         if ((StrToDate(GetControlText('EDEBUT'))>= DateDebBuf) and
            (StrToDate(GetControlText('EDEBUT'))<= DateFinBuf)) then
            begin
            StDate:= 'La date de début '+GetControlText('EDEBUT')+' est#13#10'+
                     'comprise entre le '+DateToStr(DateDebBuf)+' et le#13#10'+
                     DateToStr(DateFinBuf)+' qui sont les dates de début#13#10'+
                     'et de fin d''une autre période';
            PGIBox (StDate,'Contrôle DADS-U');
            result := FALSE;
            break;
            end;

         if ((StrToDate(GetControlText('EFIN'))>= DateDebBuf) and
            (StrToDate(GetControlText('EFIN'))<= DateFinBuf)) then
            begin
            StDate:= 'La date de fin '+GetControlText('EFIN')+' est#13#10'+
                     'comprise entre le '+DateToStr(DateDebBuf)+' et le#13#10'+
                     DateToStr(DateFinBuf)+' qui sont les dates de début#13#10'+
                     'et de fin d''une autre période';
            PGIBox (StDate,'Contrôle DADS-U');
            result := FALSE;
            break;
            end;

         if ((StrToDate(GetControlText('EDEBUT'))<= DateDebBuf) and
            (StrToDate(GetControlText('EFIN'))>= DateFinBuf)) then
            begin
            StDate:= 'Une autre période d''inactivité contient déjà les#13#10'+
                     ' dates '+GetControlText('EDEBUT')+' et '+
                     GetControlText('EFIN');
            PGIBox (StDate,'Contrôle DADS-U');
            result := FALSE;
            break;
            end;
         end;
      TControlDateD := TControlDate.FindNext([''],[''],FALSE);
      end;

FreeAndNil(TControlDate);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :   /  /
Description .. : Mise à jour des tables
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
function TOF_PG_DADSINACT.UpdateTable(): boolean;
var
Ordre, Rep : integer;
begin
result := FALSE;
Ordre := StrToInt(GetControlText('ORDREDADS'));

Rep:=PGIAsk ('Voulez vous sauvegarder votre saisie ?', 'Saisie des périodes d''inactivité DADS-U') ;
if Rep=mrNo then
   exit
else
   result := TRUE;

try
   begintrans;
   DeleteErreur (Sal, 'S', Ordre);    //PT15-1
   ControleOK := ControleConform;
   if ControleOK = TRUE then
      begin
{PT14
      DeletePeriode(Sal, TypeD, Ordre);
      DeleteDetail(Sal, TypeD, Ordre);
}
      DeletePeriode(Sal, Ordre);
      DeleteDetail (Sal, Ordre);

      ChargeTOBDADS;
      DeleteErreur ('', 'SKO');	//PT13-4
      SauveZones;
      EcrireErreurKO;	//PT13-4
      LibereTOBDADS;

      MAJQuery;

      TFVierge (Ecran).Binsert.Enabled:= True;
      TFVierge (Ecran).BDelete.Enabled:= True;
      end;
   CommitTrans;
Except
   result:= FALSE;
   Rollback;
   PGIBox ('Une erreur est survenue lors de la mise à jour de la base', 'Mise à jour DADS-U');
   END;
end;

{$IFNDEF EAGLCLIENT}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "SalPrem"
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSINACT.SalPremClick(Sender: TObject);
begin
BougeSal(nbFirst) ;
SalPrem.Enabled := FALSE;
SalPrec.Enabled := FALSE;
SalSuiv.Enabled := TRUE;
SalDern.Enabled := TRUE;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "SalPrec"
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSINACT.SalPrecClick(Sender: TObject);
begin
BougeSal(nbPrior) ;
if QMul.BOF then
   begin
   SalPrem.Enabled := FALSE;
   SalPrec.Enabled := FALSE;
   end;
SalSuiv.Enabled := TRUE;
SalDern.Enabled := TRUE;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "SalSuiv"
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSINACT.SalSuivClick(Sender: TObject);
begin
BougeSal(nbNext) ;
SalPrem.Enabled := TRUE;
SalPrec.Enabled := TRUE;
if QMul.EOF then
   begin
   SalSuiv.Enabled := FALSE;
   SalDern.Enabled := FALSE;
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "SalDern"
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSINACT.SalDernClick(Sender: TObject);
begin
BougeSal(nbLast) ;
SalPrem.Enabled := TRUE;
SalPrec.Enabled := TRUE;
SalSuiv.Enabled := FALSE;
SalDern.Enabled := FALSE;
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :   /  /    
Description .. : Déplacement au niveau du salarié
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
Function TOF_PG_DADSINACT.BougeSal(Button: TNavigateBtn) : boolean ;
BEGIN
UpdateTable;
result:=TRUE ;

if Button=nbDelete then
   begin
   if QMul.EOF = FALSE then
      begin
      QMul.Next;
      if QMul.EOF = TRUE then
         begin
         QMul.prior ;
         if QMul.BOF then
            Close;
         end
      end
   else
      begin
      if QMul.BOF = FALSE then
         begin
         QMul.prior;
         if QMul.BOF = TRUE then
            Close;
         end;
      end;
   end;

if QMul <> NIL then
   begin
   Case Button of
        nblast : QMul.Last;
        nbfirst : QMul.First;
        nbnext : QMul.Next;
        nbprior : QMul.prior;
        end;
   end;
GereQuerySal;
END ;


procedure TOF_PG_DADSINACT.GereQuerySal();
begin
if QMul = NIL then
   exit;
Sal := QMul.FindField('PSA_SALARIE').AsString;
if T_Periode <> NIL then
   FreeAndNil (T_Periode);

SalChange := True;
ChargeZones;
AfficheCaption;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "PerPrec"
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSINACT.PerPrecClick(Sender: TObject);
begin
BougePer(nbPrior) ;
Nbre := IntToStr(StrToInt(Nbre)+1);
if QPer.BOF then
   begin
   Nbre := IntToStr(StrToInt(Nbre)-1);
   PerPrec.Enabled := FALSE;
   end;

AfficheCaption;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "PerSuiv"
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSINACT.PerSuivClick(Sender: TObject);
begin
BougePer(nbNext) ;
Nbre := IntToStr(StrToInt(Nbre)-1);
if QPer.EOF then
   begin
   Nbre := IntToStr(StrToInt(Nbre)+1);
   PerSuiv.Enabled := FALSE;
   end;

AfficheCaption;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :   /  /
Description .. : MAJ de la query des périodes
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSINACT.MAJQuery();
var
StPer : string;
QCount : TQuery;
begin
Ferme(QPer);
StPer:= 'SELECT COUNT(PDE_SALARIE) AS NBRE'+
        ' FROM DADSPERIODES WHERE'+
        ' PDE_SALARIE = "'+Sal+'" AND'+
        ' PDE_ORDRE < 0 AND'+
        ' PDE_EXERCICEDADS = "'+PGExercice+'"';

QCount:=OpenSql(StPer, True);
if (not QCount.EOF) then
   NbreTot := IntToStr(-QCount.FindField('NBRE').AsInteger)
else
   NbreTot := '0';
Ferme(QCount);

StPer:= 'SELECT *'+
        ' FROM DADSPERIODES WHERE'+
        ' PDE_SALARIE = "'+Sal+'" AND'+
        ' PDE_ORDRE < 0 AND'+
        ' PDE_EXERCICEDADS = "'+PGExercice+'"'+
        ' ORDER BY PDE_DATEDEBUT';

QPer:=OpenSql(StPer, True);
QPer.First;
Nbre := IntToStr(-1);
{PT5-2
while ((not QPer.Eof) and
      (QPer.FindField('PDE_ORDRE').AsString <> OrdreDADS.Text)) do
}
while ((not QPer.Eof) and
      (IntToStr(QPer.FindField('PDE_ORDRE').AsInteger) <> GetControlText('ORDREDADS'))) do    //DB2
      begin
      QPer.Next;
      Nbre := IntToStr(StrToInt(Nbre)-1);
      end;

if (QPer.EOF) then
   begin
   Nbre := IntToStr(StrToInt(Nbre)+1);
   QPer.First;
{PT5-2
   OrdreDADS.Text := '';
}
{PT11-3
   SetControlText('ORDREDADS', '');
}
   if (not QPer.Eof) then
      SetControlText ('ORDREDADS',
                      IntToStr(QPer.FindField('PDE_ORDRE').AsInteger))
   else
      SetControlText ('ORDREDADS', '0');
   ChargeZones;
   end;
//FIN PT11-3

AfficheCaption;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :   /  /
Description .. : Déplacement au niveau de la période
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
Function TOF_PG_DADSINACT.BougePer(Button: TNavigateBtn) : boolean ;
BEGIN
UpdateTable;
result:=TRUE ;

if QPer <> NIL then
   begin
   Case Button of
        nbfirst : QPer.First;
        nbprior : QPer.prior;
        nbnext : QPer.Next;
        nblast : QPer.Last;
        end;
   end;
GereQueryPer;
END ;


procedure TOF_PG_DADSINACT.GereQueryPer();
begin
if QPer = NIL then
   exit;
if T_Periode <> NIL then
   FreeAndNil (T_Periode);

ChargeZones;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :   /  /
Description .. : Click sur le bouton Imprimer
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSINACT.Impression(Sender: TObject);
var
Pages : TPageControl;
Rech : String;
begin
Rech := 'SELECT *'+
        ' FROM DADSPERIODES WHERE'+
        ' PDE_SALARIE = "'+Sal+'" AND'+
        ' PDE_TYPE = "'+TypeD+'" AND'+
        ' PDE_ORDRE = '+GetControlText('ORDREDADS')+' AND'+
        ' PDE_EXERCICEDADS = "'+PGExercice+'"';

Pages := TPageControl(GetControl('PAGES'));
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/10/2001
Modifié le ... :   /  /
Description .. : Click sur le bouton valider
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSINACT.Validation(Sender: TObject);
begin
UpdateTable;
if (ControleOK = TRUE) then
   begin
   if (State = 'CREATION') then
      begin
      SalPrem.Enabled := TRUE;
      SalPrec.Enabled := TRUE;
      SalSuiv.Enabled := TRUE;
      SalDern.Enabled := TRUE;
      end;
   State := 'MODIFICATION';
   end;
end;


//PT6
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 04/05/2004
Modifié le ... :   /  /
Description .. : Complément des raccourcis claviers
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSINACT.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
TFVierge(Ecran).FormKeyDown(Sender, Key, Shift);
case Key of
     VK_F3: begin
            if ((GetControlVisible ('BSALPREM')) and
               (GetControlEnabled ('BSALPREM')) and (ssCtrl in Shift)) then
               SalPrem.Click      //Premier salarié
            else
               if ((GetControlVisible ('BSALPREC')) and
                  (GetControlEnabled ('BSALPREC')) and (Shift = [])) then
                  SalPrec.Click;     //Salarié précédent
            end;
     VK_F4: begin
            if ((GetControlVisible ('BSALDERN')) and
               (GetControlEnabled ('BSALDERN')) and (ssCtrl in Shift)) then
                SalDern.Click     //Dernier salarié
            else
                if ((GetControlVisible('BSALSUIV')) and
                   (GetControlEnabled('BSALSUIV')) and (Shift = [])) then
                   SalSuiv.Click;     //Salarié suivant
            end;
     ord('P'): if ((GetControlVisible('BPERPREC')) and
                  (GetControlEnabled('BPERPREC')) and (ssAlt in Shift)) then
                  PerPrec.Click; //Période précédente
     ord('N'): if ((GetControlVisible('BPERSUIV')) and
                  (GetControlEnabled('BPERSUIV')) and (ssAlt in Shift)) then
                  PerSuiv.Click; //Période suivante
     end;
end;
//FIN PT6

//PT13-1
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 18/09/2006
Modifié le ... :   /  /
Description .. : Changement du type de régime
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSINACT.RegimeClick(Sender: TObject);
begin
inherited;
case Regime.ItemIndex of
     0 : BEGIN
         SetControlVisible ('GBELEMCOMPL', False);
         SetControlVisible ('GBIRCANTECSITPARTIC', True);
         END;

     1 : BEGIN
         SetControlVisible ('GBELEMCOMPL', True);
         SetControlVisible ('GBIRCANTECSITPARTIC', False);
         END;
     END;
setcontroltext('CMOTIF', ''); // PT18 remise à blanc du code motif
choixmotif;  // PT18 changement de la liste des codes motif
end;
//FIN PT13-1

// Deb  pt18
{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 19/11/2007
Modifié le ... :   /  /    
Description .. : Initialise la liste des motifs en fonction du régime
Mots clefs ... : 
*****************************************************************}
procedure TOF_PG_DADSINACT.choixmotif;
var
stplus, listeplus : string;
begin
inherited;
listeplus := '04,08,09,10,11,13,52';
if Regime.ItemIndex =  0 then
stplus := 'AND CO_ABREGE in ('+listeplus+')'
else
stplus := 'AND CO_ABREGE <> "" and CO_ABREGE <> "52"';

setcontrolproperty('CMOTIF', 'plus' , stplus);
end;
// fin PT18

Initialization
registerclasses([TOF_PG_DADSINACT]) ;

end.
