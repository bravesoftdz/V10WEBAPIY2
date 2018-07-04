{***********UNITE*************************************************
Auteur  ...... : EPI
Créé le ...... : 21/03/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PROCESSUSTRT ()
Mots clefs ... : TOF;PROCESSUSTRT
*****************************************************************}
{
PT1   : 14/04/2006 EPI V_65 FQ 12791 appel automatique des options
                                 arret en cours de traitement
PT2   : 04/05/2006 VG V_65 Gestion multi-attestations par salarié - FQ N°12425
                           et FQ N°12704
PT3   : 15/05/2006 EPI V_65 Test L.G. après un arret en cours de traitement,
                            plus de boucle sur le processus et message à tort
PT4   : 18/05/2006 EPI V_65 FQ N°13151 gestion bouton détail

PT5   : 30/05/2006 EPI V_65 FQ N°13201 ergonomie bouton détail
                            FQ N°13200 idem MUL Salarié exclure les salariés sortis
PT6   : 05/02/2007 FC V_70  FQ N°13349 Ajout de l'action préparation automatique des bulletins
PT7   : 07/03/2007 FC V_70  Edition du bulletin, choisir le mode d'édition duplicata ou specimen
PT8   : 31/03/2007 GGS V_80 FQ 14060 Passage d'argument pour processus sortie de salarié
}
Unit PROCESSUSTRT_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}  mul, FE_Main, DBGrids, Hdb,
{$else}
     eMul,  MainEAGL, UTOB, //@@
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF , UTOM,  HTB97,
     Grids,
     UTofPG_MulAssedic,                 //PT1 attestation assedic
     ParamSoc,PGOutils,EntPaie,PGOutils2;//PT7

Type
  TOF_PROCESSUSTRT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    private
{$IFDEF EAGLCLIENT}
  Grille: THGrid;
{$ELSE}
  Grille: THDBGrid;
{$ENDIF}
    Retour : String;         // PT1
    procedure BTnTrtClick(Sender: TObject);
    procedure ProcesTrtSalarie(Salarie : String);
    procedure ProcesOption(Option,Salarie : String);
    procedure ChangeProc(Sender: Tobject);
    procedure ProcesTrtOption(Proces : String);
    // PT1 ajout bouton pour appel détail du processus
    procedure BTnDetClick(Sender: TObject);
    // PT5
    procedure OnClickSalarieSortie(Sender: TObject);
  end ;

Implementation

uses Buttons;      // Bouton Détail

procedure TOF_PROCESSUSTRT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PROCESSUSTRT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PROCESSUSTRT.OnUpdate ;
begin
  Inherited ;
end ;

// PT5 ajout sélection sur la date de sortie du salarié
procedure TOF_PROCESSUSTRT.OnLoad ;
var
  DateArret, DateDebutPaie, DebExer, FinExer : TdateTime;
  StDateArret,StDateDebutPaie : string;
  Zdate: TDateTime;
  MoisE, AnneeE, ComboExer : string;
begin
  Inherited ;
  StDateArret := '';
  if  TCheckBox(GetControl('CKSORTIE'))<>nil then
  Begin
    if (GetControlText('CKSORTIE')='X') and (IsValidDate(GetControlText('DATEARRET')))then   //DEB PT2
    Begin
      DateArret:=StrtoDate(GetControlText('DATEARRET'));
      StDateArret:=' AND (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) ';
      StDateArret:=StDateArret + ' AND PSA_DATEENTREE <="'+UsDateTime(DateArret)+'"';
    End
  Else
    StDateArret:='';
  end;

  //DEB PT6 - Ne pas tenir compte des salariés dont la date de sortie est < à la date de début période paie
  //          dans le cas du processus de préparation automatique des bulletins
  StDateDebutPaie := '';
  if (GetControlText('COMBOPROC') = 'BULTIN') then
  begin
    RendExerSocialEnCours(MoisE, AnneeE, ComboExer, DebExer, FinExer);
    if VH_Paie.PgDatePrepAuto = 'MOI' then
      DateDebutPaie := FinExer
    else
    begin
      if not VH_Paie.PGDecalage then
        ZDate := EncodeDate(StrToInt(AnneeE), StrToInt(MoisE), 01)
      else
      begin
        if MoisE = '12' then ZDate := EncodeDate(StrToInt(AnneeE) - 1, StrToInt(MoisE), 01)
        else ZDate := EncodeDate(StrToInt(AnneeE), StrToInt(MoisE), 01);
      end;
      DateDebutPaie := FINDEMOIS(Zdate);
    end;
    StDateDebutPaie := ' AND (PSA_DATESORTIE>="'+UsDateTime(DateDebutPaie)+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) ';
    StDateDebutPaie := StDateDebutPaie + ' AND PSA_DATEENTREE <="'+UsDateTime(DateDebutPaie)+'"';
  end;
  //FIN PT6

  SetControlText('XX_WHERE',StDateArret+StDateDebutPaie);
end ;

procedure TOF_PROCESSUSTRT.OnArgument (S : String ) ;
var
  // PT1 BtnTrt : TToolBarButton97;
  BtnTrt : TToolBarButton97;
  BtnDet : TToolBarButton97;       // PT5
  // PT5   BitDet : tBitBtn;
  Trt: THValCombobox;
  Check : TCheckBox;             // PT5

begin
  Inherited ;
{$IFDEF EAGLCLIENT}
  Grille := THGrid(GetControl('Fliste'));
{$ELSE}
  Grille := THDBGrid(GetControl('Fliste'));
{$ENDIF}

  // PT5 début
  SetControlVisible('CKSORTIE',True);
  SetControlvisible('DATEARRET',True);
  SetControlvisible('TDATEARRET',True);
  SetControlEnabled('DATEARRET',False);
  SetControlEnabled('TDATEARRET',False);
  // PT5 fin

  BtnTrt := TToolBarButton97(GetControl('BOuvrir'));
  if BtnTrt <> nil then BtnTrt.OnClick := BTnTrtClick;

  SetControlVisible('BDetail', False);        // PT4
  // PT5 BitDet := TBitBtn(GetControl('BDetail'));
  // PT5 if BitDet <> nil then BitDet.OnClick := BTnDetClick;
  BtnDet := TToolBarButton97(GetControl('BDetail'));
  if BtnDet <> nil then BtnDet.OnClick := BtnDetClick;

  Trt := THValComboBox(GetControl('COMBOPROC'));
   if Trt <> nil then
   begin
     Trt.OnChange := ChangeProc;
   end;

   // PT5 début
   Check := TCheckBox(GetControl('CKSORTIE'));
   if Check = nil then
     Begin
      SetControlVisible('DATEARRET',False);
      SetControlVisible('TDATEARRET',False);
    End
    else
      Check.OnClick:=OnClickSalarieSortie;
   // PT5 fin
end ;

procedure TOF_PROCESSUSTRT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PROCESSUSTRT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PROCESSUSTRT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_PROCESSUSTRT.ChangeProc(Sender: Tobject);
var
  Trt,Ordre, Proces : String;
  Q : TQuery;

begin
  // Alimentation du type de traitement
  Proces := GetControlText('COMBOPROC');
  If Proces <> '' then
  begin
    SetControlVisible('BDetail', True);        // PT4
    Ordre := 'SELECT PRO_PROSELECT FROM PROCESSUS WHERE PRO_CODEPROC="';
 	  Ordre := Ordre + Proces + '"';
    Q:= OpenSQL(Ordre,TRUE) ;
    if not q.eof then
    begin
      Trt := Q.FindField('PRO_PROSELECT').AsString;
    end;
    Ferme(Q) ;
    SetControlText('COMBOTRT',Trt);
  end;
end;

//  Gestion du traitement
procedure TOF_PROCESSUSTRT.BTnTrtClick(Sender: TObject);
var
  Trt,Proces,Sel_salarie,Txt : String;
  i : integer;

begin
  Inherited ;
  if not (Ecran is TFMul) then exit;

  Proces := GetControlText('COMBOPROC');
  if (Proces = '') then
   begin
     MessageAlerte('Vous devez sélectionner un processus');
     exit;
   end;

   Trt := GetControlText('COMBOTRT');
   If Trt = '001' then
   begin
     // Traitement sans sélection
     if (Grille.NbSelected <> 0) then
	   begin
  	   MessageAlerte('Sélection de salarié inutile pour ce traitement');
       Grille.ClearSelected;
    	 exit;
     end
   end;
   If Trt <> '001' then
     begin
		   if (Grille.NbSelected = 0) and (not Grille.AllSelected) then
		   begin
    		 MessageAlerte('Sélection de salariés obligatoire');
		     exit;
  	   end;
     end;

   // Gestion des 3 types de traitement

   // Code 001 boucle sur les traitements sans sélection
   If Trt = '001' then
   begin
     Sel_Salarie := '';
   	 ProcesTrtSalarie(Sel_Salarie);
   end;

   // Code 002 boucle sur les traitements puis les salariés
   If Trt = '002' then
   begin
     ProcesTrtOption(Proces);
   end;

   // Code 003 boucle sur les salariés puis les traitements
   If Trt = '003' then
   begin
	   for i := 0 to Grille.NbSelected-1 do
  	 begin
	     Grille.GotoLeBookmark(i);
  	  {$IFDEF EAGLCLIENT}
    	 TFMul(Ecran).Q.TQ.Seek(Grille.Row-1) ;
	    {$ENDIF}
  	   Sel_Salarie := TFMul(Ecran).Q.FindField('PSA_SALARIE').asstring;
    	 ProcesTrtSalarie(Sel_Salarie);
       // PT1
       If Retour = 'STOP' then
       begin
        Grille.ClearSelected;
        Break;
       end;
	   end;
  	 Grille.ClearSelected;
   end;
   Txt := RechDom('PGPROCESSUS', GetControlText('COMBOPROC'), False);
   PGModeEdition := ''; //PT7
   if Retour <> 'STOP'
   then
     PgiBox('Fin du processus ' + Txt,Ecran.Caption)
   else
     PgiBox ('Le processus ' + Txt + ' a été interrompu.', Ecran.Caption);
     Retour := '';      // PT3
end;

procedure TOF_PROCESSUSTRT.ProcesTrtSalarie(Salarie : String);
var
  Ordre,Trt,Proces : String;
  QTrt : TQuery;

begin
  // Recherche des options à enchainer
  Proces := GetControlText('COMBOPROC');
  Ordre := 'SELECT PRD_PROCESACTION FROM PROCESSUSDET WHERE ';
	Ordre := Ordre + ' PRD_CODEPROC="'+ Proces + '"';
  Ordre := Ordre + ' ORDER BY PRD_CODEPROC,PRD_NOORDRE ASC';
	QTrt :=  OpenSQL(Ordre,TRUE);
	 while not QTrt.eof do
 	 begin
     Trt := QTrt.FindField('PRD_PROCESACTION').AsString;
     // Appel du traitement option
     ProcesOption(Trt,Salarie);
     if Retour = 'STOP' then Break;
     QTrt.NEXT ;
   end;
 	 Ferme(QTrt);
end;

procedure TOF_PROCESSUSTRT.ProcesTrtOption(Proces : String);
var
Ordre,Trt,Salarie : String;
QTrt : TQuery;
i : integer;
begin
// Recherche des options à enchainer
Ordre:= 'SELECT PRD_PROCESACTION FROM PROCESSUSDET WHERE ';
Ordre:= Ordre+' PRD_CODEPROC="'+Proces+'"';
Ordre:= Ordre+' ORDER BY PRD_CODEPROC,PRD_NOORDRE ASC';
QTrt:= OpenSQL(Ordre,TRUE);
while not QTrt.eof do
      begin
      Trt:= QTrt.FindField('PRD_PROCESACTION').AsString;
      for i:= 0 to Grille.NbSelected-1 do
  	  begin
	  Grille.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
    	  TFMul(Ecran).Q.TQ.Seek(Grille.Row-1) ;
{$ENDIF}
 	  Salarie:= TFMul(Ecran).Q.FindField('PSA_SALARIE').asstring;
{PT2
//PT1 Recherche si attestation déjà faite
          Attes := TFMul(Ecran).Q.FindField('PSA_ASSEDIC').AsString;
}
          ProcesOption(Trt,Salarie);
// PT1
          If Retour = 'STOP' then
             begin
             Grille.ClearSelected;
             Break;
             end;
          end;
      QTrt.NEXT ;
      end;
Ferme(QTrt);
Grille.ClearSelected;
end;

procedure TOF_PROCESSUSTRT.ProcesOption(Option,Salarie : String);
var
Saltrt, StSal : String;
Attes, Ordre : String;   // PT1
QRechSal: TQuery; // PT1
Plusieurs : boolean;
begin
Saltrt := '';

If Salarie <> '' then
   begin
   If option <> 'E03' then
      Saltrt:= 'PSA_SALARIE="'+Salarie+'"'
   else
      Saltrt:= 'PCI_SALARIE="'+Salarie+'"';
   end;

// Saisie salarié si salarié renseigné affichage fiche sinon liste
If Option = 'E01' then
   Begin
   If Salarie = '' then
      AglLanceFiche ('PAY', 'SALARIE_MUL', 'GRILLE=S', '', 'S')
   else
      AGLLanceFiche('PAY', 'SALARIE', '', Salarie, '');
   End;

// Traitement DUE
If Option = 'E02' then
   Begin
// Modif appel pour passer le salarié en paramètre
// PT1 AglLanceFiche('PAY', 'SALARIE_ATTES', 'OBJET=DUE', '', 'DUE;'+ Saltrt);
// PT1 début modification appel
   If Salarie = '' then
// PT1 AglLanceFiche('PAY', 'SALARIE_ATTES', 'OBJET=DUE', '', 'DUE;P;')
      Retour:= AglLanceFiche('PAY', 'SALARIE_ATTES', 'OBJET=DUE', '', 'DUE;P;')
   else
      AglLanceFiche ('PAY','DUE','','',Salarie + ';MODIFICATION');
   End;
// PT1 fin

// Contrat de travail
If Option = 'E03' then
// PT1 AglLanceFiche('PAY', 'SALARIE_RTF', 'OBJET=CONTRAT', '', 'CONTRAT;'+Saltrt);
   Retour := AglLanceFiche('PAY', 'SALARIE_RTF', 'OBJET=CONTRAT', '', 'CONTRAT;P;'+Saltrt);

// Attestation ASSEDIC
If Option = 'S01' then
   Begin
// PT1 Modification appel attestation Assedic pour arriver en saisie
// PT1 AglLanceFiche('PAY', 'MUL_ATTESTASSED', '', '', Saltrt);
// modif appel attestation pour passer le salarié en paramètre
   If Salarie = '' then
// PT1 AglLanceFiche('PAY', 'MUL_ATTESTASSED', '', '', Saltrt)
      Retour:= AglLanceFiche('PAY', 'MUL_ATTESTASSED', '', '', 'P')
   else
// PT1 début appel attestation sans passer par la liste
      begin
      PGAttesSalarie:= Salarie;
//Recherche si attestation déjà faite
       Attes := TFMul(Ecran).Q.FindField('PSA_ASSEDIC').AsString;  //PT2
      if Attes = 'X' then
         begin
//PT2
         StSal:= 'SELECT COUNT(PAS_SALARIE) AS NBRE'+
                 ' FROM ATTESTATIONS WHERE'+
                 ' PAS_SALARIE = "'+Salarie+'" AND'+
                 ' PAS_TYPEATTEST = "ASS"';
         QRechSal:= OpenSQL (StSal, TRUE);
         Plusieurs:= QRechSal.FindField('NBRE').AsInteger>1;
         Ferme (QRechSal);
//FIN PT2
         QRechSal:= OpenSQL ('SELECT PAS_ORDRE FROM ATTESTATIONS WHERE'+
                             ' PAS_SALARIE = "'+Salarie+'" AND'+
                             ' PAS_TYPEATTEST = "ASS"'+
                             ' ORDER BY PAS_DERNIERJOUR DESC', TRUE);
         Ordre := IntToStr(QRechSal.FindField('PAS_ORDRE').AsInteger);
         Ferme(QRechSal);
{$IFNDEF EAGLCLIENT}
{PT2
         AGLLanceFiche ('PAY', 'ASSEDIC', '', Salarie+';ASS;'+Ordre,
                        'ACTION=MODIFICATION');
}
         if (Plusieurs=True) then
            AGLLanceFiche ('PAY', 'MUL_ATTESTASSED', '', '', 'S')
         else
            AGLLanceFiche ('PAY', 'ASSEDIC', '', Salarie+';ASS;'+Ordre,
                           'ACTION=MODIFICATION');
//FIN PT2
         end
      else
         AGLLanceFiche ('PAY', 'ASSEDIC', '', '', Salarie+';ASS;'+Ordre+
                        ';ACTION=CREATION');
{$ELSE}
{PT2
         AGLLanceFiche ('PAY', 'ASSEDIC_EAGL', '', Salarie+';ASS;'+Ordre,
                        'ACTION=MODIFICATION');
}
         if (Plusieurs=True) then
            AGLLanceFiche ('PAY', 'MUL_ATTESTASSED', '', '', 'S')
         else
            AGLLanceFiche ('PAY', 'ASSEDIC_EAGL', '', Salarie+';ASS;'+Ordre,
                           'ACTION=MODIFICATION');
//FIN PT2
         end
      else
         AGLLanceFiche ('PAY', 'ASSEDIC_EAGL', '', '', Salarie+';ASS;'+Ordre+
                        ';ACTION=CREATION');
{$ENDIF}
// PT1 fin
      End;
   End;

// Certificat de travail
If Option = 'S02' then
// PT1 AglLanceFiche('PAY', 'SALARIE_RTF', 'OBJET=CERTIFICAT', '', 'CERTIFICAT;'+Saltrt);
   Retour:= AglLanceFiche('PAY', 'SALARIE_RTF', 'OBJET=CERTIFICAT', '', 'CERTIFICAT;P;'+Saltrt);

// Solde de tout compte
If Option = 'S03' then
// PT1 AglLanceFiche('PAY', 'SALARIE_RTF', 'OBJET=SOLDE', '', 'SOLDE;'+Saltrt);
   Retour:= AglLanceFiche('PAY', 'SALARIE_RTF', 'OBJET=SOLDE', '', 'SOLDE;P;'+Saltrt);

// Saisie date de sortie salarié si sélection affichage fiche sinon liste
If Option = 'S04' then
   Begin
   If Salarie = '' then
      AglLanceFiche('PAY', 'SALARIE_MUL', 'GRILLE=S', '', 'S')
   else
      AGLLanceFiche('PAY', 'SALARIE', '', Salarie, 'S04');  // PT8 'S04' pour affichage direct onglet emploi de fiche salarié
   End;

// Calcul bulletin
If Option = 'S05' then
   Retour:= AglLanceFiche('PAY', 'MUL_BULLETIN', '', '', 'P;'+Salarie);

// Edition bulletin
If Option = 'S06' then
    //DEB PT7
    // Si on ne gère pas le mode d'édition
    if not GetParamSocSecur('SO_PGGESTORIDUPSPE',False) then
      AglLanceFiche('PAY', 'EDITBUL_ETAT', '', '', 'MENU;'+Salarie)
    else
    begin 
      // Vérifier que l'utilisateur a au moins le droit d'éditer un original, un duplicata ou un specimen
      if JaiLeDroitTag(42313) or JaiLeDroitTag(42314) or JaiLeDroitTag(42316) then
      begin
        if PGModeEdition = '' then
        begin
          AglLanceFiche('PAY', 'MODEEDT_BUL', '', '','PROCESSUSTRT');
          if PGModeEdition <> '' then
            AglLanceFiche('PAY', 'EDITBUL_ETAT', '', '', 'MENU;'+Salarie + ';' + PGModeEdition);
        end
        else
          AglLanceFiche('PAY', 'EDITBUL_ETAT', '', '', 'MENU;'+Salarie + ';' + PGModeEdition);
      end
      else
        PGIBox(TraduireMemoire('Vous n''êtes pas autorisé à lancer une édition de bulletins'));
    end;
    //FIN PT7

//DEB PT6
// Préparation automatique des bulletins
If Option = 'S07' then
   AglLanceFiche('PAY', 'BULLETIN_MUL', '', '', ';'+Salarie);
//FIN PT6
end;

// PT1 Ajout icone appel du détail processus
procedure TOF_PROCESSUSTRT.BTnDetClick(Sender: TObject);
var
  Lib,St,Ordre,Proces,Pred,Nodos,ProType : String;
  Q : TQuery;

begin
  Inherited ;
  if not (Ecran is TFMul) then exit;
  Proces := GetControlText('COMBOPROC');
  if (Proces <> '') then
  begin
    Ordre := 'SELECT PRO_PREDEFINI,PRO_NODOSSIER,PRO_PROCESTYPE,PRO_LIBELLEPROC ';
    Ordre := Ordre + 'FROM PROCESSUS WHERE PRO_CODEPROC="';
   	Ordre := Ordre + Proces + '"';
    Q:= OpenSQL(Ordre,TRUE) ;
    if not q.eof then
    begin
      Pred := Q.FindField('PRO_PREDEFINI').AsString;
      Nodos := Q.FindField('PRO_NODOSSIER').AsString;
      ProType := Q.FindField('PRO_PROCESTYPE').AsString;
      Lib := Q.FindField('PRO_LIBELLEPROC').AsString;
    end;
    Ferme(Q) ;
    St :=   Pred + ';' + Nodos + ';' + Proces + ';' + Protype + ';'+ Lib + ';';
    AGLLanceFiche('PAY','PROCESSUSDET_MUL','','',St);
    End;
  End;

// PT5
procedure TOF_PROCESSUSTRT.OnClickSalarieSortie(Sender: TObject);
  begin
    SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
    SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
  end;

Initialization
  registerclasses ( [ TOF_PROCESSUSTRT ] ) ;
end.
