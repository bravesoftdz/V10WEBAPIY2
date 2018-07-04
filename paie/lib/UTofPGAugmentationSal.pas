{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 16/12/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGAUGMENTATIONSAL ()
Mots clefs ... : TOF;PGAUGMENTATIONSAL
*****************************************************************
PT1  | 28/02/2008 | V_800 | FL | Le responsable prime doit pouvoir changer l'état de l'augmentation
}
Unit UTofPGAugmentationSal ;                       

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Graphics,
     HTB97,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     FE_Main,
     Fiche,
     EdtREtat,
{$else}
     eMul,
     MainEAGL,
     UTilEAGL,
{$ENDIF}
     forms,
     uTob,
     sysutils,
     ComCtrls,
     ParamSoc,
     HCtrls,
     HEnt1,
     Vierge,
     HMsgBox,
     UTofPGMul_Augmentation,
     UTobDebug,
     HSysMenu,
     Grids,
     ed_tools,
     P5Util,
     ExtCtrls,
     Menus,
     ShellAPI,
     HPanel,
     windows,
     EntPaie,
     MailOl,
     UTOF ;

Const
{     ColSal = 0;
     ColNom = 1;
     ColEmploi = 2;
     ColDateE = 3;
     ColEtat = 4;
     ColF = 5;
     ColPctF = 6;
     ColFap = 7;
     ColV = 8;
     ColPctV = 9;
     ColVap = 10;
     ColCommentaire = 11;
     ColBrut = 12;
     ColPctB = 13;
     ColBrutap = 14;
     ColMotif = 15;
     ColEtatBis = 16;
     ColMotifBis =17;
     ColAge = 18;
     ColAnc = 19;
     ColEtab = 20;
     ColEff = 21;
     ColHor = 22;
     ColTN1= 23;
     ColTN2 = 24;
     ColTN3 = 25;
     ColTN4 = 26;
     ColStat = 27;
}
     ColSal = 1;
     ColNom = 2;
     ColEmploi = 3;
     ColTpsPart = 4;
     ColDateE = 5;
     ColEtat = 6;
     ColF = 7;
     ColPctF = 8;
     ColFap = 9;
     ColV = 10;
     ColPctV = 11;
     ColVap = 12;
     ColCommentaire = 13;
     ColBrut = 14;
     ColPctB = 15;
     ColBrutap = 16;
     ColMotif = 17;
     ColEtatBis = 18;
     ColMotifBis =19;
     ColAge = 20;
     ColAnc = 21;
     ColEtab = 22;
     ColEff = 23;
     ColHor = 24;
     ColTN1= 25;
     ColTN2 = 26;
     ColTN3 = 27;
     ColTN4 = 28;
     ColStat = 29;
     ColTpsPlein = 30;
     ColSalAnnAv = 31;
     ColSalAnnAp = 32;
     ColSalAnnTpsP = 33;



Type
  TOF_PGAUGMENTATIONSAL = Class (TOF)
 //   procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnClose                   ; override ;
    procedure OnArgument (S : String ) ; override ;

    private
    GAug : THGrid;
    AnneeSaisie,TypeSaisie,StChampGrid : String;
    GestionSalFixe,GestionSalVariable,NewSaisie : Boolean;
    TobTestModif : Tob;
    IArrondiAugm,PrecisionArrondi,PctAugmDec : Integer;
    LastColEnter,LastColExit,NbMEquivalenceAnn : Integer;
    Consultation : Boolean;
    FormatPct : String;
    procedure GrilleCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GrillePostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
    Function FormatageDouble(Chaine : String) : Double;
    procedure VoirEmploi(Sender : TObject);
    procedure CalculerResume(Ligne : Integer;Init : Boolean = False);
    procedure BSimuClick(Sender : Tobject);
    procedure MiseEnFormeGrille;
    procedure ChargerDonnees;
    procedure DupliquerSalaire(PctFixe,PctVar : Double;MotifAug : String);
//    procedure NonSaisie(const GS: THGrid; const Acol, Arow: integer; Canvas: Tcanvas);
    procedure GrilleRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GrilleRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GrilleGetCellCanvas( ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
    procedure GrilleDblClick(Sender : TObject);
    procedure VoirHistoAug(Sender : TObject);
    procedure AccesSal(Sender : Tobject);
    Procedure BSelectClick(Sender : Tobject);
    Procedure AfficheInfosSal(Ou : Integer);
    Procedure ValiderEtat(Sender : Tobject);
    Procedure RefuserEtat(Sender : Tobject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure ZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: boolean);
    function ZoneAccessible(var ACol, ARow: Longint): Boolean;
    procedure VerifierModifs(Champ : String;MontantVal : Double;Ligne : Integer);
    procedure ValidationSaisie(Sender : TObject);
    Function ChampSalaires(TypeSalaire : String) : String;
    procedure EntrerGrille(Sender : TObject);
    procedure ImprimerGrille(Sender : Tobject);
    Procedure ExporterGrille(Sender : TObject);
    Procedure AnnulerModifs(Sender : TObject);
    procedure AfficheLegende(Sender : TObject);
    Function ArrondiAugm(Montant : Double) : Double;
    {$IFDEF EMANAGER}
    Procedure EnvoyerMailResp(Sender : TObject);
    {$ENDIF}
    procedure VoirTempsPlein(Sender : TObject);
  end ;
  var TobEmploiAug : Tob;

Implementation

uses
PgOutilsGrids;

procedure TOF_PGAUGMENTATIONSAL.OnLoad ;
var HMTRAD : THSystemMenu;
begin
  Inherited ;

            HMTrad:=THSystemMenu(GetControl('HMTrad'));
//           TFVierge(Ecran).WindowState := wsMaximized;
           HMTrad.ResizeGridColumns(GAug) ;
end;

procedure TOF_PGAUGMENTATIONSAL.OnClose ;
var TestModif : Boolean;
    i : Integer;
    Rep : Word;
begin
  Inherited ;
  {$IFDEF EMANAGER}
  If Consultation then
  begin
       TobTestModif.Free;
       Exit;
  end;
  {$ENDIF}
  TobAug.GetGridDetail (GAug,GAug.RowCount-1,'','INDICATEUR;'+ StChampGrid);
  TobAug.Detail.Sort('PSA_SALARIE');
  TobTestModif.Detail.Sort('PSA_SALARIE');
  TestModif := False;
  For i := 0 to TobAug.Detail.Count -1 do
  begin
       If TobTestModif.detail[i].getValue('PBG_PCTFIXE') <> TobAug.detail[i].GetValue('PBG_PCTFIXE') then TestModif := True;
       If TobTestModif.detail[i].getValue('PBG_PCTVARIABLE') <> TobAug.detail[i].GetValue('PBG_PCTVARIABLE') then TestModif := True;
       If TobTestModif.detail[i].getValue('ETAT') <> TobAug.detail[i].GetValue('ETAT') then TestModif := True;
       If TobTestModif.detail[i].getValue('MOTIF') <> TobAug.detail[i].GetValue('MOTIF') then TestModif := True;
       If TobTestModif.detail[i].getValue('PBG_COMMENTAIREABR') <> TobAug.detail[i].GetValue('PBG_COMMENTAIREABR') then TestModif := True;
       If testModif = True then Break;
  end;
  If testModif = True then
  begin
       Rep := PGIAskcancel('Voulez-vous sauvegarder vos modifications ?',Ecran.Caption);
       If Rep = MrYes then ValidationSaisie(Nil);
       If Rep = MrCancel then Lasterror := 1;
       If Rep = MrNo then TobTestModif.Free;
  end
  else TobTestModif.Free;
end;


//procedure TOF_PGAUGMENTATIONSAL.OnUpdate ;
procedure TOF_PGAUGMENTATIONSAL.ValidationSaisie(Sender : TObject);
var TobMaj,TM : Tob;
    i : Integer;
    Q : TQuery;
    Salarie,Commentaire : String;
    BTemps : TToolBarButton97;
    Etat : String; //PT1
begin
  BTemps := TToolBarButton97(GetControl('BTEMPS'));
  If BTemps <> Nil then
  begin
       If BTemps.Down = True then
       begin
            BTemps.Down := False;
            VoirTempsPlein(BTemps);
       end;
  end;
  ValideComboDansGrid (GAug, GAug.Col, 4, 1);
  Commentaire := GetControlText('COMMENTAIRE');
  GAug.CellValues[ColCommentaire,GAug.Row] := Commentaire;
  NextPrevControl(TFVierge(Ecran));
  SetFocusControl('COMMENTAIRE');
  TobAug.GetGridDetail (GAug,GAug.RowCount-1,'','INDICATEUR;'+ StChampGrid);
  InitMoveProgressForm (NIL,'Sauvegarde en cours',
                    'Veuillez patienter SVP ...', TobAug.Detail.Count,
                    False,True);
  Gaug.ColFormats[ColEtat] := '';
  Gaug.ColFormats[ColMotif] := '';
  TobMaj := Tob.Create('BUDGETPAIE',Nil,-1);
  Q := OpenSQL('SELECT * FROM BUDGETPAIE WHERE PBG_ANNEE="'+AnneeSaisie+'"',True);
  TobMaj.LoadDetailDB('BUDGETPAIE','','',Q,False);
  Ferme(Q);
  For i := 0 to TobAug.Detail.Count - 1 do
  begin
        Salarie := TobAug.Detail[i].GetValue('PSA_SALARIE');
        TM := TobMaj.FindFirst(['PBG_SALARIE','PBG_ANNEE'],[Salarie,AnneeSaisie],False);
        If TM = Nil then TM := Tob.Create('BUDGETPAIE',TobMaj,-1);
        TM.PutValue('PBG_SALARIE',Salarie);
        TM.PutValue('PBG_LIBELLE',TobAug.Detail[i].GetValue('PSA_LIBELLE'));
//        TM.PutValue('PBG_LIBELLEEMPLOI',TobAug.Detail[i].GetValue('PBG_LIBELLEEMPLOI'));
        TM.PutValue('PBG_FIXEAV',TobAug.Detail[i].GetValue('PBG_FIXEAV'));
        TM.PutValue('PBG_FIXEAP',TobAug.Detail[i].GetValue('PBG_FIXEAP'));
        TM.PutValue('PBG_PCTFIXE',TobAug.Detail[i].GetValue('PBG_PCTFIXE'));
        TM.PutValue('PBG_VARIABLEAV',TobAug.Detail[i].GetValue('PBG_VARIABLEAV'));
        TM.PutValue('PBG_VARIABLEAP',TobAug.Detail[i].GetValue('PBG_VARIABLEAP'));
        TM.PutValue('PBG_PCTVARIABLE',TobAug.Detail[i].GetValue('PBG_PCTVARIABLE'));
        TM.PutValue('PBG_COMMENTAIREABR',TobAug.Detail[i].GetValue('PBG_COMMENTAIREABR'));
        TM.PutValue('PBG_MOTIFAUGM',TobAug.Detail[i].GetValue('PBG_MOTIFAUGM'));
        TobAug.Detail[i].PutValue('PBG_MOTIFAUGM',TobAug.Detail[i].GetValue('PBG_MOTIFAUGM'));
        If ((TypeSaisie = 'SAISIE') Or (TypeSaisie = 'PROPOSITION')) then //PT1
        begin
             TM.PutValue('PBG_PCTFIXESAISIE',TobAug.Detail[i].GetValue('PBG_PCTFIXE'));
             TM.PutValue('PBG_PCTVARSAISIE',TobAug.Detail[i].GetValue('PBG_PCTVARIABLE'));
             TM.PutValue('PBG_DATESAISIE',Date);
        end;
        If TypeSaisie= 'VALIDRESP' then
        begin
             TM.PutValue('PBG_PCTFIXERESP',TobAug.Detail[i].GetValue('PBG_PCTFIXE'));
             TM.PutValue('PBG_PCTVARRESP',TobAug.Detail[i].GetValue('PBG_PCTVARIABLE'));
             TM.PutValue('PBG_VALIDELE',Date);
             {$IFDEF EMANAGER}
             TM.PutValue('PBG_VALIDERPAR',V_PGI.UserSalarie);
             {$ENDIF}
        end;
        If TypeSaisie= 'VALIDDRH' then
        begin
             TM.PutValue('PBG_ACCEPTELE',Date);
        end;
        If ((TypeSaisie = 'SAISIE') Or (TypeSaisie = 'PROPOSITION')) then //PT1
        begin
             If Consultation then
             begin
                  if TobAug.Detail[i].GetValue('ETAT') = '001' then
                  begin
                    TM.PutValue('PBG_ETATINTAUGM','000');
                    TobAug.Detail[i].PutValue('PBG_ETATINTAUGM','000');
                    TobAug.Detail[i].PutValue('ETAT','000');
                  end;
             end
             else
             begin
               If (TobAug.Detail[i].GetValue('ETAT') = '001') or (TobAug.Detail[i].GetValue('ETAT') = '000') then
               begin
                  //PT1 - Début
(*                  If TypeSaisie = 'SAISIE' Then
                    Etat := TobAug.Detail[i].GetValue('PBG_ETATINTAUGM')
                  Else
                    Etat := '002';
                  TM.PutValue('PBG_ETATINTAUGM', Etat);
                  TobAug.Detail[i].PutValue('PBG_ETATINTAUGM',Etat);
                  TobAug.Detail[i].PutValue('ETAT',Etat);*)
                  //PT1 - Fin
                  TM.PutValue('PBG_ETATINTAUGM', '002');
                  TobAug.Detail[i].PutValue('PBG_ETATINTAUGM','002');
                  TobAug.Detail[i].PutValue('ETAT','002');
               end
               else
               begin
                  TM.PutValue('PBG_ETATINTAUGM',TobAug.Detail[i].GetValue('PBG_ETATINTAUGM'));
                  TobAug.Detail[i].PutValue('PBG_ETATINTAUGM',TobAug.Detail[i].GetValue('PBG_ETATINTAUGM'));
                  TobAug.Detail[i].PutValue('ETAT',TobAug.Detail[i].GetValue('PBG_ETATINTAUGM'));
               end;
             end;
        end
        else
        If TobAug.Detail[i].GetValue('PBG_ETATINTAUGM') <> '' then
        begin
             TM.PutValue('PBG_ETATINTAUGM',TobAug.Detail[i].GetValue('PBG_ETATINTAUGM'));
              TobAug.Detail[i].PutValue('PBG_ETATINTAUGM',TobAug.Detail[i].GetValue('PBG_ETATINTAUGM'));
              TobAug.Detail[i].PutValue('ETAT',TobAug.Detail[i].GetValue('PBG_ETATINTAUGM'));
        end;
        TM.PutValue('PBG_NUMORDRE',1);
        TM.PutValue('PBG_TYPEBUDG','AUG');
        TM.PutValue('PBG_ANNEE',AnneeSaisie);
        TM.InsertOrUpdateDB();
        MoveCurProgressForm ('Salarié : '+
                       TobAug.Detail[i].GetValue('PSA_LIBELLE'));
  end;
  TFVierge(Ecran).Retour := 'MODIF';
//  If TypeSaisie = 'VALIDRESP' then Gaug.ColFormats[4] := 'CB=PGAUGMETATVALID2'
//  else Gaug.ColFormats[4] := 'CB=PGAUGMETATVALID';
  If TypeSaisie = 'SAISIE' Then
    Gaug.ColFormats[ColEtat] := 'CB=PGAUGMETATVALID1' //PT1
  Else
    Gaug.ColFormats[ColEtat] := 'CB=PGAUGMETATVALID';
  Gaug.ColFormats[ColMotif] := 'CB=PGMOTIFAUGM';
  TobAug.PutGridDetail(GAug,False,False,StChampGrid,False);
  GAug.SortedCol := -1;
  TobTestModif.ClearDetail;
  TobTestModif.Dupliquer(TobAug,True,True,True);
  FiniMoveProgressForm;
end ;

procedure TOF_PGAUGMENTATIONSAL.OnArgument (S : String ) ;
var BHisto,BValider,BSimu,BtnSal,BEmploi,BSelect,BValid,BRefus: TToolBarButton97;
    i : Integer;
    BTemps,BImp,BExport,BDefaire,BLegende : TToolBarButton97;
    CodeArrondi : String;
    Q : TQuery;
    {$IFDEF EMANAGER}
    BMail : TToolBarButton97;
    {$ENDIF}
    Titres : HTStringList;
    NbDecPct : Integer;
begin
  Inherited ;
     NbDecPct := GetparamSocSecur('SO_PGAUGMPCTDEC',False);
     FormatPct := '# ##0.';
     For i := 1 to NbDecPct do
     begin
          FormatPct := FormatPct + '0';
     end;
        For i := 1 to 4 do
     begin
          SetControlCaption('TRAVAILN'+IntToStr(i),'');
     end;    
        SetControlVisible('LIBA',True);
        SetControlVisible('LIBFXAV1',True);
        SetControlVisible('LIBFXAP1',True);
        SetControlVisible('LIBFXPCT1',True);
        SetControlVisible('ANMTFXAV',True);
        SetControlVisible('ANMTFXAP',True);
        SetControlVisible('ANNMTFXPCT',True);
        SetControlVisible('ANMTVARAV',True);
        SetControlVisible('ANMTVARAP',True);
        SetControlVisible('ANNMTVARPCT',True);
        SetControlVisible('ANMTTOTAV',True);
        SetControlVisible('ANMTTOTAP',True);
        SetControlVisible('ANNMTTOTPCT',True);
//        SearchTimer.Interval := 50 ;
//        SearchTimer.Enabled:=TRUE ;
        NbMEquivalenceAnn := GetParamSocSecur('SO_PGAUGMNBMOISANN',False);
        If NbMEquivalenceAnn <= 0 then NbMEquivalenceAnn := 12;
        BTemps := TToolBarButton97(GetControl('BTEMPS'));
        If BTemps <> Nil then BTemps.ONClick := VoirTempsPlein;
        LastColEnter := -1;
        LastColExit := -1;
        TFVierge(Ecran).WindowState := wsMaximized;
        TFVierge(Ecran).Align := AlClient;
        NewSaisie := True;
        CodeArrondi := GetParamSoc('SO_PGAUGMSALARR');
        PctAugmDec := GetParamSoc('SO_PGAUGMPCTDEC');
        Q := OpenSQL('SELECT * FROM COMMUN WHERE CO_TYPE="PRR" AND CO_CODE="'+CodeArrondi+'"',True);
        If Not Q.Eof then
        begin
             IArrondiAugm := StRtoInt(Q.FindField('CO_ABREGE').AsString);
             PrecisionArrondi := StRtoInt(Q.FindField('CO_LIBRE').AsString);
        end
        else
        begin
             IArrondiAugm := 0;
             PrecisionArrondi := 0;
        end;
        Ferme(Q);
        StChampGrid := 'PSA_SALARIE;PSA_LIBELLE;PBG_LIBELLEEMPLOI;TPSPART;PSA_DATEENTREE;PBG_ETATINTAUGM;PBG_FIXEAV;'+
        'PBG_PCTFIXE;PBG_FIXEAP;PBG_VARIABLEAV;PBG_PCTVARIABLE;PBG_VARIABLEAP;PBG_COMMENTAIREABR;TOTALAV;PCTTOTAL;TOTALAP;'+
        'PBG_MOTIFAUGM;ETAT;MOTIF;AGE;ANCIENETE;PSA_ETABLISSEMENT;EFFECTIF;HORAIRE;PSA_TRAVAILN1;PSA_TRAVAILN2;PSA_TRAVAILN3;PSA_TRAVAILN4;PSA_CODESTAT;TPSPLEIN;SALANNAV;SALANNAP;SALANNTPSP';
        TobTestModif := Tob.Create('TesterModif',Nil,-1);
        TobTestModif.Dupliquer(TobAug,True,True,True);
        BValider := TToolBarButton97(GetControl('VALIDSAISIE'));
        BValider.OnClick := ValidationSaisie;
        GestionSalFixe := (GetParamSoc('SO_PGAUGFIXE') <> '');
        GestionSalVariable := (GetParamSoc('SO_PGAUGVARIABLE') <> '');
        TFVierge(Ecran).Retour := '';
        TFVierge(Ecran).OnKeyDown := FormKeyDown;
        TypeSaisie := readTokenPipe(S,';');
        AnneeSaisie := readTokenPipe(S,';');
        Consultation := False;
        If TypeSaisie = 'PROPOSITION' then
        begin
             TFVierge(Ecran).Caption := 'Propositions d''augmentations';
             //UpdateCaption(TFVierge(Ecran)); //PT1
             //TypeSaisie := 'SAISIE'; //PT1
             Consultation := True;
        end
        Else
            Ecran.Caption := 'Saisie des augmentations de l''année '+ AnneeSaisie;
        UpdateCaption(Ecran);
        SetControlText('ANNEE',AnneeSaisie);
        SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));
        GAug := THGrid(GeTControl('GAUGMENTATION'));
        GAug.DBIndicator := True;
        //GAug.ColCount := 28;
        GAug.ColCount := 33 + GAug.FixedCols;
        Titres := HTStringList.Create;
        Titres.Insert(0, '');
        Titres.Insert(ColSal, 'Salarié');
        Titres.Insert(ColNom, 'Nom du salarié');
        Titres.Insert(ColEmploi, 'Emploi');
        Titres.Insert(ColTpsPart, 'Temps');
        Titres.Insert(ColDateE, 'Date d''entrée');
        Titres.Insert(ColEtat, 'Etat');
        Titres.Insert(ColF, 'Fixe');
        Titres.Insert(ColPctF, '%');
        Titres.Insert(ColFAp, 'Fixe Ap');
        Titres.Insert(ColV, 'Variable');
        Titres.Insert(ColPctV, '%');
        Titres.Insert(ColVap, 'Varia. Ap');
        Titres.Insert(ColCommentaire, 'Commentaire');
        Titres.Insert(ColBrut, 'Brut');
        Titres.Insert(ColPctB, '%');
        Titres.Insert(ColBrutAp, 'Brut Ap');
        Titres.Insert(ColMotif, 'Motif');
        Titres.Insert(ColEtatBis, 'Etat');
        Titres.Insert(ColMotifBis, 'MOTIF');
        Titres.Insert(ColAge, 'AGE');
        Titres.Insert(ColAnc, 'ANC');
        Titres.Insert(ColEtab, 'ETABLISSEMENT');
        Titres.Insert(ColEff, 'EFFECTIF');
        Titres.Insert(ColHor, 'HORAIRE');
        Titres.Insert(ColTN1, 'TRAVAILN1');
        Titres.Insert(ColTN2, 'TRAVAILN2');
        Titres.Insert(ColTN3, 'TRAVAILN3');
        Titres.Insert(ColTN4, 'TRAVAILN4');
        Titres.Insert(ColStat, 'CODESTAT');
        Titres.Insert(ColTpsPlein, 'Tps plein');
        Titres.Insert(ColSalAnnAv, 'Annuel av');
        Titres.Insert(ColSalAnnAp, 'Annuel ap');
        Titres.Insert(ColSalAnnTpsP, 'Annuel tps');
        GAug.Titres := Titres;
        Titres.free;
        BDefaire := TToolBarButton97(GetControl('BDEFAIRE'));
        If BDefaire <> Nil then BDefaire.OnClick := AnnulerModifs;
//        SetFocusControl('GAUGMENTATION');
         if GAug <> nil then
        begin
                GAug.PostDrawCell := GrillePostDrawCell;
                GAug.OnCellExit := GrilleCellExit;
                GAug.OnRowEnter := GrilleRowEnter;
                GAug.OnRowExit := grilleRowExit;
                GAug.GetCellCanvas := GrilleGetCellCanvas;
                GAug.OnDblClick := GrilleDblClick;
                GAug.OnCellEnter := GrilleCellEnter;
                GAug.OnEnter := EntrerGrille;
        end;
        BHisto := TToolBarButton97(GetControl('BHISTOSAL'));
        If BHisto <> Nil then BHisto.OnClick := VoirHistoAug;
{       MenuPop := TPopUpMenu(GetControl('MENU'));
        MenuPop.Items[0].OnClick := MasquerLigne;
        MenuPop.Items[1].OnClick := VoirHistoSal;
        MenuPop.Items[2].OnClick := VoirEmploi;
        MenuMasquer := TPopUpMenu(GetControl('MENUMASQUER'));}
        BSimu := TToolBarButton97(GetControl('BSIMU'));
        If BSimu <> Nil then BSimu.OnClick := BSimuClick;
        BEmploi := TToolBarButton97(GetControl('BEMPLOi'));
        If BEmploi <> Nil then BEmploi.OnClick := VoirEmploi;
        BtnSal := TToolbarButton97(GetControl('BTNSAL'));
        if BtnSal <> nil then Btnsal.OnClick := AccesSal;
        MiseEnFormeGrille;
        ChargerDonnees;
        CalculerResume(-1,False);
//        ComboEtat := THValComboBox(GetControl('PBG_ETATINTAUGM'));
//        ComboEtat.OnEnter := AccesCombo;
//        ComboEtat.OnExit := SortieCombo;
        BSelect := TToolBarButton97(GetControl('BSelectAll'));
        If BSelect <> Nil then BSelect.OnClick := BSelectClick;
        If TypeSaisie = 'VALIDRESP' then
        begin
             BValid := TToolBarButton97(GetControl('BVALIDRESP'));
             If BValid <> Nil then BValid.OnClick := ValiderEtat;
             BRefus  := TToolBarButton97(GetControl('BREFUSRESP'));
             If BRefus <> Nil then BRefus.OnClick := RefuserEtat;
             SetControlVisible('BVALIDRESP',True);
             SetControlVisible('BREFUSRESP',True);
        end;
        If TypeSaisie = 'VALIDDRH' then
        begin
             BValid := TToolBarButton97(GetControl('BVALIDDRH'));
             If BValid <> Nil then BValid.OnClick := ValiderEtat;
             BRefus  := TToolBarButton97(GetControl('BREFUSDRH'));
             If BRefus <> Nil then BRefus.OnClick := RefuserEtat;
             SetControlVisible('BVALIDDRH',True);
             SetControlVisible('BREFUSDRH',True);
        end;
{     SetControlCaption('ANCIENNETE','');
     SetControlCaption('NAISSANCE','');
     SetControlCaption('ETABLISS','');
     SetControlCaption('CODESTAT','');}
     Bimp := TToolBarButton97(GetControl('BImprimer'));
     If BImp <> Nil then BImp.OnClick := ImprimerGrille;
     BExport := TToolBarButton97(GetControl('BEXPORT'));
     If BExport <> Nil then BExport.OnClick := ExporterGrille;
     SetControlVisible('BEXPORT',False);
     BLegende := TToolBarButton97(GetControl('BLEGENDE'));
     If BLegende <> Nil then BLegende.OnClick := AfficheLegende;
     {$IFDEF EMANAGER}
//     If Consultation then SetControlVisible('VALIDSAISIE',False);
     If ((TypeSaisie = 'SAISIE') Or (TypeSaisie = 'PROPOSITION')) and Not Consultation then //PT1
     begin
          BMail := TToolBarButton97(GetControl('BMAILRESP'));
          If BMail <> Nil then
          begin
               BMail.OnClick := EnvoyerMailresp;
               BMail.Visible := True;
          end;
     end;
     {$ENDIF}
     //HMTrad.ResizeGridColumns(GAug) ;

end ;

procedure TOF_PGAUGMENTATIONSAL.GrilleCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var Test,Etat : String;
    Pct,salAp,SalAv,New,AutreSal : Double;
    NePasChanger : Boolean;
    AncBrut,NewBrut,PctBrut,Coeff : Double;
begin
        If (ACol = ColFap) or (ACol = ColPctF) then
        begin
                Pct := FormatageDouble(GAug.Cellvalues[ColPctF,ARow]);
                Pct := Arrondi(Pct,PctAugmDec);
                salAp := FormatageDouble(GAug.Cellvalues[ColFap,ARow]);
                salAv := FormatageDouble(GAug.Cellvalues[ColF,ARow]);
                If SalAv = 0 then
                begin
                     GAug.Cellvalues[ColPctF,ARow] := FormatFloat(FormatPct,0);
                     GAug.Cellvalues[ColFap,ARow] := FormatFloat('# ##0.00',0);
//                     CalculerResume;
                     AutreSal := FormatageDouble(GAug.Cellvalues[ColVap,ARow]);
                     NewBrut := AutreSal + SalAp;
                     GAug.CellValues[ColBrutap,ARow] := FormatFloat('# ##0.00',NewBrut);
                     GAug.CellValues[ColSalAnnAp,ARow] := FormatFloat('# ##0',AutreSal*12 + SalAp*NbMEquivalenceAnn);
                     Coeff := 1 / FormatageDouble(GAug.CellValues[ColEff,ARow]);
                     GAug.CellValues[ColTpsPlein,ARow] := FormatFloat('# ##0.00',NewBrut*Coeff);
                     GAug.CellValues[ColSalAnnTpsP,ARow] := FormatFloat('# ##0',(AutreSal*12 + SalAp*NbMEquivalenceAnn)*Coeff);
                     AncBrut := FormatageDouble(GAug.Cellvalues[ColBrut,ARow]);
                     If AncBrut <> 0 then PctBrut := Arrondi(((NewBrut - AncBrut)/AncBrut)*100,PctAugmDec)
                     else PctBrut := 0;
                     GAug.CellValues[ColPctB,ARow] := FormatFloat(FormatPct,Pctbrut);
                     exit;
                end;
                If ACol = ColFap then
                begin
                     New := Arrondi(((SalAp - SalAv)/SalAv)*100,PctAugmDec);
                     If (ArrondiAugm((SalAv * (Pct/100))+ SalAv)) = SalAp then NePasChanger := True
                     else NePasChanger := False;
                     If  (New <> Pct) and (NePasChanger = False) then
                     begin
                        VerifierModifs('PBG_FIXEAP',SalAp,ARow);
                        GAug.Cellvalues[ColPctF,ARow] := FormatFloat(FormatPct,New);
                        CalculerResume(ARow,False);
                        AutreSal := FormatageDouble(GAug.Cellvalues[ColVap,ARow]);
                        NewBrut := AutreSal + SalAp;
                        GAug.CellValues[ColBrutap,ARow] := FormatFloat('# ##0.00',NewBrut);
                        GAug.CellValues[ColSalAnnAp,ARow] := FormatFloat('# ##0',AutreSal*12 + SalAp*NbMEquivalenceAnn);
                        Coeff := 1 / FormatageDouble(GAug.CellValues[ColEff,ARow]);
                        GAug.CellValues[ColTpsPlein,ARow] := FormatFloat('# ##0.00',NewBrut*Coeff);
                        GAug.CellValues[ColSalAnnTpsP,ARow] := FormatFloat('# ##0',(AutreSal*12 + SalAp*NbMEquivalenceAnn)*Coeff);
                        AncBrut := FormatageDouble(GAug.Cellvalues[ColBrut,ARow]);
                        If AncBrut <> 0 then PctBrut := Arrondi(((NewBrut - AncBrut)/AncBrut)*100,PctAugmDec)
                        else PctBrut := 0;
                        GAug.CellValues[ColPctB,ARow] := FormatFloat(FormatPct,Pctbrut);
                     end;
//                     If GAug.CellValues[7,ARow] = '' then GAug.CellValues[7,ARow] := FormatFloat('# ##0.00',0);
                     GAug.Cellvalues[ColFap,ARow] := FormatFloat('# ##0.00',SalAp);
                end;
                If ACol = ColPctF then
                begin
                     New := ArrondiAugm((SalAv * (Pct/100))+ SalAv);
                     If  New <> SalAp then
                     begin
                        VerifierModifs('PBG_PCTFIXE',Pct,ARow);
                        GAug.Cellvalues[ColFap,ARow] := FormatFloat('# ##0.00',New);
                        CalculerResume(ARow,False);
                        AutreSal := FormatageDouble(GAug.Cellvalues[ColVap,ARow]);
                        Newbrut := AutreSal + New;
                        GAug.CellValues[ColBrutap,ARow] := FormatFloat('# ##0.00',NewBrut);
                        GAug.CellValues[ColSalAnnAp,ARow] := FormatFloat('# ##0',AutreSal*12 + New*NbMEquivalenceAnn);
                        Coeff := 1 / FormatageDouble(GAug.CellValues[ColEff,ARow]);
                        GAug.CellValues[ColTpsPlein,ARow] := FormatFloat('# ##0.00',NewBrut* Coeff);
                        GAug.CellValues[ColSalAnnTpsP,ARow] := FormatFloat('# ##0',(AutreSal*12 + New*NbMEquivalenceAnn)*Coeff);
                        AncBrut := FormatageDouble(GAug.Cellvalues[ColBrut,ARow]);
                        If AncBrut <> 0 then PctBrut := Arrondi(((NewBrut - AncBrut)/AncBrut)*100,PctAugmDec)
                        else PctBrut := 0;
                        GAug.CellValues[ColPctB,ARow] := FormatFloat(FormatPct,Pctbrut);
                     end;
    //                 If GAug.CellValues[ColPctF,ARow] = '' then GAug.CellValues[6,ARow] := FormatFloat('# ##0.00',0);
                     GAug.Cellvalues[ColPctF,ARow] := FormatFloat(FormatPct,Pct);
                end;
        end;
        If (ACol = ColVap) or (ACol = ColPctV) then
        begin
                Pct := FormatageDouble(GAug.Cellvalues[ColPctV,ARow]);
                Pct := Arrondi(Pct,PctAugmDec);
                salAp := FormatageDouble(GAug.Cellvalues[ColVap,ARow]);
                salAv := FormatageDouble(GAug.Cellvalues[ColV,ARow]);
                If SalAv = 0 then
                begin
                     GAug.Cellvalues[ColPctV,ARow] := FormatFloat(FormatPct,0);
                     GAug.Cellvalues[ColVap,ARow] := FormatFloat('# ##0.00',SalAp);
//                     CalculerResume;
                     AutreSal := FormatageDouble(GAug.Cellvalues[ColFap,ARow]);
                     NewBrut := AutreSal + SalAp;
                     GAug.CellValues[ColBrutap,ARow] := FormatFloat('# ##0.00',NewBrut);
                     GAug.CellValues[ColSalAnnAp,ARow] := FormatFloat('# ##0',SalAp*12 + AutreSal*NbMEquivalenceAnn);
                     Coeff := 1 / FormatageDouble(GAug.CellValues[ColEff,ARow]);
                     GAug.CellValues[ColTpsPlein,ARow] := FormatFloat('# ##0.00',NewBrut*Coeff);
                     GAug.CellValues[ColSalAnnTpsP,ARow] := FormatFloat('# ##0',(SalAp*12 + AutreSal*NbMEquivalenceAnn)*Coeff);
                     AncBrut := FormatageDouble(GAug.Cellvalues[ColBrut,ARow]);
                     If AncBrut <> 0 then PctBrut := Arrondi(((NewBrut - AncBrut)/AncBrut)*100,PctAugmDec)
                     else PctBrut := 0;
                     GAug.CellValues[ColPctB,ARow] := FormatFloat(FormatPct,Pctbrut);
                     exit;
                end;
                If ACol = ColVap then
                begin
                     New := Arrondi(((SalAp - SalAv)/SalAv)*100,PctAugmDec);
                     If (ArrondiAugm((SalAv * (Pct/100))+ SalAv)) = SalAp then NePasChanger := True
                     else NePasChanger := False;
                     If  (New <> Pct) and (NePasChanger = False) then
                     begin
                        VerifierModifs('PBG_VARIABLEAP',SalAp,ARow);
                        GAug.Cellvalues[ColPctV,ARow] := FormatFloat(FormatPct,New);
                        CalculerResume(ARow,False);
                        AutreSal := FormatageDouble(GAug.Cellvalues[ColFap,ARow]);
                        NewBrut := AutreSal + SalAp;
                        GAug.CellValues[ColBrutap,ARow] := FormatFloat('# ##0.00',NewBrut);
                        GAug.CellValues[ColSalAnnAp,ARow] := FormatFloat('# ##0',SalAp*12 + AutreSal*NbMEquivalenceAnn);
                        Coeff := 1 / FormatageDouble(GAug.CellValues[ColEff,ARow]);
                        GAug.CellValues[ColTpsplein,ARow] := FormatFloat('# ##0.00',NewBrut * Coeff);
                        GAug.CellValues[ColSalAnnTpsP,ARow] := FormatFloat('# ##0',(SalAp*12 + AutreSal*NbMEquivalenceAnn)*Coeff);
                        AncBrut := FormatageDouble(GAug.Cellvalues[ColBrut,ARow]);
                        If AncBrut <> 0 then PctBrut := Arrondi(((NewBrut - AncBrut)/AncBrut)*100,PctAugmDec)
                        else PctBrut := 0;
                        GAug.CellValues[ColPctB,ARow] := FormatFloat(FormatPct,Pctbrut);
                     end;
//                     If GAug.CellValues[10,ARow] = '' then GAug.CellValues[10,ARow] := FormatFloat('# ##0.00',0);
                     GAug.Cellvalues[ColVap,ARow] := FormatFloat('# ##0.00',SalAp);
                end;
                If ACol = ColPctV then
                begin
                     New := ArrondiAugm((SalAv * (Pct/100))+ SalAv);
                     If  New <> SalAp then
                     begin
                        VerifierModifs('PBG_PCTVARIABLE',Pct,ARow);
                        GAug.Cellvalues[ColVap,ARow] := FormatFloat('# ##0.00',New);
                        CalculerResume(ARow,False);
                        AutreSal := FormatageDouble(GAug.Cellvalues[ColFap,ARow]);
                        NewBrut := AutreSal + New;
                        GAug.CellValues[ColBrutap,ARow] := FormatFloat('# ##0.00',NewBrut);
                        GAug.CellValues[ColSalAnnAp,ARow] := FormatFloat('# ##0',New*12 + AutreSal*NbMEquivalenceAnn);
                        Coeff := 1 / FormatageDouble(GAug.CellValues[ColEff,ARow]);
                        GAug.CellValues[ColTpsPlein,ARow] := FormatFloat('# ##0.00',NewBrut * Coeff);
                        GAug.CellValues[ColSalAnnTpsP,ARow] := FormatFloat('# ##0',(New*12 + AutreSal*NbMEquivalenceAnn)*Coeff);
                        AncBrut := FormatageDouble(GAug.Cellvalues[ColBrut,ARow]);
                        If AncBrut <> 0 then PctBrut := Arrondi(((NewBrut - AncBrut)/AncBrut)*100,PctAugmDec)
                        else PctBrut := 0;
                        GAug.CellValues[ColPctB,ARow] := FormatFloat(FormatPct,Pctbrut);
                     end;
//                     If GAug.CellValues[9,ARow] = '' then GAug.CellValues[9,ARow] := FormatFloat('# ##0.00',0);
                     GAug.Cellvalues[ColPctV,ARow] := FormatFloat(FormatPct,Pct);
                end;
        end;
        If ACol = ColEtat then
        begin
             If (GAug.CellValues[ColEtatBis,ARow] <> GAug.CellValues[ColEtat,ARow]) then
             begin
                If (TypeSaisie = 'SAISIE') Then
                    GAug.CellValues[ColEtatBis,ARow] := 'F' //PT1
                Else
                    GAug.CellValues[ColEtatBis,ARow] := GAug.CellValues[ColEtat,ARow];
                  CalculerResume(ARow,False);
             end;
        end;
        If ACol = ColMotif then
        begin
             Etat := GAug.CellValues[ColEtat,ARow];
             If ((TypeSaisie = 'SAISIE') Or (TypeSaisie = 'PROPOSITION')) then //PT1
             begin
                  If (Etat <> '001') and (Etat <>'002') and (Etat <>'000') then GAug.CellValues[ColMotif,ARow] := GAug.CellValues[ColMotifBis,ARow];
                  if (COnsultation) and (Etat = '002') then GAug.CellValues[ColMotif,ARow] := GAug.CellValues[ColMotifBis,ARow];
             end;
             If TypeSaisie= 'VALIDRESP' then
             begin
                  If Etat >= '006' then GAug.CellValues[ColMotif,ARow] := GAug.CellValues[ColMotifBis,ARow]
             end;
             Test := GAug.CellValues[ColMotif,ARow];
             GAug.CellValues[ColMotifBis,ARow] := Test;
        end;
end;

procedure TOF_PGAUGMENTATIONSAL.GrillePostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
  If (ACol = ColF) or (ACol = ColV) or (ACol = ColBrut) or (ACol = ColBrutap) or (ACol = ColPctB) or (ACol = ColTpsPlein)
     or (ACol = ColSalAnnAv) or (ACol = ColSalAnnAp) or (ACol = ColSalAnnTpsP) then GridGriseCell(GAug, Acol, Arow, Canvas);
  If (ACol > ColEmploi) and (ACol < ColEtat) then GridGriseCell(GAug, Acol, Arow, Canvas);

end;

Function TOF_PGAUGMENTATIONSAL.FormatageDouble(Chaine : String) : Double;
var Longueur,Indice : Integer;
    St : String;
begin
        longueur:=Length (Chaine);
        If Longueur < 1 then
        begin
         Result := 0;
         exit;
        end;
        indice:=1;
        repeat
        if (Chaine [Indice]<>' ') then
        St := St + Chaine [Indice];
        inc (Indice);
        until (Indice=Longueur+1);
        If Not IsNumeric (St) then
        begin
             Result := 0;
             PGIBox(St + ' n''est pas une valeur numérique',Ecran.Caption);
             Exit;
        end;
        result := StrToFloat(St);
end;

procedure TOF_PGAUGMENTATIONSAL.VoirEmploi(Sender : TObject);
var i : Integer;
    LibEmploi,Salarie : String;
    Q : TQuery;
    T : Tob;
    Total : Double;
    NbSal : Integer;
    ChampSalaire : String;
begin
        TobAug.GetGridDetail (GAug,GAug.RowCount-1,'','INDICATEUR;'+ StChampGrid);
        TobEmploiAug := Tob.Create('LesEmplois',Nil,-1);
        ChampSalaire := ChampSalaires('TOUS');
        For i := 0 to TobAug.Detail.Count - 1 do
        begin
                Salarie := TobAug.Detail[i].GetValue('PSA_SALARIE');
                Q := OpenSQL('SELECT PSA_LIBELLEEMPLOI FROM SALARIES WHERE PSA_SALARIE="'+Salarie+'"',True);
                If Not Q.Eof then LibEmploi := Q.FindField('PSA_LIBELLEEMPLOI').AsString;
                ferme(Q);
                T := tobEmploiAug.FindFirst(['LIBEMPLOI'],[LibEmploi],False);
                If T = Nil then
                begin
                        T := tob.Create('Lesemploisfille',TobEmploiAug,-1);
                        T.AddChampSupValeur('LIBEMPLOI',LibEmploi,False);
                        If ChampSalaire <> '' then
                        begin
                             Q := OpenSQL('SELECT COUNT(PSA_SALARIE) NBSAL,MIN(1/PSA_UNITEPRISEFF*('+ChampSalaire+')) SAMIN,MAX(1/PSA_UNITEPRISEFF*('+ChampSalaire+')) SAMAX,'+
                             'SUM(1/PSA_UNITEPRISEFF*('+ChampSalaire+')) TOTAL '+
                              'FROM SALARIES WHERE PSA_LIBELLEEMPLOI="'+LibEmploi+'" AND (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'") AND PSA_UNITEPRISEFF>0',True);
                              If Not Q.Eof then
                              begin
                                      T.AddChampSupValeur('NBSAL',Q.FindField('NBSAL').Asfloat,False);
                                      T.AddChampSupValeur('MIN',Q.FindField('SAMIN').Asfloat,False);
                                      T.AddChampSupValeur('MAX',Q.FindField('SAMAX').Asfloat,False);
                                      Total := Q.FindField('TOTAL').Asfloat;
                                      NbSal := Q.FindField('NBSAL').AsInteger;
                              end
                              else
                              begin
                                      T.AddChampSupValeur('NBSAL',0,False);
                                      T.AddChampSupValeur('MIN',0,False);
                                      T.AddChampSupValeur('MAX',0,False);
                                      Total := 0;
                                      NbSal := 0;
                              end;
                              Ferme(Q);
                        end
                        else
                        begin
                             T.AddChampSupValeur('MIN',0,False);
                             T.AddChampSupValeur('MAX',0,False);
                             Total := 0;
                             NbSal := 0;
                        end;
                        If NbSal  <> 0 then T.AddChampSupValeur('MOY',ArrondiAugm((Total) / NbSal),False)
                        else T.AddChampSupValeur('MOY',0,False);
                end;
        end;
        AGLLanceFiche('PAY','AUGMENTATIONCAT','','','');
        TobEmploiAUg.Free;
end;

procedure TOF_PGAUGMENTATIONSAL.CalculerResume(Ligne : Integer;Init : Boolean = False);
Var i,NbSal : Integer;
    Fixe,NewFixe,Variable,NewVariable,Tot,TotNew : Double;
    FixeAnn,NewFixeAnn,VariableAnn,NewVariableAnn,TotAnn,TotNewAnn : Double;
    Pct,PctAnn : Double;
    Etat : String;
begin
        If Init then TobAug.GetGridDetail (GAug,GAug.RowCount-1,'','INDICATEUR;'+ StChampGrid);
        Fixe := 0;
        NewFixe := 0;
        Variable := 0;
        NewVariable := 0;
        NbSal := 0;
        FixeAnn := 0;
        VariableAnn := 0;
        NewFixeAnn := 0;
        NewVariableAnn := 0;
        For i := 0 to TobAug.Detail.Count - 1 do
        begin
                If Ligne <> -1 then
                begin
                     If GAug.CellValues[ColSal,Ligne] = TobAug.Detail[i].GetValue('PSA_SALARIE') then
                     begin
                          TobAug.Detail[i].PutValue('PBG_ETATINTAUGM',GAug.Cellvalues[ColEtat,Ligne]);
                          TobAug.Detail[i].PutValue('PBG_FIXEAV',FormatageDouble(GAug.Cellvalues[ColF,Ligne]));
                          TobAug.Detail[i].PutValue('PBG_VARIABLEAV',FormatageDouble(GAug.Cellvalues[ColV,Ligne]));
                          TobAug.Detail[i].PutValue('PBG_VARIABLEAP',FormatageDouble(GAug.Cellvalues[ColVAp,Ligne]));
                          TobAug.Detail[i].PutValue('PBG_FIXEAP',FormatageDouble(GAug.Cellvalues[ColFAp,Ligne]));
                     end;
                end;
                Etat := TobAug.Detail[i].GetValue('PBG_ETATINTAUGM');
                Fixe := Fixe + TobAug.Detail[i].GetValue('PBG_FIXEAV');
                FixeAnn := Fixe * NbMEquivalenceAnn;
                Variable := Variable + TobAug.Detail[i].GetValue('PBG_VARIABLEAV');
                VariableAnn := Variable * 12;
                If (Etat <> '004') and (Etat <> '007') then
                begin
                     NewVariable := NewVariable + TobAug.Detail[i].GetValue('PBG_VARIABLEAP');
                     NewVariableAnn := NewVariable * 12;
                     NewFixe := NewFixe + TobAug.Detail[i].GetValue('PBG_FIXEAP');
                     NewFixeAnn := NewFixe * NbMEquivalenceAnn;
                end
                else
                begin
                     NewVariable := NewVariable + TobAug.Detail[i].GetValue('PBG_VARIABLEAV');
                     NewFixe := NewFixe + TobAug.Detail[i].GetValue('PBG_FIXEAV');
                     NewVariableAnn := NewVariable * 12;
                     NewFixeAnn := NewFixe * NbMEquivalenceAnn;
                end;
                NbSal := NbSal + 1;
        end;
    {   For i := 1 to GAug.RowCount - 1 do
        begin
                Etat := GAug.CellValues[ColEtat,i];
                Fixe := Fixe + FormatageDouble(GAug.Cellvalues[ColF,i]);;
                FixeAnn := Fixe * NbMEquivalenceAnn;
                Variable := Variable + FormatageDouble(GAug.Cellvalues[ColV,i]);;
                VariableAnn := Variable * 12;
                If (Etat <> '004') and (Etat <> '007') then
                begin
                     NewVariable := NewVariable + FormatageDouble(GAug.Cellvalues[ColVap,i]);;
                     NewVariableAnn := NewVariable * 12;
                     NewFixe := NewFixe + FormatageDouble(GAug.Cellvalues[ColFap,i]);;
                     NewFixeAnn := NewFixe * NbMEquivalenceAnn;
                end
                else
                begin
                     NewVariable := NewVariable + FormatageDouble(GAug.Cellvalues[ColV,i]);;
                     NewFixe := NewFixe + FormatageDouble(GAug.Cellvalues[ColF,i]);
                     NewVariableAnn := NewVariable * 12;
                     NewFixeAnn := NewFixe * NbMEquivalenceAnn;
                end;
                NbSal := NbSal + 1;
        end;   }
        Tot := Fixe + Variable;
        TotAnn := FixeAnn + VariableAnn;
        TotNew := NewFixe + NewVariable;
        TotNewAnn := NewFixeAnn + NewVariableAnn;
        SetControlCaption('LIBSAL',IntToStr(NbSal) + ' salarié(s)');
        SetControlCaption('MTFXAV',FormatFloat('#,##0.00',Fixe));
        SetControlCaption('MTFXAP',FormatFloat('#,##0.00',NewFixe));
        SetControlCaption('ANMTFXAV',FormatFloat('#,##0',FixeAnn));
        SetControlCaption('ANMTFXAP',FormatFloat('#,##0',NewFixeAnn));
        If fixe <> 0 then Pct := ((NewFixe - Fixe) / Fixe) * 100
        else Pct := 0;
        Pct := Arrondi(Pct,PctAugmDec);
        SetControlCaption('MTFXPCT',FormatFloat('#,##0.00',Arrondi(Pct,PctAugmDec)));
        If fixeAnn <> 0 then PctAnn := ((NewFixeAnn - FixeAnn) / FixeAnn) * 100
        else PctAnn := 0;
        SetControlCaption('ANNMTFXPCT',FormatFloat('#,##0.00',Arrondi(PctAnn,PctAugmDec)));
        SetControlText('EDITPCTF',FormatFloat('#,##0.00',Arrondi(Pct,PctAugmDec)));
        SetControlCaption('MTVARAV',FormatFloat('#,##0.00',Variable));
        SetControlCaption('MTVARAP',FormatFloat('#,##0.00',NewVariable));
        SetControlCaption('ANMTVARAV',FormatFloat('#,##0',VariableAnn));
        SetControlCaption('ANMTVARAP',FormatFloat('#,##0',NewVariableAnn));
        If Variable <> 0 then Pct := ((NewVariable - Variable) / Variable) * 100
        else Pct := 0;
        Pct := Arrondi(Pct,PctAugmDec);
        SetControlCaption('MTVARPCT',FormatFloat('#,##0.00',Arrondi(Pct,PctAugmDec)));
        If VariableAnn <> 0 then PctAnn := ((NewVariableAnn - VariableAnn) / VariableAnn) * 100
        else PctAnn := 0;
        PctAnn := Arrondi(PctAnn,PctAugmDec);
        SetControlCaption('ANNMTVARPCT',FormatFloat('#,##0.00',Arrondi(PctAnn,PctAugmDec)));
        SetControlText('EDITPCTV',FormatFloat('#,##0.00',Arrondi(Pct,PctAugmDec)));
        SetControlCaption('MTTOTAV',FormatFloat('#,##0.00',Tot));
        SetControlCaption('MTTOTAP',FormatFloat('#,##0.00',TotNew));
        SetControlCaption('ANMTTOTAV',FormatFloat('#,##0',TotAnn));
        SetControlCaption('ANMTTOTAP',FormatFloat('#,##0',TotNewAnn));
        If Tot <> 0 then Pct := ((totNew - Tot) / Tot) * 100
        else Pct := 0;
        Pct := Arrondi(Pct,PctAugmDec);
        If TotAnn <> 0 then PctAnn := ((totNewAnn - TotAnn) / TotAnn) * 100
        else PctAnn := 0;
        PctAnn := Arrondi(PctAnn,PctAugmDec);
        SetControlCaption('ANNMTTOTPCT',FormatFloat('#,##0.00',Arrondi(pctAnn,PctAugmDec)));
        SetControlCaption('MTTOTPCT',FormatFloat('#,##0.00',Arrondi(pct,PctAugmDec)));
        SetControlText('EDITPCTT',FormatFloat('#,##0.00',Arrondi(Pct,PctAugmDec)));
end;


procedure TOF_PGAUGMENTATIONSAL.BSimuClick(Sender : Tobject);
var Fixe,Variable,PctFixe,PctVar : Double;
    i : Integer;
    StPctF,StPctVar,LePct,MotifAug : String;
    BTemps : TToolBarButton97;
begin
         BTemps := TToolBarButton97(GetControl('BTEMPS'));
         If BTemps <> Nil then
         begin
              If BTemps.Down = True then
              begin
                   BTemps.Down := False;
                   VoirTempsPlein(BTemps);
              end;
        end;
        If (GAug.nbSelected < 1) and (GAug.AllSelected = False) then
        begin
             PGIBox('Vous devez sélectionner 1 ou plusieurs salariés',Ecran.Caption);
             Exit;
        end;
        TobAug.GetGridDetail (GAug,GAug.RowCount-1,'','INDICATEUR;'+ StChampGrid);
        Fixe := 0;
        Variable := 0;
        For i := 0 to TobAug.Detail.Count - 1 do
        begin
                If Gaug.IsSelected(i+1) then
                begin
                     Fixe := Fixe + TobAug.Detail[i].GetValue('PBG_FIXEAV');
                     Variable := Variable + TobAug.Detail[i].GetValue('PBG_VARIABLEAV');
                end;
        end;
        LePct := AGLLanceFiche('PAY','AUGMGLOB','','',FloatToStr(Fixe)+';'+FloatToStr(Variable));
        StPctF := ReadTokenPipe(LePct,';');
        StPctVar := ReadTokenPipe(LePct,';');
        MotifAug := ReadTokenPipe(LePct,';');
        If IsNumeric(StPctF) then PctFixe := StrToFloat(StPctF)
        else PctFixe := 0;
        If IsNumeric(StPctVar) then PctVar := StrToFloat(StPctVar)
        else PctVar := 0;
        If (PctFixe <> 0) or (PctVar<>0) or (MotifAug <> '') then DupliquerSalaire(PctFixe,PctVar,MotifAug);
        SetControlProperty('BSelectAll','Down',False);
        GAug.ClearSelected;
        GAug.AllSelected := False;
end;

procedure TOF_PGAUGMENTATIONSAL.MiseEnFormeGrille;
var i : Integer;
begin
//        Gaug.ColWidths[0] := -1;
        Gaug.ColDrawingModes[ColTpsPart]:= 'IMAGE';
        GAug.ColEditables[ColTpsPart] := False;
        GAug.ColWidths[ColTpsPart] := -1;
        GAug.ColEditables[ColNom] := False;
        GAug.ColEditables[ColEmploi] := False;
//        Gaug.ColFormats[2] := 'CB=PGLIBEMPLOI';
        GAug.ColEditables[ColDateE] := False;
        GAug.ColFormats[ColDateE] := ShortDateFormat;
        If TypeSaisie = 'PROPOSITION' then GAug.ColEditables[ColEtat] := False; //PT1
        GAug.ColEditables[ColF] := False;
        GAug.ColEditables[ColV] := False;
        For i := ColF to ColVap do
        begin
                GAug.ColFormats[i] := '# ##0.00';
                GAug.ColAligns[i] := taRightJustify;
        end;
        Gaug.ColDrawingModes[ColEtat]:= 'IMAGE';
        Gaug.ColFormats[ColEtat] := 'CB=PGAUGMETATVALID';
        //Gaug.ColFormats[12] := 'CB=PGMOTIFAUGM';
        GAug.ColFormats[ColBrut] := '# ##0.00';
        GAug.ColAligns[ColBrut] := taRightJustify;
        GAug.ColFormats[ColPctB] := FormatPct;
        GAug.ColFormats[ColPctF] := FormatPct;
        GAug.ColFormats[ColPctV] := FormatPct;
        GAug.ColAligns[ColPctB] := taRightJustify;
        GAug.ColFormats[ColBrutap] := '# ##0.00';
        GAug.ColAligns[ColBrutap] := taRightJustify;
        GAug.ColFormats[ColTpsPlein] := '# ##0.00';
        GAug.ColAligns[ColTpsPlein] := taRightJustify;
        GAug.ColFormats[ColSalAnnAv] := '# ##0';
        GAug.ColAligns[ColSalAnnAv] := taRightJustify;
        GAug.ColFormats[ColSalAnnAp] := '# ##0';
        GAug.ColAligns[ColSalAnnAp] := taRightJustify;
        GAug.ColFormats[ColSalAnnTpsP] := '# ##0';
        GAug.ColAligns[ColSalAnnTpsP] := taRightJustify;
        GAug.ColWidths[ColSal] := -1;
        GAug.ColWidths[ColNom] := 100;
        GAug.ColWidths[ColEmploi] := 100;
        GAug.ColWidths[ColDateE] := -1;
        GAug.ColWidths[ColF] := 50;
        GAug.ColWidths[ColPctF] := 40;//25 ggg
        GAug.ColWidths[ColFap] := 50;
        GAug.ColWidths[ColV] := 50;
        GAug.ColWidths[ColPctV] := 40;//25 ggg
        GAug.ColWidths[ColVap] := 50;
        GAug.ColWidths[ColCommentaire] := -1;
        GAug.ColWidths[ColEtat] := 40;
        GAug.ColWidths[ColBrut] := 50;
        GAug.ColWidths[ColPctB] := 40;//25 ggg
        GAug.ColWidths[ColBrutap] := 50;
        GAug.ColWidths[ColTpsPlein] := 50;
        GAug.ColWidths[ColSalAnnAv] := 60;
        GAug.ColWidths[ColSalAnnAp] := 60;
        GAug.ColWidths[ColSalAnnTpsP] := 60;
        GAug.ColWidths[ColMotif] := 60;
        If Not GestionSalFixe then
        begin
             GAug.ColWidths[ColF] := -1;
             GAug.ColWidths[ColPctF] := -1;
             GAug.ColWidths[ColFap] := -1;
        end;
        If not GestionSalVariable then
        begin
             GAug.ColWidths[ColV] := -1;
             GAug.ColWidths[ColPctV] := -1;
             GAug.ColWidths[ColVap] := -1;
        end;
//        GAug.ColWidths[15] := -1;
        Gaug.ColFormats[ColMotif] := 'CB=PGMOTIFAUGM';
        GAug.ColWidths[ColEtatBis] := -1;
        GAug.ColWidths[ColMotifBis] := -1;
        GAug.ColWidths[ColAge] := -1;
        GAug.ColWidths[ColAnc] := -1;
        GAug.ColWidths[ColEtab] := -1;
        GAug.ColWidths[ColEff] := -1;
        GAug.ColWidths[ColHor] := -1;
        GAug.ColWidths[ColTN1] := -1;
        GAug.ColWidths[ColTN2] := -1;
        GAug.ColWidths[ColTN3] := -1;
        GAug.ColWidths[ColTN4] := -1;
        GAug.ColWidths[ColStat] := -1;
end;

procedure TOF_PGAUGMENTATIONSAL.ChargerDonnees;
begin
        GAug.RowCount := TobAug.Detail.Count + 1;
        TobAug.PutGridDetail(GAug,False,False,StChampGrid,False);
        AfficheInfosSal(GAug.Row);
        SetControlText('COMMENTAIRE',GAug.CellValues[ColCommentaire,GAug.Row]);
end;

procedure TOF_PGAUGMENTATIONSAL.DupliquerSalaire(PctFixe,PctVar : Double;MotifAug : String);
Var InitFixe,InitVar,SalF,SalV : Double;
    Rep,i : Integer;
    Etat,LibQuest : String;
    EtatOk : Boolean;
    NewBrut,AncBrut,PctBrut,Coeff : Double;
begin
        If GAug.nbSelected = 0 then
        begin
             PGIBox('Affectation impossible car aucun salarié n''a été sélectionné',Ecran.Caption);
             Exit;
        end;
        LibQuest := '';
        If PctFixe <> 0 then LibQuest := '#13#10 '+FloatToStr(PctFixe) +' % d''augmentation de salaire fixe';
        If PctVar <> 0 then LibQuest := LibQuest + '#13#10 '+FloatToStr(PctVar) +' % d''augmentation de salaire variable';
        Rep := PGIAsk('Voulez-vous affecter aux salariés sélectionnés : '+LibQuest,Ecran.Caption);
//        '#13#10 '+FloatToStr(PctFixe) +' % d''augmentation de salaire fixe'+
//        '#13#10 '+FloatToStr(PctVar) +' % d''augmentation de salaire variable',Ecran.Caption);
        If Rep <> MrYes then Exit;
        For i := 1 to GAug.RowCount - 1 do
        begin
                Etat := GAug.Cellvalues[ColEtatBis,i];
                If ((TypeSaisie = 'SAISIE') Or (TypeSaisie = 'PROPOSITION')) then //PT1
                begin
                     If (Etat = '001') or ((Etat = '002') and (Consultation = False)) or (Etat = '000') then EtatOk := True
                     else EtatOk := False;
                end
                else
                If TypeSaisie = 'VALIDRESP' then
                begin
                     If (Etat < '006') then EtatOk := True
                     else EtatOk := False;
                end
                else
                EtatOk := True;
                If ((Etat   <> '011') and (Etat <> '012')) and (EtatOk) and (Gaug.IsSelected(i)) then
                begin
                        If PctFixe <> 0 then
                        begin
                             GAug.CellValues[ColPctF,i] := FormatFloat(FormatPct,PctFixe);
                             InitFixe := FormatageDouble(GAug.Cellvalues[ColF,i]);
                             If InitFixe <> 0 then GAug.Cellvalues[ColFap,i] := FormatFloat('# ##0.00',ArrondiAugm((InitFixe * (PctFixe/100))+ InitFixe))
                             else GAug.Cellvalues[ColFap,i] := FormatFloat('# ##0.00',0);
                        end;
                        If PctVar <> 0 then
                        begin
                             GAug.CellValues[ColPctV,i] := FormatFloat(FormatPct,PctVar);
                             InitVar := FormatageDouble(GAug.Cellvalues[ColV,i]);
                             If InitVar <> 0 then GAug.Cellvalues[ColVap,i] := FormatFloat('# ##0.00',ArrondiAugm((InitVar * (PctVar/100))+ InitVar))
                             else GAug.Cellvalues[ColVap,i] := FormatFloat('# ##0.00',0);
                        end;
                        If MotifAug <> '' then
                        begin
                             GAug.CellValues[ColMotif,i] := MotifAug;
                        end;
                        If ((TypeSaisie = 'SAISIE') Or (TypeSaisie = 'PROPOSITION')) and (Etat = '001') and (not Consultation) then GAug.CellValues[ColEtat,i] := '002'; //PT1
                        If ((TypeSaisie = 'SAISIE') Or (TypeSaisie = 'PROPOSITION')) and (Etat = '001') and (Consultation) then GAug.CellValues[ColEtat,i] := '000'; //PT1
                        If (TypeSaisie = 'VALIDRESP') and (Etat < '005') then GAug.CellValues[ColEtat,i] := '005';
                        If (TypeSaisie = 'VALIDDRH') and (Etat < '008') then GAug.CellValues[ColEtat,i] := '008';
                        SalF := FormatageDouble(GAug.Cellvalues[ColF,i]);
                        SalV := FormatageDouble(GAug.Cellvalues[ColV,i]);
                        GAug.CellValues[ColBrut,i] := FormatFloat('# ##0.00',SalF+SalV);
                        SalF := FormatageDouble(GAug.Cellvalues[ColFap,i]);
                        SalV := FormatageDouble(GAug.Cellvalues[ColVap,i]);
                        NewBrut := SalF+SalV;
                        GAug.CellValues[ColBrutap,i] := FormatFloat('# ##0.00',NewBrut);
                        GAug.CellValues[ColSalAnnAp,i] := FormatFloat('# ##0',SalV*12 + SalF*NbMEquivalenceAnn);
                        Coeff := 1 / FormatageDouble(GAug.CellValues[ColEff,i]);
                        GAug.CellValues[ColTpsPlein,i] := FormatFloat('# ##0.00',NewBrut * Coeff);
                        GAug.CellValues[ColSalAnnTpsP,i] := FormatFloat('# ##0',(SalV*12 + SalF*NbMEquivalenceAnn) * Coeff);
                        AncBrut := FormatageDouble(GAug.Cellvalues[ColBrut,i]);
                        If AncBrut <> 0 then PctBrut := Arrondi(((NewBrut - AncBrut)/AncBrut)*100,PctAugmDec)
                        else PctBrut := 0;
                        GAug.CellValues[ColPctB,i] := FormatFloat(FormatPct,Pctbrut);
                end;
        end;
        CalculerResume(-1,True);
end;

{procedure TOF_PGAUGMENTATIONSAL.NonSaisie(const GS: THGrid; const Acol, Arow: integer; Canvas: Tcanvas);
var
  R: TRect;
begin
  Canvas.Font.Color := Clred;
  GAug.Font.Color := ClRed;
//  Canvas.Brush.color := ClRed;// $00C9C9CB; // $00E4E4E4;
  Canvas.Pen.Color := Canvas.Brush.color;
  Canvas.Pen.mode := pmMask; //pmCopy;
  Canvas.Pen.Style := psDot; //psClear;
  Canvas.Pen.Width := 0;
  R := GS.CellRect(Acol, ARow);
  Canvas.Rectangle(R.Left, R.Top, R.Right, R.Bottom);
end;      }

procedure TOF_PGAUGMENTATIONSAL.GrilleRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var Commentaire : String;
begin
        If LastColEnter <> Ou then
        begin
             Commentaire := GAug.CellValues[ColCommentaire,Ou];
             SetControltext('COMMENTAIRE',Commentaire);
             LastColEnter := Ou;
        end;
        AfficheInfosSal(Ou);
        If TypeSaisie= 'VALIDRESP' then
        begin
             If GAug.CellValues[ColEtatBis,Ou] < '006' then GAug.ColFormats[ColEtat] := 'CB=PGAUGMETATVALID2'
             else GAug.ColFormats[ColEtat] := 'CB=PGAUGMETATVALID';
        end
        Else If TypeSaisie = 'SAISIE' Then GAug.ColFormats[ColEtat] := 'CB=PGAUGMETATVALID1'; //PT1
end;

procedure TOF_PGAUGMENTATIONSAL.GrilleRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var Commentaire : String;
begin
     If LastColexit <> Ou then
     begin
          Commentaire := GetControlText('COMMENTAIRE');
          GAug.CellValues[ColCommentaire,Ou] := Commentaire;
          LastColExit := Ou;
     end;
    If TypeSaisie = 'VALIDRESP' then Gaug.ColFormats[ColEtat] := 'CB=PGAUGMETATVALID'
    Else If TypeSaisie = 'SAISIE' Then GAug.ColFormats[ColEtat] := 'CB=PGAUGMETATVALID1'; //PT1
 end;

procedure TOF_PGAUGMENTATIONSAL.GrilleGetCellCanvas( ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
Var Effectif : Double;
begin
      If ARow = 0 then Exit;
      If (GAug.CellValues[ColEtatBis,ARow] = '011') or (GAug.CellValues[ColEtatBis,ARow] = '012') then Canvas.Font.Color := Clred;
//      If (ACol = ColPctF) or  (ACol = ColFap) or (ACol = ColPctV) or (ACol = ColVap) or (ACol = ColBrutap) or (ACol = ColPctB) or (ACol=ColTpsPlein)
//      or (ACol = ColSalAnnAp) or (ACol = ColSalAnnAv)or (ACol = ColSalAnnTpsP)then Canvas.Font.Style := [fsBold];
      If (ACol = ColF) or  (ACol = ColFap) or (ACol = ColV) or (ACol = ColVap) or (ACol = ColBrutap) or (ACol = ColBrut) then
      begin
            Effectif := StrToFloat(GAug.CellValues[ColEff,ARow]);
            If (Effectif <> 1) and (effectif<> 0)then
            begin
                 If TToolBarButton97(GetControl('BTEMPS')).Down = True then Canvas.Font.Color := ClRed
                 else Canvas.Font.Color := ClBlack;
            end;
      end;
end;

procedure TOF_PGAUGMENTATIONSAL.GrilleDblClick(Sender : TObject);
Var Salarie : String;
    Ligne : Integer;
begin
        Ligne := GAug.Row;
        Salarie := GAug.CellValues[ColSal,Ligne];
        If Salarie = '' then Exit;
        AGLLanceFiche('PAY','AUGMDETAILSAL','','',Salarie+';'+AnneeSaisie);
end;

procedure TOF_PGAUGMENTATIONSAL.VoirHistoAug(Sender : TObject);
Var Salarie : String;
    Ligne : Integer;
begin
        Ligne := GAug.Row;
        Salarie := GAug.CellValues[ColSal,Ligne];
        If Salarie = '' then
        begin
             PGIBox('Vous devez vous positionner sur un salarié',ecran.Caption);
             Exit;
        end;
        AGLLanceFiche('PAY','AUGMENTASUIVI','','',Salarie);
end;

procedure TOF_PGAUGMENTATIONSAL.AccesSal(Sender: TObject);
var
  CodeSalarie: string;
begin
  CodeSalarie := GAug.Cells[ColSal, GAug.Row];
  if (CodeSalarie <> '') and (CodeSalarie[1] <> '-') then AglLanceFiche('PAY', 'SALARIE_PRIM', '', CodeSalarie, 'ACTION=CONSULTATION');
end;

Function TOF_PGAUGMENTATIONSAL.ChampSalaires(TypeSalaire : String) : String;
var LesSalaires,TempSalaire,Champ,Requete : String;
    i : Integer;
begin
     LesSalaires := '';
     If TypeSalaire = 'FIXE' then LesSalaires := GetParamSoc('SO_PGAUGFIXE');
     If TypeSalaire = 'VARIABLE' then LesSalaires := GetParamSoc('SO_PGAUGVARIABLE');
     If TypeSalaire = 'TOUS' then LesSalaires := GetParamSoc('SO_PGAUGFIXE')+';'+GetParamSoc('SO_PGAUGVARIABLE');
     If LesSalaires = '' then
     begin
          Result := '';
          Exit;
     end;
     Requete := '';
     While LesSalaires <> '' do
     begin
          Champ := '';
          TempSalaire := ReadTokenPipe(LesSalaires,';');
          For i := 1 to 5 do
          begin
               if TempSalaire = 'AN'+IntToStr(i) then Champ := 'PSA_SALAIRANN'+IntToStr(i);
          end;
          For i := 1 to 5 do
          begin
               If TempSalaire = 'MO'+IntToStr(i) then Champ := 'PSA_SALAIREMOIS'+IntToStr(i);
          end;
          If Champ <> '' then
          begin
               If Requete <> '' then Requete := Requete + ' + '+Champ
               else Requete := Champ;
          end;
     end;
     Result := Requete;
end;

Procedure TOF_PGAUGMENTATIONSAL.BSelectClick(Sender : Tobject);
var Selection : Boolean;
begin
     Selection := TToolBarButton97(Sender).Down;
     If Selection then GAug.AllSelected := True
     else GAug.AllSelected := False;
end;

Procedure TOF_PGAUGMENTATIONSAL.AfficheInfosSal(Ou : Integer);
var SalariE : String;
   LeLabel : THLabel;
begin
     Salarie := GAug.CellValues[ColSal,Ou];
     SetControlCaption('GRBXSALARIE',GAug.CellValues[ColNom,Ou]);
{
     Q :=OpenSQL('SELECT PSA_DATENAISSANCE,PSA_DATEENTREE,PSA_UNITEPRISEFF,PSA_HORAIREMOIS,PSA_ETABLISSEMENT,PSA_CODESTAT,PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4 FROM SALARIES WHERE PSA_SALARIE="'+Salarie+'"',True);
     If Not Q.Eof then
     begin
          DateAnc := Q.FindField('PSA_DATEENTREE').AsDateTime;
          DateNaiss := Q.FindField('PSA_DATENAISSANCE').AsDateTime;
          Etab := Q.FindField('PSA_ETABLISSEMENT').AsString;
          CodeStat := Q.FindField('PSA_CODESTAT').AsString;
          TN1 := Q.FindField('PSA_TRAVAILN1').AsString;
          TN2 := Q.FindField('PSA_TRAVAILN2').AsString;
          TN3 := Q.FindField('PSA_TRAVAILN3').AsString;
          TN4 := Q.FindField('PSA_TRAVAILN4').AsString;
          HoraireMens := Q.FindField('PSA_HORAIREMOIS').AsFloat;
          Effectif := Q.FindField('PSA_UNITEPRISEFF').AsInteger;
     end
     else
     begin
          DateAnc := IDate1900;
          DateNaiss := IDate1900;
          Etab := '';
          Codestat := '';
          TN1 := '';
          TN2 := '';
          TN3 := '';
          TN4 := '';
          HoraireMens := 0;
          Effectif := 0;
     end;
     Ferme(Q);
     Anciennete := AncienneteMois(DateANC, Date);
     AncAnnee := Anciennete div 12;
     AncMois := Anciennete - AncAnnee;
     Age := 0;
     PremMois := 0;
     PremAnnee := 0;
     If DateNaiss > IDate1900 then AglNombreDeMoisComplet(DateNaiss, Date, PremMois, PremAnnee, Age)
     else Age := 0;
     Age := Age div 12;
     SetControlCaption('ANCIENNETE','Ancienneté : '+IntToStr(AncAnnee) +' an(s) et '+ IntToStr(AncMois)+ ' Mois');
     SetControlCaption('NAISSANCE','Age : '+IntToStr(Age));
     SetControlCaption('ETABLISS','Etablissement : '+RechDom('TTETABLISSEMENT',Etab,False));
     SetControlCaption('HORAIRE','Horaire mensuel : '+FloatToStr(HoraireMens)+' heures');
     SetControlCaption('EFFECTIF','Pris dans l''effectif pour : '+IntToStr(Effectif));
     if VH_Paie.PGLibCodeStat <> '' then SetControlCaption('CODESTAT',VH_Paie.PGLibCodeStat + ' '+RechDom('PGCODESTAT',Codestat,False))
     else SetControlCaption('CODESTAT','');
     For i := 1 to 4 do
     begin
          If i = 1 then
          begin
               valeur := TN1;
               Libelle := VH_Paie.PGLibelleOrgStat1;
          end
          else if i = 2 then
          begin
               Valeur := TN2;
               Libelle := VH_Paie.PGLibelleOrgStat2;
          end
          else if i = 3 then
          begin
               Valeur := TN3;
               Libelle := VH_Paie.PGLibelleOrgStat3;
          end
          else
          begin
               Valeur := TN4;
               Libelle := VH_Paie.PGLibelleOrgStat4;
          end;
          if i <= VH_Paie.PGNbreStatOrg then SetControlCaption('TRAVAILN'+IntToStr(i),Libelle + ' : '+RechDom('PGTRAVAILN'+IntToStr(i),Valeur,False))
          else SetControlCaption('TRAVAILN'+IntToStr(i),'');
     end;          }
     SetControlCaption('NAISSANCE','Ancienneté : '+ GAug.CellValues[ColAnc,Ou]+', âge : '+GAug.CellValues[ColAge,Ou]);
     SetControlCaption('ANCIENNETE','Date d''entrée : '+GAug.CellValues[ColDateE,Ou]);
     SetControlCaption('ETABLISS','Etablissement : '+GAug.CellValues[ColEtab,Ou]);
     LeLabel := THLabel(GetControl('EFFECTIF'));
     If GAug.CellValues[ColEff,Ou] <> '1,00' then
     begin
          SetControlCaption('EFFECTIF','Pris dans l''effectif pour '+GAug.CellValues[ColEff,Ou]+', temps partiel');
          If LeLabel <> Nil then LeLabel.Font.Style := [fsBold];
     end
     else
     begin
          SetControlCaption('EFFECTIF','Pris dans l''effectif pour '+GAug.CellValues[ColEff,Ou]);
          If LeLabel <> Nil then LeLabel.Font.Style := [];
     end;
     SetControlCaption('HORAIRE',GAug.CellValues[ColHor,Ou]+ ' heures');
     SetControlCaption('TRAVAILN1',GAug.CellValues[ColTN1,Ou]);
     SetControlCaption('TRAVAILN2',GAug.CellValues[ColTN2,Ou]);
     SetControlCaption('TRAVAILN3',GAug.CellValues[ColTN3,Ou]);
     SetControlCaption('TRAVAILN4',GAug.CellValues[ColTN4,Ou]);
     SetControlCaption('CODESTAT',GAug.CellValues[ColStat,Ou]);
end;

Procedure TOF_PGAUGMENTATIONSAL.ValiderEtat(Sender : Tobject);
Var EtatAug,EtatMax,EtatI : String;
    i : Integer;
begin
        If TypeSaisie='VALIDRESP' then
        begin
             EtatAug := '003';
             EtatMax := '006';
        end;
        If TypeSaisie='VALIDDRH' then
        begin
             EtatAug := '006';
             EtatMax := '009';
        end;
        If GAug.nbSelected = 0 then
        begin
             PGIBox('Opération impossible car aucun salarié sélectionné',Ecran.Caption);
             Exit;
        end;
        For i := 1 to GAug.RowCount - 1 do
        begin
                If GAug.IsSelected(i) then
                begin
                     EtatI := GAug.Cellvalues[ColEtatBis,i];
                     If EtatI < EtatMax then GAug.CellValues[ColEtat,i] := EtatAug;
                end;
        end;
        SetControlProperty('BSelectAll','Down',False);
        GAug.ClearSelected;
        GAug.AllSelected := False;
end;


Procedure TOF_PGAUGMENTATIONSAL.RefuserEtat(Sender : Tobject);
Var EtatAug,EtatMax,EtatI : String;
    i : Integer;
begin
        If TypeSaisie='VALIDRESP' then
        begin
             EtatAug := '004';
             EtatMax := '006';
        end;
        If TypeSaisie='VALIDDRH' then
        begin
             EtatAug := '007';
             EtatMax := '009';
        end;
        If GAug.nbSelected = 0 then
        begin
             PGIBox('Opération impossible car aucun salarié sélectionné',Ecran.Caption);
             Exit;
        end;
        For i := 1 to GAug.RowCount - 1 do
        begin
                If GAug.IsSelected(i) then
                begin
                     EtatI := GAug.Cellvalues[ColEtatBis,i];
                     If EtatI < EtatMax then GAug.CellValues[ColEtat,i] := EtatAug;
                end;
        end;
        SetControlProperty('BSelectAll','Down',False);
        GAug.ClearSelected;
        GAug.AllSelected := False;
end;

procedure TOF_PGAUGMENTATIONSAL.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var OkG, Vide: Boolean;
begin
  OkG := (Screen.ActiveControl = GAug);
  Vide := (Shift = []);
  case Key of
    VK_RETURN: if ((OkG) and (Vide)) then Key := VK_TAB;
    VK_SPACE:
    begin
         //If GAUg.IsSelected(GAUg.Row) then GAUG.FlipSelection[
         If GAug.Focused then GAug.FlipSelection(GAug.Row);
    end;
    38 : If Not Vide then If GAug.Focused then GAug.FlipSelection(GAug.Row);
    40 : If Not Vide then If GAug.Focused then GAug.FlipSelection(GAug.Row);
      end;
 inherited;
end;

procedure TOF_PGAUGMENTATIONSAL.GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  ZoneSuivanteouOk(ACol, ARow, Cancel);
  GAug.SetFocus;
end;

procedure TOF_PGAUGMENTATIONSAL.ZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: boolean);
var
  Sens, ii: integer;
  OldEna: boolean;
begin
  OldEna := GAug.SynEnabled;
  GAug.SynEnabled := False;
  Sens := -1;
  if GAug.Row > ARow then Sens := 1 else if ((GAug.Row = ARow) and (ACol < GAug.Col)) then Sens := 1;
  if (sens = -1) and ((ACol = ColSal) or (GAug.Col = ColSal)) and (ARow = 1) then
  begin
    GAug.SynEnabled := OldEna;
    Cancel := TRUE;
    exit;
  end;
  ACol := GAug.Col;
  ARow := GAug.Row;
  ii := 0;
  while not ZoneAccessible(ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 1000 then Break;
    if (ACol = ColSal) and (ARow = 1) then
    begin
      ACol := 1;
      break;
    end;
    if Sens = 1 then
    begin
      if ((ACol = GAug.ColCount - 1) and (ARow = GAug.RowCount - 1)) then
      begin
        ACol := GAug.FixedCols;
        ARow := 1;
        Break;
      end;
      if ACol < GAug.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        ACol := GAug.FixedCols;
      end;
    end else
    begin
      if ((ACol = GAug.FixedCols) and (ARow = 1)) then Break;
      if ACol > GAug.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := GAug.ColCount - 1;
      end;
    end;
  end;
  GAug.SynEnabled := OldEna;
end;

function TOF_PGAUGMENTATIONSAL.ZoneAccessible(var ACol, ARow: Longint): Boolean;
var
  T1: TOB;
  Etat: string;
  EtatOk : Boolean;
  Effectif : Double;
begin
  result := FALSE;
  EtatOk := True;
  If ACol = ColNom then
  begin
       Result := True;
       Exit;
  end;
  Effectif := StrToFloat(GAug.CellValues[ColEff,ARow]);
  If (Effectif <> 1) and (effectif<> 0)then
  begin
       If TToolBarButton97(GetControl('BTEMPS')).Down = True then
       begin
            Result := False;
            Exit;
       end;
  end;
  Etat := GAug.CellValues[ColEtatBis,ARow];
  If ((TypeSaisie = 'SAISIE') Or (TypeSaisie = 'PROPOSITION')) then //PT1
  begin
       If ACol = ColEtat then EtatOk := False;
       If (Etat <> '001') and (Etat <>'002') and (Etat <>'000') then EtatOk := False;
       if (COnsultation) and (Etat >= '002') then EtatOk := False;
  end;
  If TypeSaisie= 'VALIDRESP' then
  begin
       If Etat >= '006' then EtatOk := False;
  end;
  if (((Acol = ColPctF) or (Acol = ColFap)) and GestionSalFixe) or (((Acol = ColPctV) or (Acol = ColVap)) and GestionSalVariable) or (Acol = ColEtat) then
  begin
       If EtatOk = False then Result := False
       else Result := True;
       If (Not Consultation) and (ACol=ColEtat) and (Etat<'003') then Result := True;
       Exit;
  end;
  If (ACol = ColMotif) and (EtatOk) then
  begin
       Result := True;
       Exit;
  end;
  T1 := TOB(GAug.Objects[ColSal, GAug.Row]);
  if T1 = nil then
  begin
    result := FALSE;
    exit;
  end;
end;

procedure TOF_PGAUGMENTATIONSAL.VerifierModifs(Champ : String;MontantVal : Double;Ligne : Integer);
var Etat,EtatInit,Salarie : String;
    Colonne : Integer;
    T : Tob;
begin
     If TypeSaisie='VALIDDRH' then Etat := '008'
     else If TypeSaisie='VALIDRESP' then Etat := '005'
     else
     begin
          If Consultation then Etat := '000'
          else Etat := '002';
     end;
     EtatInit := GAug.CellValues[ColEtatBis,Ligne];
     If EtatInit = Etat then exit;
     Colonne := -1;
     If Champ = 'PBG_PCTFIXE' then Colonne := ColPctF;
     If Champ = 'PBG_PCTVARIABLE' then Colonne := ColPctV;
     If Champ = 'PBG_FIXEAP' then Colonne := ColFap;
     If Champ = 'PBG_VARIABLEAP' then Colonne := ColVap;
     If Colonne = -1 then Exit;
     Salarie := GAug.CellValues[ ColSal,Ligne];
     T := TobAug.FindFirst(['PSA_SALARIE'],[Salarie],False);
     If T <> Nil then
     begin
          If T.GetDouble(Champ) <> MontantVal then
            GAug.CellValues[ColEtat,Ligne] := Etat;
     end;
end;

procedure TOF_PGAUGMENTATIONSAL.EntrerGrille(Sender : TObject);
var Ligne : Integer;
    Cancel : Boolean;
begin
     If NewSaisie then
     begin
          Ligne := GAug.Row;
          Cancel := False;
          GrilleRowEnter(GAug,Ligne,Cancel,False);
          NewSaisie := False;
     end;
end;

procedure TOF_PGAUGMENTATIONSAL.ImprimerGrille(Sender : Tobject);
Var Pages : TPageControl;
    StPages : String;
begin
     NextPrevControl(TFVierge(Ecran));
     SetControlChecked('FLISTE',False);
     Pages := TPageControl(GetControl('PAGES'));
     StPages := AglGetCriteres (Pages, FALSE);
     TobAug.GetGridDetail (GAug,GAug.RowCount-1,'','INDICATEUR;'+ StChampGrid);
     LanceEtatTOB('E','PAU','PAT',TobAug,True,False,False,Pages,'','',False);
end;

Procedure TOF_PGAUGMENTATIONSAL.ExporterGrille(Sender : TObject);
Var Fichier : String;
    TobTemp,T : Tob;
    i : Integer;
begin
     NextPrevControl(TFVierge(Ecran));
{     SetControlChecked('FLISTE',True);
     Pages := TPageControl(GetControl('PAGES'));
     StPages := AglGetCriteres (Pages, FALSE);
     TobAug.GetGridDetail (GAug,GAug.RowCount-1,'',StChampGrid);
     LanceEtatTOB('E','PAU','PAT',TobAug,True,True,False,Pages,'','',False,0,StPages)}
   //  LanceEtat('E','PAU','PAU',True,True,False,Pages,'','',False,0,StPages)
     Fichier := AGLLanceFiche('PAY','AUGM_XLS','','','');
     If Fichier <> '' then
     begin
          TobAug.GetGridDetail (GAug,GAug.RowCount-1,'','INDICATEUR;'+ StChampGrid);
          TobTemp := Tob.Create('Temporaire',Nil,-1);
          For i := 0 to TobAug.Detail.Count -1 do
          begin
               T := Tob.Create('Fille_temp',TobTemp,-1);
               T.AddChampSupValeur('SALARIE',TobAug.Detail[i].GetValue('PSA_LIBELLE'),False);
               T.AddChampSupValeur('EMPLOI',TobAug.Detail[i].GetValue('PSA_LIBELLEEMPLOI'),False);
               T.AddChampSupValeur('DATE_ENTREE',TobAug.Detail[i].GetValue('PSA_DATEENTREE'),False);
               T.AddChampSupValeur('ETAT',RechDom('PGAUGMETATVALID',TobAug.Detail[i].GetValue('PBG_ETATINTAUGM'),False),False);
               T.AddChampSupValeur('FIXE',TobAug.Detail[i].GetValue('PBG_FIXEAV'),False);
               T.AddChampSupValeur('POURCENTAGE_FIXE',TobAug.Detail[i].GetValue('PBG_PCTFIXE'),False);
               T.AddChampSupValeur('FIXE_AUGMENTE',TobAug.Detail[i].GetValue('PBG_FIXEAP'),False);
               T.AddChampSupValeur('VARIABLE',TobAug.Detail[i].GetValue('PBG_VARIABLEAV'),False);
               T.AddChampSupValeur('POURCENTAGE_VARIABLE',TobAug.Detail[i].GetValue('PBG_PCTVARIABLE'),False);
               T.AddChampSupValeur('VARIABLE_AUGMENTE',TobAug.Detail[i].GetValue('PBG_VARIABLEAP'),False);
               T.AddChampSupValeur('BRUT',TobAug.Detail[i].GetValue('TOTALAV'),False);
               T.AddChampSupValeur('POURCENTAGE_BRUT',TobAug.Detail[i].GetValue('PCTTOTAL'),False);
               T.AddChampSupValeur('BRUT_AUGMENTE',TobAug.Detail[i].GetValue('TOTALAP'),False);
               T.AddChampSupValeur('PBG_COMMENTAIREABR',TobAug.Detail[i].GetValue('TOTALAP'),False);
          end;
          TobTemp.SaveToExcelFile(Fichier);
          tobTemp.Free;
    //      Rep := PGIAsk('Voulez vous ouvrir le fichier','Exportation de la saisie');
    //      If rep = MrYes then ShellExecute(0, PCHAR('open'), PChar('Excel'), PChar(Fichier), nil, SW_RESTORE);
     end;
end;

Procedure TOF_PGAUGMENTATIONSAL.AnnulerModifs(Sender : TObject);
begin
     TobAug.PutGridDetail(GAug,False,False,StChampGrid,False);
     GAug.SortedCol := -1;
end;

procedure TOF_PGAUGMENTATIONSAL.AfficheLegende(Sender : TObject);
begin
     AGLLanceFiche('PAY','AUGMLEGENDE','','','');
end;

Function TOF_PGAUGMENTATIONSAL.ArrondiAugm(Montant : Double) : Double;
var Calcul : Double;
begin
     If IArrondiAugm >= 0 then
     begin
          If PrecisionArrondi = 5 then Calcul := 5 * (Arrondi((Montant/5),IArrondiAugm))
          else Calcul := Arrondi(Montant,IArrondiAugm);
     end
     else
     begin
          Calcul := 10 * (Arrondi((Montant/10),0));
     end;
     Result := Calcul;
end;

{$IFDEF EMANAGER}
Procedure TOF_PGAUGMENTATIONSAL.EnvoyerMailResp(Sender : TObject);
var Texte: HTStrings;
    Titre, Destinataire: string;
    ResponsValid,Service : string;
    Q: TQuery;
begin
  Destinataire := '';
  Q := OpenSQL('SELECT PSE_CODESERVICE FROM DEPORTSAL WHERE PSE_SALARIE="'+V_PGI.UserSalarie+'"',True);
  If Not Q.Eof then Service := Q.FindField('PSE_CODESERVICE').AsString
  else Service := '';
  Ferme(Q);
  If Service <> '' then
  begin
       Q := OpenSQL('SELECT PGS_RESPONSVAR FROM SERVICES '+
       'LEFT JOIN HIERARCHIE ON PGS_HIERARCHIE=PHO_HIERARCHIE '+
       'LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP '+
       'WHERE PHO_NIVEAUH<=2 AND PSO_CODESERVICE="'+Service+'"',True);
       If Not Q.Eof then ResponsValid := Q.FindField('PGS_RESPONSVAR').AsString
       else ResponsValid := '';
  end     
  else ResponsValid := '';
  Ferme(Q);
  If ResponsValid <> '' then
  begin
       Q := OpenSQL('SELECT PSE_EMAILPROF FROM DEPORTSAL WHERE PSE_SALARIE="' + ResponsValid + '"', true);
       if not Q.eof then Destinataire := Q.FindField('PSE_EMAILPROF').AsString; //portageCWAS
       Ferme(Q);
  end;
  Titre := 'Saisie des augmentations'; //+RechDom('PGTYPVISITEMED',GetField('PGTYPVISITEMED'),false);
  Texte := HTStringList.Create;
  texte.Insert(0, 'Le responsable des éléments variables '+RechDom('PGSALARIE',V_PGI.UserSalarie,False)+' a terminé la saisie de ses augmentations.');
  texte.Insert(1, 'Merci de procéder à la validation sur EManager');
  texte.Insert(2, '');
  texte.Insert(3, 'Cordialement,');
  SendMail(Titre, Destinataire, '', Texte, '', false);
  Texte.free;

end;
{$ENDIF}

procedure TOF_PGAUGMENTATIONSAL.VoirTempsPlein(Sender : TObject);
var Effectif,Coeff : Double;
    i : Integer;
    SalApF,SalAvF,SalApV,SalAvV : Double;
begin
     If TToolBarButton97(Sender).Down = True then
     begin
          TToolBarButton97(Sender).Hint := 'Voir salaires temps partiel';
          For i := 1 to GAug.RowCount - 1 do
          begin
               Effectif := StrToFloat(GAug.CellValues[ColEff,i]);
               If (Effectif <> 1) and (effectif<> 0)then
               begin
                    Coeff := 1 / Effectif;
                    salApF := FormatageDouble(GAug.Cellvalues[ColFap,i]);
                    salAvF := FormatageDouble(GAug.Cellvalues[ColF,i]);
                    SalApF := SalApF * Coeff;
                    salAvF := SalAvF * Coeff;
                    GAug.CellValues[ColF,i] := FormatFloat('# ##0.00',salAvF);
                    GAug.CellValues[ColFap,i] := FormatFloat('# ##0.00',salApF);
                    salApV := FormatageDouble(GAug.Cellvalues[ColVap,i]);
                    salAvV := FormatageDouble(GAug.Cellvalues[ColV,i]);
                    SalApV := SalApV * Coeff;
                    salAvV := SalAvV * Coeff;
                    GAug.CellValues[ColV,i] := FormatFloat('# ##0.00',salAvV);
                    GAug.CellValues[ColVap,i] := FormatFloat('# ##0.00',salAvV);
                    GAug.CellValues[ColBrut,i] := FormatFloat('# ##0.00',salAvV + SalAvF);
                    GAug.CellValues[ColBrutap,i] := FormatFloat('# ##0.00',salApV + SalApF);
                    GAug.CellValues[ColSalAnnAp,i] := FormatFloat('# ##0',SalApV*12 + SalApF*NbMEquivalenceAnn);
                    GAug.CellValues[ColTpsPlein,i] := FormatFloat('# ##0.00',salApV + SalApF);
                    GAug.CellValues[ColSalAnnTpsP,i] := FormatFloat('# ##0',(SalApV*12 + SalApF*NbMEquivalenceAnn) * Coeff);
               end;
          end;
     end
     else
     begin
          TToolBarButton97(Sender).Hint := 'Voir équivalence temps plein';
          For i := 1 to GAug.RowCount - 1 do
          begin
               Effectif := StrToFloat(GAug.CellValues[ColEff,i]);
               If (Effectif <> 1) and (effectif<> 0)then
               begin
                    Coeff := 1 / Effectif;
                    salApF := FormatageDouble(GAug.Cellvalues[ColFap,i]);
                    salAvF := FormatageDouble(GAug.Cellvalues[ColF,i]);
                    SalApF := SalApF / Coeff;
                    salAvF := SalAvF / Coeff;
                    GAug.CellValues[ColF,i] := FormatFloat('# ##0.00',salAvF);
                    GAug.CellValues[ColFap,i] := FormatFloat('# ##0.00',salApF);
                    salApV := FormatageDouble(GAug.Cellvalues[ColVap,i]);
                    salAvV := FormatageDouble(GAug.Cellvalues[ColV,i]);
                    SalApV := SalApV / Coeff;
                    salAvV := SalAvV / Coeff;
                    GAug.CellValues[ColV,i] := FormatFloat('# ##0.00',salAvV);
                    GAug.CellValues[ColVap,i] := FormatFloat('# ##0.00',salAvV);
                    GAug.CellValues[ColBrut,i] := FormatFloat('# ##0.00',salAvV + SalAvF);
                    GAug.CellValues[ColBrutap,i] := FormatFloat('# ##0.00',salApV + SalApF);
                    GAug.CellValues[ColSalAnnAp,i] := FormatFloat('# ##0',salApF*NbMEquivalenceAnn + salApV*12);
                    GAug.CellValues[ColTpsPlein,i] := FormatFloat('# ##0.00',Coeff *  (salApV + SalApF));
                    GAug.CellValues[ColSalAnnTpsP,i] := FormatFloat('# ##0',(salApF*NbMEquivalenceAnn + salApV*12) * Coeff);
               end;
          end;
     end;
     CalculerResume(-1,True);
end;

Initialization
  registerclasses ( [ TOF_PGAUGMENTATIONSAL ] ) ;
end.

