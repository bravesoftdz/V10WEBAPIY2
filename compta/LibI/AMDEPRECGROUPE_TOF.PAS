{***********UNITE*************************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 16/09/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AMDEPRECGROUPE ()
Mots clefs ... : TOF;AMDEPRECGROUPE
Suite..........: FQ 17183 BTY 12/05 Transformer en fonction pour avoir un retour d'opération et pouvoir sortir de la fiche immo
Suite..........: FQ 17171 BTY 01/06 Message d'erreur si pb durant la transaction d'enregistrement
Suite..........: FQ 17259 BTY 01/06 Nelle colonne dans IMMO
Suite..........: FQ 17476 MBO 02/06 Stockage de la dépréciation sur le plan fiscal
Suite..........: FQ 18119 BTY 05/06 Positionner les indicateurs de modif de compta du PARAMSOC
Suite..........: FQ 18386 BTY 06/06 Ne pas déprécier les immos de date début amortissement ECO > date fin exo en cours
Suite..........: FQ 17500 BTY 06/06 Prise en compte de l'antérieur dépréciation saisi
Suite..........: BTY 06/06 Message bloquant si aucune immo du regroupement n'est dépréciable (dépréciées, cédées, date début amort ECO > date fin exo en cours)
Suite..........: MBO     25/06/2007 Ne pas présenter les immos ayant une gestion fiscale    
*****************************************************************}
Unit AMDEPRECGROUPE_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     windows,
     Graphics,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     MaineAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     ImPlan,
     HRichOle,
     uTOB,
     AGLInit,
     UTOF,
     Vierge;

Type
  TEnregDeprecGroupe = class
    Code : string;
    //JournalA : string;
    Operation : string;
    LibelleLog : string;
    TypeOpe : string;
    DateOpe : TDateTime;
    TypeDot : string;
    Ordre : integer;
    OrdreSerie : integer;
    PlanActifAv : integer;
    PlanActifAp : integer;
    MontantDot : double;
    //CodeMutation : string;
    BlocNote : String;
    OpeDeprec : string;
  end;

  TOF_AMDEPRECGROUPE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (StCodeRegroupement : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
      fRegroupement : String;
      fTRepartition : TOB;
      fGrille : THGrid;
      fCurMontant : double;
      fMontantOpe : double;
      fInitZero : double;
      fCumulPlafondAjout : double;
      fCumulPlafondReprise : double;
      fOrdreSerie : integer;
      DP_BlocNote : THRichEditOle;
      fDateErr : boolean;
      fMontantErr : boolean;
      fCellErr : boolean;
      fConfirme: boolean;
      fVide : boolean;   // BTY 06/06 Message bloquant si grille vide
      //procedure OnAfterFormShowEcran;
      procedure OnbFermeClick(Sender: TObject);
      procedure OnClickRadioB(Sender: TObject);
      procedure OnChangeRepartition(Sender: TObject);
      //procedure OnKeyPressDateOpe(Sender: TObject);
      //procedure OnExitDateOpe(Sender: TObject);
      // BTY 06/06 Message bloquant si grille vide
      procedure OnEnterDateOpe(Sender: TObject);
      procedure OnExitMontantOpe(Sender: TObject);
      procedure FormaterfGrille;
      procedure InitZonesfGrille;
      procedure CalculerRepartitionEtRecharger;
      procedure RechargerGrille;
      procedure AffichPlafond;
      procedure MontantOpeRestant;
      function  ControleMontantOpeOK : boolean;
      function  ControleZonesOK : boolean;
      function  ControleDateOpeOK : boolean;
      function  ControleSaisieOK : boolean;
      procedure OnGrilleCellExit(Sender: TObject; var ACol,
        ARow: Integer; var Cancel: Boolean);
      procedure OnGrilleCellEnter(Sender: TObject; var ACol,
        ARow: Integer; var Cancel: Boolean);
      procedure OnGrilleRowEnter(Sender: TObject;
        Ou: Integer; var Cancel: Boolean; Chg: Boolean);
      procedure OnGrilleKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
      procedure EnregistrerDepreciationImmo;
      procedure EnregUneImmo (rang : integer; stCode : string);
      procedure MAJUneImmo (EnregImmo : TEnregDeprecGroupe) ;
      procedure AjoutUneImmoLog (EnregImmo : TEnregDeprecGroupe);
end ;

const
      DP_COL_IMMO   = 0;
      DP_COL_LIB    = 1;
      DP_COL_VNC    = 2;
      DP_COL_PLAFOND = 3;
      DP_COL_DEPREC = 4;

//Procedure AMLanceFiche_DeprecGroupe(stCodeRegroupement: string);
Function AMLanceFiche_DeprecGroupe(stCodeRegroupement: string):String;

Implementation

uses

Messages,  // pour WM_CLOSE BTY 06/06 Message bloquant si grille vide
Iment,     // pour VHImmo
Outils;    // pour n° ordre d'opération
//ImouPlan;  // pour NombreJour360


const
HMD: array[1..10] of string =
 {1} ('1;Dépréciation;Voulez-vous enregistrer cette opération ?;Q;YNC;N;C;',
 {2} '2;Dépréciation;La date de l''opération est incorrecte.;E;O;O;O;',
 {3} '3;Dépréciation;La date d''opération doit être une date de l''exercice en cours;E;O;O;O;',
 {4} '4;Dépréciation;Une opération a déjà été effectuée à une date postérieure.;E;O;O;O;',
 {5} '5;Dépréciation;Le montant de l''opération est supérieur au plafond.;E;O;O;O;',
 {6} '6;Dépréciation;Le montant saisi est supérieur au plafond de l''immobilisation.;E;O;O;O;',
 {7} '7;Dépréciation;Le total des répartitions est différent du montant de l''opération.;E;O;O;O;',
 {8} '8;Dépréciation;Le montant de l''une des répartitions est supérieur au plafond de l''immobilisation.#10#13Veuillez rectifier en répartition libre.;E;O;O;O;',
 {9} '9;Dépréciation;Après l''enregistrement de la dépréciation, toute modification rétrospective des plans d''amortissement sera interdite. Voulez-vous continuer ?;Q;YN;N;C;',
 {10} '10;Dépréciation;L''enregistrement de la dépréciation a échoué!;W;O;O;O;'
     );


// BTY 12/05 FQ 17183
{*procedure AMLanceFiche_DeprecGroupe(StCodeRegroupement: string);
 begin
   AGLLanceFiche('AM', 'AMDEPRECGROUPE','' , '', StCodeRegroupement);
 end;*}

function AMLanceFiche_DeprecGroupe(StCodeRegroupement: string):String;
 begin
   result:= AGLLanceFiche('AM', 'AMDEPRECGROUPE','' , '', StCodeRegroupement);
 end;


procedure TOF_AMDEPRECGROUPE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AMDEPRECGROUPE.OnDelete ;
begin
  Inherited ;
end ;

// Action Valider l'écran
procedure TOF_AMDEPRECGROUPE.OnUpdate ;
var NumColEnCours,NumRowEnCours : integer;
    Sender: Tobject;
    Verif: boolean;
begin
  Inherited ;
  // Ne pas sortir a priori
  Ecran.ModalResult := mrNone;
  // Forcer le OnExit sous le curseur
  // NextPrevControl(Ecran);
  // OnExitMontant a été envoyé => fMontantErr positionné
  if fMontantErr then exit;
  // Vérifier la valeur de la cellule sous le curseur en saisie libre
  Sender:=nil;
  NumColEnCours := FGrille.Col;
  NumRowEnCours := FGrille.Row;
  OnGrilleCellExit (Sender,NumColEnCours,NumRowEnCours,Verif);
  if fCellErr then exit;
  // Vérifier la date
  if ControleDateOpeOK then fDateErr := False else fDateErr := True;
  if fDateErr then exit;
  // En dernier : contrôler que le total réparti est égal au montant
  if ControleSaisieOK then
     begin
     // Enregistrer l'opération et sortir
     EnregistrerDepreciationImmo;
     Ecran.ModalResult := mrYes;
     end;
end ;


procedure TOF_AMDEPRECGROUPE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AMDEPRECGROUPE.OnArgument (StCodeRegroupement : String ) ;

begin
  Inherited ;
  // Initialisations
  fRegroupement :=  StCodeRegroupement;
  fMontantOpe := 0;
  fCumulPlafondAjout := 0;
  fCumulPlafondReprise := 0;
  fInitZero := 0;
  fConfirme := False;
  // BTY 12/05 FQ 17183
  TFVIERGE(ECRAN).retour := IntToStr(MrNo);

  // Initialisation des champs
  SetControlCaption ('DP_Regroupement', fRegroupement+' - '+RechDom('AMREGROUPEMENT',fRegroupement,False));
  SetControlProperty('DP_Plafond', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('DP_MontantOpe', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('DP_Repartition', 'DataType', 'AMREPARTITIONVNC');
  SetControlProperty('DP_Repartition', 'Value', 'VNC');
  SetControlEnabled('DP_Plafond', False);
  SetControlProperty('DP_MontantOpe', 'Text', StrFMontant(fMOntantOpe,15 , V_PGI.OkDecV, '' , true));
  DP_BlocNote := THRichEditOle (GetControl('DP_BN'));

  // Evènements
  TToolbarButton97(GetControl('BFerme')).Onclick := OnbFermeClick;
  THRadioGroup(GetControl('DP_RadioB')).Onclick := OnClickRadioB;
  THValComboBox(GetControl('DP_Repartition')).OnChange  := OnChangeRepartition;
  //THEdit(GetControl('DP_DateOpe')).OnKeyPress  := OnKeyPressDateOpe;
  //THEdit(GetControl('DP_DateOpe')).OnExit    := OnExitDateOpe;
  THEdit(GetControl('DP_MontantOpe')).OnExit    := OnExitMontantOpe;

  // FQ 18386
  THEdit(GetControl('DP_DateOpe')).OnEnter := OnEnterDateOpe;

  // Initialisation du tableau
  fGrille := THGrid ( GetControl ('DP_Liste') );
  fGrille.OnCellExit  := OnGrilleCellExit;
  fGrille.OnCellEnter  := OnGrilleCellEnter;
  fGrille.OnRowEnter := OnGrilleRowEnter;
  fGrille.OnKeyDown := OnGrilleKeyDown;
  FormaterfGrille;

  // Chargement du tableau via des TOB
  InitZonesfGrille;

  fCurMontant := Valeur (fGrille.Cells[DP_COL_DEPREC,1]);

  //OnAfterFormShowEcran;
end ;

procedure TOF_AMDEPRECGROUPE.OnClose;
begin
  Inherited ;
  fTRepartition.Free;
end ;

procedure TOF_AMDEPRECGROUPE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AMDEPRECGROUPE.OnCancel () ;
begin
  Inherited ;
end ;


//*******************************************************************************
//    Evènements supplémentaires
//*******************************************************************************


{
procedure TOF_AMDEPRECGroupe.OnExitDateOpe(Sender: TObject);
begin
  if ControleDateOpeOK then fDateErr := False else fDateErr :=True;
end;
}

procedure TOF_AMDEPRECGroupe.OnExitMontantOpe(Sender: TObject);
begin
   // Le montant a changé
   //if fMontantOpe <> Valeur (GetControlText('DP_MontantOpe')) then
   //   ControleMontantOpeOK;
   if ControleMontantOpeOK then fMontantErr :=False else fMontantErr :=True;
end;


procedure TOF_AMDEPRECGroupe.OnChangeRepartition(Sender: TObject);
begin
  // Recalcul des répartitions si Prorata retenu
  if GetControlText('DP_Repartition')<>'ZLI' then
     begin
     RechargerGrille;
     end;
  //Ouvrir la colonne Dépréciation à la saisie si répartition libre
  fGrille.ColEditables[DP_COL_DEPREC]:=(GetControlText('DP_Repartition')='ZLI');

end;


procedure TOF_AMDEPRECGroupe.OnClickRadioB(Sender: TObject);
begin
 // Mode de répartition et Plafond diffèrent selon AJOut ou REPrise
 if (GetControlText('DP_RadioB')= 'AJO') then
    begin
    SetControlProperty('DP_Repartition', 'DataType', 'AMREPARTITIONVNC');
    SetControlProperty('DP_Repartition', 'Value', 'VNC');
    fGrille.Cells[DP_COL_DEPREC,0] := TraduireMemoire('Dépréciation');
    end
 else
    begin
    SetControlProperty('DP_Repartition', 'DataType', 'AMREPARTITIONPLAFOND');
    SetControlProperty('DP_Repartition', 'Value', 'PLA');
    fGrille.Cells[DP_COL_DEPREC,0] := TraduireMemoire('Reprise');
    end;
 RechargerGrille;
 SetControlProperty('DP_MontantOpe', 'Text', StrFMontant( fInitZero, 15 , V_PGI.OkDecV, '' , true));
 fMontantOpe := Valeur (GetControlText('DP_MontantOpe'));
end;


procedure TOF_AMDEPRECGroupe.OnGrilleCellExit(Sender: TObject; var ACol, ARow: Integer;
          var Cancel: Boolean);
var RdMontant : double;
begin
  inherited;
  fCellErr := False;
  case ACol of
    DP_COL_DEPREC :
    begin
      if IsNumeric(fGrille.Cells[ACol,ARow]) then
      begin
        RdMontant := Valeur(fGrille.Cells[ACol,ARow]);
        // on ne peut pas saisir de valeur negative
        fCellErr  := RdMontant < 0;
        if fCellErr then  begin PGIInfo('Montant négatif !','Attention'); end;
        // test montant dépasse le plafond de l'immo
        if (RdMontant > Valeur (fGrille.Cells[DP_COL_PLAFOND,ARow])) then
            begin
            HShowMessage (HMD[6],'',''); RdMontant := 0; fCellErr :=True;
            end;
        fGrille.Cells[ACol,ARow] := STRFMONTANT (RdMontant , 15 , V_PGI.OkDecV, '' , true);
        fCurMontant := Valeur(FGrille.Cells[ACol,ARow]);
      end;
    end;
  end;
end;


// Touche F6 pour avoir le restant à déprécier
procedure TOF_AMDEPRECGroupe.OnGrilleKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    VK_F6 :
     if (GetControlText('DP_Repartition')='ZLI') then MontantOpeRestant;
  end;
end;


// Action Echap ou Fermer par la croix
procedure TOF_AMDEPRECGroupe.OnbFermeClick(Sender: TObject);
begin

  case HShowMessage (HMD[1],'','') of
    mrYes:
      begin
      if ControleZonesOK then
         begin
         // OK enregistrer l'opération et sortir
         EnregistrerDepreciationImmo;
         Ecran.ModalResult := mrYes;
         end
      else Ecran.ModalResult := mrNone; // ne pas sortir
      end;
    mrNo:
      begin
      // Sortir
      Ecran.ModalResult := mrNo;
      end;
    else
      // Rester dans l'écran
      Ecran.ModalResult := mrNone;
  end;

end;

procedure TOF_AMDEPRECGroupe.OnGrilleCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
  fCurMontant := Valeur (fGrille.Cells[ACol,ARow]);
end;

procedure TOF_AMDEPRECGroupe.OnGrilleRowEnter(Sender: TObject; Ou: Integer;
var Cancel: Boolean; Chg: Boolean);
begin
end;


//*******************************************************************************
//    Fonctions supplémentaires
//*******************************************************************************

procedure TOF_AMDEPRECGroupe.FormaterfGrille;
var Masque : string;
begin
  Masque := StrfMask(V_PGI.OkDecV, '', True);

  fGrille.ColEditables[0]:=False;
  fGrille.ColWidths[0]:=70;
  fGrille.ColAligns[0]:=taLeftJustify;
  fGrille.Cells[DP_COL_IMMO,0] := TraduireMemoire('Code');

  fGrille.ColEditables[DP_COL_LIB]:= False;
  fGrille.ColWidths[DP_COL_LIB]:=100;
  fGrille.ColAligns[DP_COL_LIB]:=taLeftJustify;
  fGrille.Cells[DP_COL_LIB,0] := TraduireMemoire('Libellé');

  fGrille.ColTypes[DP_COL_VNC]:='R';
  fGrille.ColFormats[DP_COL_VNC]:=Masque;
  fGrille.ColEditables[DP_COL_VNC]:= False;
  fGrille.ColWidths[DP_COL_VNC]:=93;
  fGrille.ColAligns[DP_COL_VNC]:=taRightJustify;
  fGrille.Cells[DP_COL_VNC,0] := TraduireMemoire('Valeur résiduelle');

  fGrille.ColTypes[DP_COL_PLAFOND]:='R';
  fGrille.ColFormats[DP_COL_PLAFOND]:=Masque;
  fGrille.ColEditables[DP_COL_PLAFOND]:= False;
  fGrille.ColWidths[DP_COL_PLAFOND]:=93;
  fGrille.ColAligns[DP_COL_PLAFOND]:=taRightJustify;
  fGrille.Cells[DP_COL_PLAFOND,0] := TraduireMemoire('Plafond');

  fGrille.ColTypes[DP_COL_DEPREC]:='R';
  fGrille.ColFormats[DP_COL_DEPREC]:=Masque;
  fGrille.ColEditables[DP_COL_DEPREC]:=(GetControlText('DP_Repartition')='ZLI');
  fGrille.ColWidths[DP_COL_DEPREC]:=93;
  fGrille.ColAligns[DP_COL_DEPREC]:=taRightJustify;
  fGrille.Cells[DP_COL_DEPREC,0] := TraduireMemoire('Dépréciation');
end;


procedure TOF_AMDEPRECGroupe.InitZonesfGrille;
var  Q: TQuery;
     stSQL:string;
begin
  fVide := False;
  fTRepartition := TOB.Create ('',nil,-1);

  //Q := OpenSQL('SELECT * FROM IMMO WHERE I_NATUREIMMO="PRO" AND I_GROUPEIMMO="'+fRegroupement+'"', true);
  //fTRepartition.LoadDetailDB('IMMO', '', '', Q, false);
  // Jointure interne pour exclure les immos ayant ds IMMOLOG l'indication d'une dépréciation dans l'exo
  //stSQL := 'SELECT * FROM IMMO AS a INNER JOIN IMMOLOG AS b ';
  //stSQL := stSQL + 'ON a.I_IMMO=b.IL_IMMO AND a.I_PLANACTIF=b.IL_PLANACTIFAP ';
  //stSQL := stSQL + 'WHERE a.I_NATUREIMMO="PRO" AND a.I_GROUPEIMMO="'+fRegroupement+'" ';
  //stSQL := stSQL + 'AND (b.IL_TYPEOP <> "DPR" OR (b.IL_TYPEOP = "DPR" AND IL_DATEOP<"' + USDateTime(VHImmo^.Encours.Deb)+'" ))';

  // mbo - le 25.06.07 - on ne présente pas les immos ayant une gestion fiscale
  // ajout du test sur i_nonded dans le select

  stSQL := stSQL + 'SELECT * FROM IMMO WHERE I_NATUREIMMO="PRO" AND I_GROUPEIMMO="'+fRegroupement+'" ';
  stSQL := stSQL + 'AND I_NONDED<> "X" AND I_QUANTITE<>0 AND I_DATEDEBECO<="'+ USDateTime(VHImmo.Encours.Fin) + '" AND I_IMMO NOT IN (SELECT IL_IMMO FROM IMMOLOG WHERE IL_TYPEOP = "DPR" ';
  // FQ 18386 Exclure les immos de date début amortissement ECO > date fin exo en cours
  //stSQL := stSQL + 'AND I_QUANTITE<>0 AND I_IMMO NOT IN (SELECT IL_IMMO FROM IMMOLOG WHERE IL_TYPEOP = "DPR" ';
  stSQL := stSQL + 'AND IL_DATEOP>="' + USDateTime(VHImmo^.Encours.Deb)+'" )';
  Q := OpenSQL(StSQL, True);

  // BTY 06/06 Message bloquant si grille vide
  if Q.eof then
     begin
     ferme (Q);
     fVide := True;
     end
  else
     begin
  fTRepartition.LoadDetailDB('', '', '', Q, false);
  ferme(Q);
  CalculerRepartitionEtRecharger;
     end;
end;


procedure TOF_AMDEPRECGroupe.AffichPlafond;
begin
  if (GetControlText('DP_RadioB')= 'AJO') then
    SetControlProperty('DP_Plafond', 'Text', StrFMontant(fCumulPlafondAjout,15 , V_PGI.OkDecV, '' , true))
  else
    SetControlProperty('DP_Plafond', 'Text', StrFMontant(fCumulPlafondReprise,15 , V_PGI.OkDecV, '' , true))
end ;



procedure TOF_AMDEPRECGroupe.CalculerRepartitionEtRecharger;

var i : integer;
    PlafondAjout,PlafondReprise : double;
    CumulEco, CumulFisc : double;
    TPlan : TPlanAmort;
    T : TOB;


begin
  if (fRegroupement='') or (fTRepartition=nil)
      or (fTRepartition.Detail.Count=0) then exit;

  //DureeExo := NombreJour360(VHImmo^.Encours.Deb,VHImmo^.Encours.Fin);
  fConfirme := True;

  TPlan:=TPlanAmort.Create(true) ;
  try
     for i:=0 to fTRepartition.Detail.Count - 1 do
     begin
       //PlafondAjout := 0;
       //PlafondReprise := 0;
       CumulEco := 0;
       CumulFisc := 0;
       //CumulDeprec := 0;
       // Récup TOBs fille
       T := fTRepartition.Detail[i];
       TPlan.ChargeTOB(T);
       TPlan.Recupere(T.GetValue('I_IMMO'),T.GetValue('I_PLANACTIF') );
       TPlan.GetCumulsDotExercice(VHImmo^.Encours.Fin,CumulEco, CumulFisc,false,false,true);
       //CumulDeprec := TPlan.Cumul_depreciation(T.GetValue('I_IMMO'));
       //PlafondAjout := Arrondi (
       //       T.GetValue('I_BASEECO')-T.GetValue('I_REPRISEECO')-CumulEco-CumulDeprec, V_PGI.OkDecV);
       //PlafondReprise := Arrondi (
       //CumulDeprec * (1-(DureeExo/NombreJour360(VHImmo^.Encours.Deb,TPlan.AmortEco.DateFinAmort))), V_PGI.OkDecV);

       //{$IFDEF CRC200210}
       // 10/05 CumulDepre est deja dans CumulEco
       // PlafondAjout := Arrondi (
       //        T.GetValue('I_BASEECO')-T.GetValue('I_REPRISEECO')-CumulEco-TPlan.CumulDepre, V_PGI.OkDecV);

       // BTY FQ 17500 impact antérieur dépréciation, le soustraire aussi
       PlafondAjout := Arrondi (
             TPlan.BaseDebutExoEco-TPlan.AmortEco.Reprise-CumulEco-TPlan.AmortEco.RepriseDep, V_PGI.OkDecV );
                //TPlan.BaseDebutExoEco-TPlan.AmortEco.Reprise-CumulEco, V_PGI.OkDecV );
       /// le plafond est stocké dans immolog par la cloture cf mbo amdepreciation
       //PlafondReprise := Arrondi (
       //         TPlan.CumulDepre * (1-(DureeExo/NombreJour360(VHImmo^.Encours.Deb,TPlan.AmortEco.DateFinAmort))), V_PGI.OkDecV);

       PlafondReprise := Arrondi ( TPlan.CumulDepre, V_PGI.OkDecV);

       // modif mbo - impact antérieur dépréciation - 02.06
       if (PlafondReprise = 0) and (TPlan.AmortEco.RepriseDep <> 0) then
         // il faut recalculer le montant maximum de la reprise
         PlafondReprise := TPlan.CalculRepriseDepreciation(TPlan.AmortEco, VHImmo^.EnCours.Deb,
                                               VHImmo^.EnCours.Fin, false);

       //{$ENDIF}

       fCumulPlafondAjout := Arrondi (fCumulPlafondAjout + PlafondAjout, V_PGI.OkDecV);
       fCumulPlafondReprise := Arrondi (fCumulPlafondReprise + PlafondReprise, V_PGI.OkDecV);
       // Ajout champs dans TOB fille
       fTRepartition.Detail[i].AddChampSupValeur('PLAFONDAJO',PlafondAjout);
       fTRepartition.Detail[i].AddChampSupValeur('PLAFONDREP',PlafondReprise);
       // Indicateur message de confirmation si aucune immo n'a subi de dépréciation
       if TPlan.GestDeprec then fConfirme := False;
     end;

  // Libération mémoire
  finally
     TPlan.free ;
  end ;

  RechargerGrille;
end;


// Calcul des dépréciation/reprise en fonction du plafond en cours et du montant de l'opération
// + rechargement de fGrille à partir des TOB
procedure TOF_AMDEPRECGroupe.RechargerGrille ;

var i : integer;
    Depreciation, Reprise, CumulDepreciation, CumulReprise : double;
begin
  if (fRegroupement='') or (fTRepartition=nil)
      or (fTRepartition.Detail.Count=0) then exit;

  // Par immo, calcul des proratas du plafond
  CumulDepreciation := 0;
  CumulReprise := 0;
  for i:=0 to fTRepartition.Detail.Count - 1 do
  begin
     Depreciation := 0;
     Reprise := 0;
     if i=(fTRepartition.Detail.Count - 1) then
        // dernière cellule = (MontantOpe - somme des répartitions précédentes)
        begin
          if Arrondi(fCumulPlafondAjout, V_PGI.OkDecV) <> 0 then  Depreciation := Arrondi (
             (Valeur (GetControlText('DP_MontantOpe'))-CumulDepreciation), V_PGI.OkDecV);
          if Arrondi(fCumulPlafondReprise, V_PGI.OkDecV) <> 0 then  Reprise := Arrondi (
             (Valeur (GetControlText('DP_MontantOpe'))-CumulReprise), V_PGI.OkDecV);
        end
     else
        begin
        if Arrondi(fCumulPlafondAjout, V_PGI.OkDecV) <> 0 then  Depreciation := Arrondi (
           (Valeur (GetControlText('DP_MontantOpe')) * fTRepartition.Detail[i].GetValue('PLAFONDAJO') / fCumulPlafondAjout), V_PGI.OkDecV);
        if Arrondi(fCumulPlafondReprise, V_PGI.OkDecV) <> 0 then Reprise := Arrondi (
           (Valeur (GetControlText('DP_MontantOpe')) * fTRepartition.Detail[i].GetValue('PLAFONDREP') / fCumulPlafondReprise), V_PGI.OkDecV);
        CumulDepreciation := Arrondi(CumulDepreciation+Depreciation, V_PGI.OkDecV);
        CumulReprise := Arrondi(CumulReprise+Reprise, V_PGI.OkDecV);
        end;
     // Ajout champs dans TOB fille
     fTRepartition.Detail[i].AddChampSupValeur('DEPRECIATION', Depreciation);
     fTRepartition.Detail[i].AddChampSupValeur('REPRISE', Reprise);
  end;

  // Chargement du tableau fGrille à partir des TOBs
  if (GetControlText('DP_RadioB')= 'AJO') then
     fTRepartition.PutGridDetail(fGrille,False,False,'I_IMMO;I_LIBELLE;PLAFONDAJO;PLAFONDAJO;DEPRECIATION')
  else
     fTRepartition.PutGridDetail(fGrille,False,False,'I_IMMO;I_LIBELLE;PLAFONDAJO;PLAFONDREP;REPRISE');

  // MAJ Zone Plafond
  AffichPlafond;

  // Pas d'erreur sur les cellules saisies
  fCellErr :=False;
end;


procedure TOF_AMDEPRECGroupe.MontantOpeRestant;
var lEncours : integer;
    i : integer;
    TotalSaisi, Ecart : double;
begin
  TotalSaisi := 0;
  lEncours := fGrille.Row;

  // Total des montants saisis autre que la cellule courante
  for i:=0 to fTRepartition.Detail.Count - 1 do
     if (i <> lEncours-1) then
        TotalSaisi := Arrondi
         (TotalSaisi + Valeur(fGrille.Cells[DP_COL_DEPREC,i+1]),V_PGI.OkDecV );

  Ecart := Arrondi((Valeur (GetControlText('DP_MontantOpe')) - TotalSaisi), V_PGI.OkDecV);
  if Ecart > Valeur (fGrille.Cells[DP_COL_PLAFOND,lEncours]) then
     begin
     fGrille.Cells[DP_COL_DEPREC,lEncours] := StrFMontant(Ecart,15 , V_PGI.OkDecV, '' , true);
     HShowMessage (HMD[6],'','');
     Ecart := 0;
     end
  else if  (Ecart < 0) then
     begin
     fGrille.Cells[DP_COL_DEPREC,lEncours] := StrFMontant(Ecart,15 , V_PGI.OkDecV, '' , true);
     PGIInfo('Montant négatif !','Attention');
     Ecart := 0;
     end;
  fGrille.Cells[DP_COL_DEPREC,lEncours] := StrFMontant(Ecart,15 , V_PGI.OkDecV, '' , true);
  fGrille.Refresh;
end;


function TOF_AMDEPRECGroupe.ControleDateOpeOK : boolean;
var Q : TQuery;
    i : integer;
    stop : boolean;
begin
  Result := True;
  // La date est-elle valide ?
  if (not IsValidDate (GetControlText('DP_DateOpe'))) then
  begin
    HShowMessage (HMD[2],'','');
    // Attention le déplacement du focus provoque le ONEXIT de la zone quittée
    // => messages supplémentaires d'erreur
    // Ecran.FocusControl (THEdit(GetControl('DP_DateOpe')));
    Result := false;  Exit;
  end;
  // La date est-elle un date de l'exercice en cours ?
  if (StrToDate(GetControlText('DP_DateOpe')) < VHImmo^.Encours.Deb)
  or (StrToDate(GetControlText('DP_DateOpe')) > VHImmo^.Encours.Fin) then
  begin
    HShowMessage (HMD[3],'','');
    // Attention le déplacement du focus provoque le ONEXIT de la zone quittée
    // => messages supplémentaires d'erreur
    // Ecran.FocusControl (THEdit(GetControl('DP_DateOpe')));
    Result := false;  Exit;
  end;
  // La date est-elle postérieure à la date de dernière opération ?
  stop := False;
  for i:=0 to fTRepartition.Detail.Count - 1 do
      begin
      Q := OpenSQL ('SELECT IL_DATEOP FROM IMMOLOG WHERE IL_IMMO="'
                     +fTRepartition.Detail[i].GetValue('I_IMMO')+'" ORDER BY IL_IMMO,Il_DATEOP DESC', True);
      if not Q.Eof then
         if Q.FindField ('IL_DATEOP').AsDateTime > StrToDate(GetControlText('DP_DateOpe')) then
            begin
            Ferme(Q);
            stop:= True;
            Break;
            end;
      Ferme (Q);
      end;
  if stop then
     begin
     HShowMessage (HMD[4],'','');
     // Attention le déplacement du focus provoque le ONEXIT de la zone quittée
     // => messages supplémentaires d'erreur
     // Ecran.FocusControl (THEdit(GetControl('DP_DateOpe')));
     Result := false; Exit;
     end;
end;


function TOF_AMDEPRECGroupe.ControleMontantOpeOK : boolean;
begin
   result := True;
   if Valeur (GetControlText('DP_MontantOpe')) <= 0 then
         begin
         PGIInfo('Montant négatif ou nul !','Attention');
         // Attention le déplacement du focus provoque le ONEXIT de la zone quittée
         // => messages supplémentaires d'erreur
         // Ecran.FocusControl (THEdit(GetControl('DP_MontantOpe')));
         Result := False; Exit;
         end;
   if Valeur (GetControlText('DP_MontantOpe')) > Valeur (GetControlText('DP_Plafond')) then
         begin
         HShowMessage (HMD[5],'','');
         // Attention le déplacement du focus provoque le ONEXIT de la zone quittée
         // => messages supplémentaires d'erreur
         // Ecran.FocusControl (THEdit(GetControl('DP_MontantOpe')));
         Result := False; Exit;
         end;
   fMontantOpe := Valeur (GetControlText('DP_MontantOpe'));
   // Recalcul des répartitions si Prorata retenu
   if GetControlText('DP_Repartition')<>'ZLI' then RechargerGrille;
end;


function TOF_AMDEPRECGroupe.ControleZonesOK : boolean;
var Verif: boolean;
    Sender: Tobject;
    NumColEnCours,NumRowEnCours : integer;
begin
   fDateErr := True;
   fMontantErr := True;
   fCellErr := True;
   if ControleDateOpeOK then
      begin
      fDateErr := False;
      if ControleMontantOpeOK then
         begin
         fMontantErr := False;
         // Vérif cellule sous le curseur en saisie libre
         Sender:=nil;
         NumColEnCours := FGrille.Col;
         NumRowEnCours := FGrille.Row;
         OnGrilleCellExit (Sender,NumColEnCours,NumRowEnCours,Verif);
         if not fCellErr then fCellErr := not ControleSaisieOK;
         end;
      end;
    result := ((fDateErr = False) AND (fMontantErr = False) AND (fCellErr = False));
end;



function TOF_AMDEPRECGroupe.ControleSaisieOK : boolean;
var i : integer;
    Cumul : double;

begin
  // Contrôler que chaque montant est < plafond de l'immo même en répartition automatique
  // car dernière cellule calculée en automatique = montant opération - somme des autres cellules
  //stop := 0;
  Result := True;
  for i:=0 to fTRepartition.Detail.Count - 1 do
      if  Valeur(fGrille.Cells[DP_COL_DEPREC,i+1]) > Valeur(fGrille.Cells[DP_COL_PLAFOND,i+1]) then
          begin
          //stop := i+1;
          Result := False;   break;
          end;
  if not Result then
     begin
     SetActiveTabSheet('DP_PAGELISTE');
     HShowMessage (HMD[8],'','');
     Ecran.FocusControl (THValComboBox(GetControl('DP_Repartition')));
     exit;
     end;

  // Contrôler que le total de la colonne Dépréciation ne dépasse pas le montant de l'opération
  Cumul := 0;
  for i:=0 to fTRepartition.Detail.Count - 1 do
      Cumul := Arrondi (Cumul + Valeur(fGrille.Cells[DP_COL_DEPREC,i+1]), V_PGI.OkDecV);
  Result := (Cumul = fMontantOpe);
  if not Result then
     begin
     SetActiveTabSheet('DP_PAGELISTE');
     HShowMessage (HMD[7],'','');
     end;
end;


procedure TOF_AMDEPRECGROUPE.EnregistrerDepreciationImmo ;
var i, Reponse: integer;

begin
  Reponse := mrYes;

  // Dernier message de confirmation si aucune dépréciation dans les exos précédents
  if fConfirme then  Reponse := HShowMessage (HMD[9],'','');

  if (Reponse = mrYes) then
    begin
    // Enregistrement de la colonne saisie dans les TOB colonne DEPRECIATION
    fTRepartition.GetGridDetail(fGrille,fGrille.RowCount-1,'','I_IMMO;I_LIBELLE;PLAFONDAJO;PLAFONDAJO;DEPRECIATION',True);

    // Calcul du numéro d'ordre pour opération en série
    fOrdreSerie := TrouveNumeroOrdreSerieLogSuivant;

    // Transaction pour ne pas enregister une table sans l'autre
    // BTY 01/06 FQ 17171 Message d'erreur en cas de pb de transaction
    BEGINTRANS ;
    try
      begin
      // Par immo, enregistrement de l'opération dans la base
      for i:=0 to fTRepartition.Detail.Count - 1 do
          EnregUneImmo (i, fTRepartition.Detail[i].GetValue('I_IMMO'));
      // Ecriture base
      COMMITTRANS ;
      // 05/06 FQ 18119 Positionner les indicateurs de modif
      VHImmo^.ChargeOBImmo := True;
      ImMarquerPublifi (True);
      end;
    except
      HShowMessage (HMD[10],'','');
      ROLLBACK ;
    end;

    // BTY 12/05 FQ 17183
    TFVIERGE(ECRAN).retour := IntToStr(mrYes);
    end;
end;

procedure TOF_AMDEPRECGROUPE.EnregUneImmo (rang : integer; stCode : string) ;
var EnregImmo : TEnregDeprecGroupe;
    TImmo : TOB;
begin

  TImmo := TOB.Create ('IMMO', nil, -1);
  try
    if TImmo.SelectDB('"'+stCode+'"',nil,True) then
    begin
      EnregImmo             :=TEnregDeprecGroupe.Create;
      try
        // Remplissage de l'objet EnregImmo
        EnregImmo.Code      :=TImmo.GetValue('I_IMMO');
        //EnregImmo.JournalA  := '***';
        EnregImmo.Operation :=  'X';
        // BTY 01/06 Nelle zone de IMMO
        EnregImmo.OpeDeprec :=  'X';
        EnregImmo.TypeOpe   := 'DPR';
        EnregImmo.DateOpe   := StrToDate(GetControlText('DP_DateOpe'));
        EnregImmo.TypeDot   := 'DPR';
        EnregImmo.Ordre     := TrouveNumeroOrdreLogSuivant(TImmo.GetValue('I_IMMO'));
        EnregImmo.OrdreSerie := fOrdreSerie;
        if (GetControlText('DP_RadioB')= 'AJO') then
            begin
            EnregImmo.LibelleLog  := 'Dépréciation';
            EnregImmo.MontantDot   := fTRepartition.Detail[rang].GetValue('DEPRECIATION');
            end
        else
            begin
            EnregImmo.LibelleLog  := 'Reprise de dépréciation';
            EnregImmo.MontantDot   := (-1) * fTRepartition.Detail[rang].GetValue('DEPRECIATION');
            end;
        //EnregImmo.CodeMutation   := TImmo.GetValue('I_JOURNALA');
        EnregImmo.BlocNote       := RichToString (DP_Blocnote); // Attention compatibilité CWAS DP_BLOCNOTE.LinesRTF;

        // Transaction pour ne pas enregister une table sans l'autre
        //BEGINTRANS ;
        //try
          // MAJ IMMO et son nouveau plan d'amortisseemnt
          MAJUneImmo (EnregImmo) ;
          // Ajout d'un enreg dans IMMOLOG
          AjoutUneImmoLog (EnregImmo);
          // Ecriture base
          //COMMITTRANS ;
          //except
          //ROLLBACK ;
          //end;

      finally
        EnregImmo.free ;
      end;
    end;
  finally
    TImmo.Free;
  end;
end;

// MAJ une immo
procedure TOF_AMDEPRECGROUPE.MAJUneImmo (EnregImmo : TEnregDeprecGroupe) ;
var Q : TQuery;
   Plan : TPlanAmort;
begin
  Q := OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+ EnregImmo.Code +'"', FALSE);
  Q.Edit ;
  //Q.FindField('I_JOURNALA').AsString  := EnregImmo.JournalA;
  Q.FindField('I_OPERATION').AsString  := EnregImmo.Operation;
  // BTY 01/06 Nelle zone de IMMO
  Q.FindField('I_OPEDEPREC').AsString  := EnregImmo.OpeDeprec;
  Q.FindField('I_OPECHANGEPLAN').AsString:= EnregImmo.Operation;
  Q.FindField('I_REVISIONECO').AsFloat := EnregImmo.MontantDot;

  // mbo - fq 17476
  Q.FindField('I_REVISIONFISCALE').AsFloat := EnregImmo.MontantDot;

  Q.Post;
  Q.Edit ;
  // Calcul nouveau plan d'amortissement
  Plan:=TPlanAmort.Create(true) ;
  try
      Plan.Charge(Q);
      Plan.Recupere(EnregImmo.Code,Q.FindField('I_PLANACTIF').AsString);
      EnregImmo.PlanActifAv := Plan.NumSeq;
      Plan.SetTypeOpe(EnregImmo.TypeOpe);
      Plan.Calcul(Q, idate1900);
      Plan.Sauve;
      // Attribuer à l'immo le n° du nouveau plan d'amrtissement
      Q.FindField('I_PLANACTIF').AsInteger := Plan.NumSeq;
      // Date de la dernière dotation
      Q.FindField('I_DATEDERMVTECO').AsDateTime := Plan.GetDateFinAmortEx(Plan.AmortEco);
      Q.FindField('I_DATEDERNMVTFISC').AsDateTime := Plan.GetDateFinAmortEx(Plan.AmortFisc);
      Q.Post;
      EnregImmo.PlanActifAp := Plan.NumSeq;
  finally
    Plan.Free;
  end;
  Ferme(Q);
end;

// Ajout d'un enreg dans IMMOLOG
procedure TOF_AMDEPRECGROUPE.AjoutUneImmoLog (EnregImmo : TEnregDeprecGroupe) ;
//var Q : TQuery;
var TLog : TOB;

begin
{*  Q:=OpenSQL('SELECT * FROM IMMOLOG WHERE IL_IMMO="'+stCode+'"', FALSE) ;
  Q.Insert ;
  // InitNew (Q);  // sert à initialiser les champs ?
  Q.FindField('IL_IMMO').AsString:= EnregImmo.Code ;
  Q.FindField('IL_LIBELLE').AsString:= EnregImmo.LibelleLog ;
  Q.FindField('IL_DATEOP').AsDateTime:= EnregImmo.DateOpe;
  Q.FindField('IL_TYPEOP').AsString:= EnregImmo.TypeOpe;
  Q.FindField('IL_ORDRE').AsInteger:= EnregImmo.Ordre;
  Q.FindField('IL_ORDRESERIE').AsInteger:= EnregImmo.OrdreSerie;
  Q.FindField('IL_TYPEDOT').AsString := EnregImmo.TypeDot;
  Q.FindField('IL_MONTANTDOT').AsFloat := EnregImmo.MontantDot;
  Q.FindField('IL_CODEMUTATION').AsString := EnregImmo.CodeMutation;
  //{$IFDEF EAGLCLIENT
  Q.PutValue('IL_BLOCNOTE',DP_BlocNote.LinesRTF.Text) ;
  //{$ELSE
  TBlobField(Q.FindField('IL_BLOCNOTE')).Assign(DP_BlocNote.LinesRTF) ;
  //{$ENDIF
  // ATTENTION Compléter les champs vides
  Q.FindField('IL_Cptemutation').AsString:= '' ;
  Q.FindField('IL_TypeModif').AsString:= '' ;
  Q.FindField('IL_MotifCes').AsString:= '' ;
  Q.FindField('IL_MontantCes').AsFloat:= 0 ;
  Q.Post;
  Ferme(Q);
*}

  // TOB plus avantageuse que Query car évite de compléter les champs vides
  TLog := TOB.Create ('IMMOLOG',nil,-1);
  try
    TLog.PutValue('IL_IMMO',EnregImmo.Code);
    TLog.PutValue('IL_LIBELLE', EnregImmo.LibelleLog);
    TLog.PutValue('IL_TYPEMODIF',AffecteCommentaireOperation(EnregImmo.TypeOpe));
    TLog.PutValue('IL_DATEOP', EnregImmo.DateOpe);
    TLog.PutValue('IL_TYPEOP', EnregImmo.TypeOpe);
    TLog.PutValue('IL_ORDRE', EnregImmo.Ordre);
    TLog.PutValue('IL_OrdreSerie', EnregImmo.OrdreSerie);
    TLog.PutValue('IL_TYPEDOT', EnregImmo.TypeDot);
    TLog.PutValue('IL_MONTANTDOT', EnregImmo.MontantDot);
    //TLog.PutValue('IL_CODEMUTATION', EnregImmo.CodeMutation);
    TLog.PutValue('IL_PLANACTIFAV', EnregImmo.PlanActifAv);
    TLog.PutValue('IL_PLANACTIFAP', EnregImmo.PlanActifAp);
    // TLog.PutValue('IL_BLOCNOTE', DP_BlocNote.LinesRTF.Text);
    // Attention compatibilité CWAS
    // TLog.PutValue('IL_BLOCNOTE', RichToString (DP_Blocnote));
    TLog.PutValue('IL_BLOCNOTE', EnregImmo.BlocNote);

    TLog.InsertDB(nil);
  finally
    TLog.Free;
  end;

end;



procedure TOF_AMDEPRECGroupe.OnEnterDateOpe(Sender: TObject);
begin
  // BTY 06/06 Message bloquant si grille vide :
  // Date opération = 1er champ qui a le focus à l'ouverture de la fenêtre
  // Si grille vide, message d'erreur et sortir ici car
  // sortir dans le OnArgument est impossible (exécute systématiquement OnLoad)
  // sortir dans le OnLoad est trop tôt (le message apparaît et la fenêtre pas encore,
  // et la fenêtre s'ouvre de toute façon en fin de OnLoad et PostMessage la ferme après,
  // => clignotement de la fenêtre furtive
  if fVide then
     begin
     PGIInfo ('Aucun élément à déprécier dans l''ensemble ' + fRegroupement +
              ' - ' + RechDom('AMREGROUPEMENT',fRegroupement,False), 'Dépréciation');
     PostMessage (Ecran.Handle, WM_CLOSE, 0, 0);   // passe par OnClose
     end;
end;


Initialization
  registerclasses ( [ TOF_AMDEPRECGROUPE ] ) ;
end.





