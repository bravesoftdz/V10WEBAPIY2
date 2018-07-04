{***********UNITE*************************************************
Auteur  ...... :  Bernadette Tynévez
Créé le ...... : 18/07/2006
Modifié le ... : 17/10/2006
Description .. : Source TOF de la FICHE : AMDPIANTERIEURS ()
Suite ........ : 10/06 FQ 18949 F6 dans une cellule d'un exo = MIN (solde DPI dispo, Base éco)
Suite ........ : 10/06 FQ 18954 Contrôle DPI < base éco
Mots clefs ... : TOF;AMDPIANTERIEURS
*****************************************************************}
Unit AMDPIANTERIEURS_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     windows,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     MaineAGL,
{$ENDIF}
     AmDPI,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     uTob,
     UTOF,
     HTB97,
     AGLInit, // TheData
     Vierge;

Type
  TOF_AMDPIANTERIEURS = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
     { Variables }
     fGrille: THGrid;
     fNbLig: integer;
     fMontantDPI : double;
     fDateDebEco : TDateTime;
     fReintegration : double;
     fBaseEco : double;
     fBaseTheorique : double;
     fCas : integer;    // Cas de Base fiscalement déductible
     fChampOK : boolean;
     AmDPI : TAmDPI;    // Récup des libellés des exos
     fMontantOK : boolean;
     {Procedures}
     procedure OnbFermeClick(Sender: TObject);
     procedure FormaterfGrille;
     procedure InitColonneAnnee;
     procedure OnExitMontantDPI(Sender: TObject);
     procedure OnExitANTDPI(Sender: TObject);
     procedure OnGrilleCellExit(Sender: TObject; var ACol, ARow: Integer;
               var Cancel: Boolean);
     function  ControleMontantDPIOK : boolean;
     function  ControleAntDPIOK : boolean;
     function  ControleCellOK (ACol, ARow: Integer; Zone : string) : boolean;
     function  ControleZonesOK : boolean;
     function  ControleBaseEcoOK : boolean;
     procedure CalculerRepartition(Sender: TObject);
     function  TotalRepartition(Sender: TObject) : double;
     function  TotalDPIUtilisable(Sender: TObject) : double;
     procedure StockerTheDataARetourner;
     procedure OnGrilleKeyDown(Sender:TObject; var Key:Word; Shift:TShiftState);
     procedure OnDPIKeyDown(Sender:TObject; var Key:Word; Shift:TShiftState);
  end ;

Function AMLanceFiche_DPIAnterieurs(stParam: string) : integer;

Implementation

const
 COL_ANNEE = 0;
 COL_MONT  = 1;

 HMD: array[0..10] of string =
 {0} ( 'Voulez-vous enregistrer cette opération ?',
 {1}   'Le montant doit être supérieur à zéro.',
 {2}   'Le montant affecté dépasse le solde disponible pour l''exercice traité, soit ',
 {3}   'La base fiscalement déductible tient-elle compte déjà de la DPI ?',
 {4}   'La base d''amortissement économique tient-elle compte déjà de la DPI ?',
 {5}   'L''immobilisation doit être totalement déductible.',
 {6}   'Veuillez déduire la DPI du montant de la base d''amortissement économique.',
 {7}   'Situation incorrecte.',
 {8}   'Le montant total affecté aux exercices + antérieurs doit être égal à la DPI.',
 {9}   'Le montant doit être supérieur ou égal à zéro.',
{10}   'Le montant de la DPI ne doit pas dépasser la base d''amortissement économique, soit '
     );

Function AMLanceFiche_DPIAnterieurs(stParam: string) : integer;
begin
  result:= StrToInt(AGLLanceFiche('AM', 'AMDPIANTERIEURS','' ,'',stParam) );
end;


procedure TOF_AMDPIANTERIEURS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AMDPIANTERIEURS.OnUpdate ;
begin
  Inherited;
  // Ne pas sortir a priori
  Ecran.ModalResult := mrNone;
  if not fMontantOK then exit;

  if ControleZonesOK then
     if ControleBaseEcoOK then
        begin
        Ecran.ModalResult := mrYes;
        TFVIERGE(ECRAN).retour := IntToStr(MrYes);
        // Stocker les infos à retourner dans la variable globale TheData
        StockerTheDataARetourner;
        end;
end ;

procedure TOF_AMDPIANTERIEURS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AMDPIANTERIEURS.OnArgument (S : String ) ;
var mht,tvaRable,tvaRee : double;
begin
  Inherited ;

  // Récup paramètres en entrée
  fDateDebEco := StrToDate(ReadTokenSt(S));
  fReintegration := StrToFloat(ReadTokenSt(S));
  fBaseEco := StrToFloat(ReadTokenSt(S));
  mht := StrToFloat(ReadTokenSt(S));
  tvaRable := StrToFloat(ReadTokenSt(S));
  tvaRee := StrToFloat(ReadTokenSt(S));
  fBaseTheorique := Arrondi (mht+tvaRable-tvaRee, V_PGI.OkDecV);

  // AmDPI variable locale pour connaître les libellés des exos
  AmDPI := TAmDPI.Create;
  AmDPI.ChargeDateDPI;

  // Initialisations
  TFVIERGE(ECRAN).retour := IntToStr(MrNo);
  fMontantDPI := 0;
  fChampOK := True;

  // Initialisation des champs
  SetControlEnabled('RESTE', False);
  SetControlProperty('DPI', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('ANT', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  fGrille := THGrid ( GetControl ('DPIListe') );

  // Evènements
  fGrille.OnCellExit  := OnGrilleCellExit;
  TToolbarButton97(GetControl('BFerme')).Onclick := OnbFermeClick;
  THEdit(GetControl('DPI')).OnExit    := OnExitMontantDPI;
  THEdit(GetControl('ANT')).OnExit    := OnExitANTDPI;
  fGrille.OnKeyDown := OnGrilleKeyDown;
  THEdit(GetControl ('DPI')).OnKeyDown := OnDPIKeyDown;


  // Initialisation du tableau
  FormaterfGrille;
  InitColonneAnnee;

  SetFocusControl('DPI');
end ;

procedure TOF_AMDPIANTERIEURS.FormaterfGrille;
begin
  fGrille.ColEditables[COL_ANNEE]:=False;
//  fGrille.ColWidths[COL_ANNEE]:=228;
  fGrille.ColAligns[COL_ANNEE]:=taCenter ; //tJustify;

  fGrille.ColEditables[COL_MONT]:= True;
  fGrille.ColTypes[COL_MONT]:='R';
  fGrille.ColFormats[COL_MONT]:= StrfMask(V_PGI.OkDecV, '', True);
//  fGrille.ColWidths[COL_MONT]:=118;
  fGrille.ColAligns[COL_MONT]:=taRightJustify;
end;

procedure TOF_AMDPIANTERIEURS.InitColonneAnnee;
var i, nolig, iprem : integer;
    prem : boolean;
    stDPIDispo : string;
begin
 nolig := 0;
 iprem := 0;
 prem := False;

 if (TheData <> nil) then
    if (TheData is TAmDPI) then
    // TheData variable globale pointe depuis l'appelant sur l'objet fAmDPI alimenté

     for i:= 2 to 5 do
     begin
       // Exemple i = 3 : Si Date ECO > Fin N-3 <=> Date ECO dans N-2 => afficher exo N-3
       if fDateDebEco > TAmDPI(TheData).Exercice[i].Fin then
          begin
          Inc(nolig);
          if AmDPI.Exercice[i].Code <> '' then
            fGrille.CellValues[COL_ANNEE, nolig] := AmDPI.Exercice[i].EtatCpta // contient le libellé
          else
            fGrille.CellValues[COL_ANNEE, nolig] := 'Du ' + DateToStr(TAmDPI(TheData).Exercice[i].Deb) +
                                                   ' au ' + DateToStr(TAmDPI(TheData).Exercice[i].Fin);
          if not prem then
             begin
             prem := True;
             iprem := i; // indice 1er exo affiché
             end;
          end;
     end;

 // Nb de lignes utiles
 fNbLig := nolig;
 if nolig = 0 then fGrille.Rowcount := 2
 else fGrille.Rowcount := nolig+1;

 // Bloquer en saisie si aucun exercice affichable
 if fNbLig = 0 then
    fGrille.ColEditables[COL_MONT]:= False;

  // Affichage DPIs disponibles sur l'ensemble des exos affichés
  if fNbLig = 0 then
     SetControlVisible ('StDPIDispo', False)
  else
     begin
     StDPIDispo := 'DPI utilisables du '   +
                   DateToStr(TAmDPI(TheData).Exercice[5].Deb)  + ' au ' +
                   DateToStr(TAmDPI(TheData).Exercice[iprem].Fin)  + '    ';
     SetControlText ('StDPIDispo', StDPIDispo +
                  StrFMontant(TotalDPIUtilisable(nil),15,V_PGI.OkDecV,'',true));
     end;
end;


procedure TOF_AMDPIANTERIEURS.OnGrilleCellExit(Sender: TObject; var ACol, ARow: Integer;
          var Cancel: Boolean);
begin
  inherited;

  if fNbLig <> 0 then
  begin
     case ACol of
     COL_MONT :
       begin
       if ControleCellOK (ACol, ARow, fGrille.Cells[ACol,ARow]) then
         // Formater la saisie
         fGrille.Cells[ACol,ARow] := StrfMontant
                   (Valeur(FGrille.Cells[ACol,ARow]), 15, V_PGI.OkDecV, '' , true)
       else
         // Formater à zéro
         fGrille.Cells[ACol,ARow] := StrfMontant(0, 15, V_PGI.OkDecV, '' , true);
       end;
     end;

  end else fChampOK := True;

  // Recalcul Reste à affecter
  SetControlText('RESTE',
      StrFMontant(Valeur(GetControlText('DPI'))-Valeur(GetControlText('ANT'))-TotalRepartition(nil),15,V_PGI.OkDecV,'',true));
end;

function TOF_AMDPIANTERIEURS.ControleCellOK (ACol, ARow: Integer; Zone : string) : boolean;
var RdMontant, Solde : double;
begin
  result := True;
  if IsNumeric(fGrille.Cells[ACol,ARow]) then
  begin
    RdMontant := Valeur(Zone);
    Solde := 0;

    // TheData variable globale pointe depuis l'appelant sur l'objet fAmDPI alimenté
    if (TheData <> nil) then
      if (TheData is TAmDPI) then

        if RdMontant < 0 then
          begin
          result := False;
          PGIBox(HMD[1], '');
          end
        else
          begin
          // Dernière ligne <=> exo N-5
          if ARow = fNbLig   then Solde := TAmDPI(TheData).Solde[5];
          // Avant-dernière ligne <=> exo N-4
          if ARow = fNbLig-1 then Solde := TAmDPI(TheData).Solde[4];
          if ARow = fNbLig-2 then Solde := TAmDPI(TheData).Solde[3];
          if ARow = fNbLig-3 then Solde := TAmDPI(TheData).Solde[2];
          // Comparer au solde disponible de l'exo de la ligne courante
          if RdMontant > Solde then
            begin
            result := False;
            PGIBox(HMD[2]+ StrfMontant (Solde, 15, V_PGI.OkDecV, '', true) + '.' , '');
            end;
          end;

  end else
    result := False;

  fChampOK := result;
end;

function TOF_AMDPIANTERIEURS.ControleAntDPIOK : boolean;
begin
   result := (Valeur (GetControlText('ANT')) >= 0);
   if not result then
      begin
      PGIBox(HMD[9], '');
      SetFocusControl('ANT');
      end
   // Comparer DPI avec répartitions+antérieurs
   else if Arrondi(Valeur(GetControlText('DPI'))-Valeur(GetControlText('ANT'))-TotalRepartition(nil), V_PGI.OkDecV) <> 0 then
      begin
      result := False;
      PGIBox(HMD[8], '');
      SetFocusControl('DPI');
      end;
  // fChampOK := result;
end;

function TOF_AMDPIANTERIEURS.ControleMontantDPIOK : boolean;
begin
   result := (Valeur (GetControlText('DPI')) > 0);
   if not result then
      begin
      PGIBox(HMD[1], '');
      SetFocusControl('DPI');
      exit;
      end;

   // FQ 18954
   if Arrondi(Valeur(GetControlText('DPI'))-fBaseEco, V_PGI.OkDecV) > 0 then
      begin
      result := False;
      PGIBox(HMD[10] + StrfMontant (fBaseEco,15,V_PGI.OkDecV,'',true) + '.' , '');
      SetFocusControl('DPI');
      exit;
      end;

   // Recalculer les affectations théoriques
   if fMontantDPI <> Valeur (GetControlText('DPI')) then
      CalculerRepartition(nil);
   fMontantDPI := Valeur (GetControlText('DPI'));
end;

function TOF_AMDPIANTERIEURS.ControleZonesOK : boolean;
var Verif: boolean;
    ColEnCours, RowEnCours : integer;
begin
  fMontantOK := ControleMontantDPIOK;
  if fMontantOK then
     begin
     // Vérif cellule sous le curseur, positionne fChampOK
     ColEnCours := FGrille.Col;
     RowEnCours := FGrille.Row;
     OnGrilleCellExit (nil, ColEnCours, RowEnCours, Verif);
     if fChampOK then
        fChampOK := ControleAntDPIOK;
     end;
  result := (fMontantOK and fChampOK);
end;

procedure TOF_AMDPIANTERIEURS.OnExitMontantDPI(Sender: TObject);
begin
   fMontantOK := ControleMontantDPIOK;

   {  if ControleMontantDPIOK then
        begin
        // Recalculer les affectations
        CalculerRepartition(nil);
        fMontantDPI := Valeur (GetControlText('DPI'));
        end; }
end;

procedure TOF_AMDPIANTERIEURS.OnExitAntDPI(Sender: TObject);
begin
  SetControlText('RESTE',
       StrFMontant((Valeur(GetControlText('DPI'))-Valeur(GetControlText('ANT'))-TotalRepartition(nil)),15,V_PGI.OkDecV,'',true));
end;

function TOF_AMDPIANTERIEURS.ControleBaseEcoOK : boolean;
var RepBaseEco, RepBaseFD : TmodalResult;
begin
  result := True;
  fCas := 0;

  if fReintegration = 0 then
  // fiscalement déductible = 0 ------------------------------------------------

     if fBaseEco = fBaseTheorique then
        if Arrondi(fMontantDPI-fBaseEco, V_PGI.OkDecV) > 0 then
          begin
          result := False;
          PGIBox(HMD[10] + StrfMontant (fBaseEco,15,V_PGI.OkDecV,'',true) + '.' , '');
          end
        else
          fCas := 1

     else
        // Base ECO contient-elle la DPI
        case PGIAsk(HMD[4], Ecran.Caption) of
        mrYes: begin
               result := False;
               PGIBox(HMD[6], '');
               end;
        mrNo:  if Arrondi(fMontantDPI-fBaseEco, V_PGI.OkDecV) > 0 then
                  begin
                  result := False;
                  PGIBox(HMD[10] + StrfMontant (fBaseEco,15,V_PGI.OkDecV,'',true) + '.', '');
                  end
               else fCas := 3;
        end
  else
  // fiscalement déductible <> 0 -----------------------------------------------

     if fBaseEco = fBaseTheorique then
        // Base fiscalement déductible contient-elle la DPI
        case PGIAsk(HMD[3], Ecran.Caption) of
        mrYes: if Arrondi(fMontantDPI+fReintegration, V_PGI.OkDecV) = fBaseEco then
                  fCas := 2
               else
                  begin
                  result := False;
                  PGIBox(HMD[5], '');
                  end;
        mrNo:  begin
               result := False;
               PGIBox(HMD[5], '');
               end;
        end

    else
        begin
        // Base ECO contient-elle la DPI
        RepBaseEco := PGIAsk(HMD[4], Ecran.Caption);
        // Base fiscalement déductible contient-elle la DPI
        RepBaseFD := PGIAsk(HMD[3], Ecran.Caption);

        case RepBaseEco of
        mrYes: case RepBaseFD of
               mrYes: begin
                      result := False;
                      PGIBox(HMD[7], '');
                      end;
               mrNo:  begin
                      result := False;
                      PGIBox(HMD[5], '');
                      end;
               end;
        mrNo:  case RepBaseFD of
               mrYes: if Arrondi (fReintegration+fMontantDPI, V_PGI.OkDecV) = fBaseEco then
                         fCas := 2
                      else
                         begin
                         result := False;
                         PGIBox(HMD[5], '');
                         end;
               mrNo:  begin
                         result := False;
                         PGIBox(HMD[5], '');
                      end;
               end;
        end;
        end;
end;


procedure TOF_AMDPIANTERIEURS.CalculerRepartition(Sender: TObject);
var dpi, ant, prec, restant :double ;
    i , isolde :integer;
begin
   dpi := Valeur (GetControlText('DPI'));
   ant := Valeur (GetControlText('ANT'));
   prec := 0;
   isolde := 5;

   // Affecter la DPI en commençant par l'exo le plus ancien
   for i:= fNbLig downto 1 do
   begin
     restant := Arrondi (dpi-ant-prec, V_PGI.OkDecV);

     if restant < 0 then
        fGrille.Cells[COL_MONT, i] := StrfMontanT (0, 15, V_PGI.OkDecV, '' , true)
     else
        // Si le reste à répartir sur l'exo i dépasse le solde disponible de l'exo :
        // mettre le solde et réserver le restant pour l'exo i-1 et ainsi de suite
        // jusqu'au premier exo
        begin
        if restant > TAmDPI(TheData).Solde[iSolde] then
           fGrille.Cells[COL_MONT, i] := StrfMontanT (TAmDPI(TheData).Solde[iSolde], 15, V_PGI.OkDecV, '' , true)
        else
           fGrille.Cells[COL_MONT, i] := StrfMontanT (restant, 15, V_PGI.OkDecV, '' , true);
        end;

     iSolde := iSolde - 1;
     prec := Arrondi(prec + Valeur(fGrille.Cells[COL_MONT, i]),V_PGI.OkDecV);
   end;

   // Reste à affecter
   //SetControlText('RESTE', StrFMontant((dpi-prec),15 , V_PGI.OkDecV, '' , true));
   SetControlText('RESTE', StrFMontant((dpi-ant-TotalRepartition(nil)),15 , V_PGI.OkDecV, '' , true));
end;


function TOF_AMDPIANTERIEURS.TotalRepartition(Sender: TObject) : double;
var i :integer;
begin
   // Totaliser les répartitions des exercices
   result := 0;
   for i:= 1 to fNbLig do
       result := Arrondi(result + Valeur(FGrille.Cells[COL_MONT,i]), V_PGI.OkDecV);
end;


function TOF_AMDPIANTERIEURS.TotalDPIUtilisable(Sender: TObject) : double;
var i :integer;
begin
   // Totaliser les DPI disponibles de N-5 à l'exercice précédant la date de début d'amortissement éco
   result := 0;

   if (TheData <> nil) then
      if (TheData is TAmDPI) then
        // TheData variable globale pointe depuis l'appelant sur l'objet fAmDPI alimenté
        begin
        for i:= 2 to 5 do
          // Exemple i = 3 : Si Date ECO > Fin N-3 <=> Date ECO dans N-2 => on traite exo N-3
          if fDateDebEco > TAmDPI(TheData).Exercice[i].Fin then
             break;
        // Total dispo sur les exercices antérieurs à celui passé en paramètre
        TAmDPI(TheData).ChargeSoldeDisponibleDate (TAmDPI(TheData).Exercice[i-1].Fin);
          result := TAmDPI(TheData).SoldeDispoDate;
        end;
end;


// Action Echap ou Fermer par la croix
procedure TOF_AMDPIANTERIEURS.OnbFermeClick(Sender: TObject);
begin

  case PGIAskCancel(HMD[0], Ecran.Caption) of
    mrYes:
      begin
      if ControleZonesOK then
         if ControleBaseEcoOK then
         begin
         Ecran.ModalResult := mrYes;
         TFVIERGE(ECRAN).retour := IntToStr(MrYes);
         // OK stocker les infos à retourner dans TheData (variable globale)
         StockerTheDataARetourner;
         end  else Ecran.ModalResult := mrNone
      else Ecran.ModalResult := mrNone;
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

procedure TOF_AMDPIANTERIEURS.StockerTheDataARetourner;
var i , jSolde: integer;
begin
  // Je stocke dans la variable globale TheData qui pointe sur fAmDPI depuis l'appelant
  // Stocker le tableau en commençant par l'exo le plus ancien <=> N-5
  if TheData <> nil then
     if (TheData is TAmDPI) then
       begin
       // RAZ avant d'alimenter
       TAmDPI(TheData).Montant :=  0;
       TAmDPI(TheData).St := '0';
       for i:= 0 to 5 do
          TAmDPI(TheData).Solde [i]:= 0;

       jSolde := 5;
       for i:= fNbLig downto 1 do
         begin
         TAmDPI(TheData).Solde [jSolde]:= Valeur (fGrille.Cells[COL_MONT,i]);
         jSolde := jSolde-1;
         end;
       // Indice 0 exo en cours est libre => montant antérieurs
       //TAmDPI(TheData).Solde [0]:= Valeur (GetControlText('ANT'));
       TAmDPI(TheData).Anterieur := Valeur (GetControlText('ANT'));
       // Montant DPI
       TAmDPI(TheData).Montant :=  Valeur (GetControlText('DPI'));
       // Cas Base fiscalement déductible
       TAmDPI(TheData).St := IntToStr (fCas);
       end;
end;


procedure TOF_AMDPIANTERIEURS.OnClose ;
begin
  AmDPI.Free;
  Inherited ;
end ;

procedure TOF_AMDPIANTERIEURS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AMDPIANTERIEURS.OnCancel () ;
begin
  Inherited ;
end ;

// Touche F6 pour récupérer la DPI utilisable de l'exercice de la ligne en cours
procedure TOF_AMDPIANTERIEURS.OnGrilleKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var Solde : double;
begin
  case Key of
    VK_F6 : begin
            if TheData <> nil then
               if (TheData is TAmDPI) then
                  begin
                  Solde := 0;
                  // Dernière ligne <=> exo N-5
                  if fGrille.Row = fNbLig   then Solde := TAmDPI(TheData).Solde[5];
                  // Avant-dernière ligne <=> exo N-4
                  if fGrille.Row = fNbLig-1 then Solde := TAmDPI(TheData).Solde[4];
                  if fGrille.Row = fNbLig-2 then Solde := TAmDPI(TheData).Solde[3];
                  if fGrille.Row = fNbLig-3 then Solde := TAmDPI(TheData).Solde[2];
                  // FQ 18949 Proposer MIN (solde DPI disponible, Base éco)
                  //fGrille.Cells[COL_MONT, fGrille.Row] := StrFMontant(Solde,15 , V_PGI.OkDecV, '' , true);
                  if Arrondi(Solde-fBaseEco, V_PGI.OkDecV) > 0 then
                     fGrille.Cells[COL_MONT, fGrille.Row] := StrFMontant(fBaseEco,15 , V_PGI.OkDecV, '' , true)
                  else
                     fGrille.Cells[COL_MONT, fGrille.Row] := StrFMontant(Solde,15 , V_PGI.OkDecV, '' , true);

                  fGrille.Refresh;
                  end;
            end;
  end;
end;


// Touche F6 pour récupérer le total DPI utilisables sur tous les exercices affichés
procedure TOF_AMDPIANTERIEURS.OnDPIKeyDown (Sender:TObject; var Key:Word; Shift:TShiftState);
begin
  case Key of
    VK_F6 :
      begin
      // FQ 18949 Proposer MIN (solde DPI disponible, Base éco)
      //SetControlText ('DPI', StrFMontant(TotalDPIUtilisable(nil),15,V_PGI.OkDecV,'',true));
      if Arrondi(TotalDPIUtilisable(nil)-fBaseEco, V_PGI.OkDecV) > 0 then
         SetControlText ('DPI', StrFMontant(fBaseEco,15,V_PGI.OkDecV,'',true))
      else
         SetControlText ('DPI', StrFMontant(TotalDPIUtilisable(nil),15,V_PGI.OkDecV,'',true));
      end;
  end;
end;


Initialization
  registerclasses ( [ TOF_AMDPIANTERIEURS ] ) ;
end.
