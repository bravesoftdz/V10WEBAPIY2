{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 07/09/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AMPRIME ()
Mots clefs ... : TOF;AMPRIME
*****************************************************************}
Unit AMPRIME_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Forms,
     Windows,
     Menus,
     Graphics,
     uTob,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
     HDB,
{$else}
     eMul,
     MaineAgl,
{$ENDIF}
     iment,
     ImOuPlan,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     implan,
     outils,
     HRichOle,
     ParamSoc,
     ImPlanMeth,
     HTB97,
     AGLInit,
     Vierge,
     UTOF ;

Type
  TOF_AMPRIME = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;

  private
    { Variables }
    fCode_immo : String;
    fPlan:TPlanAmort;
    fPlanActifAv: integer;
    fPlanActifAp: integer;
    fBLOCNOTE: THRichEditOLE;
    fDateDeb: TDateTime;
    okFiscal : boolean;
    NAM : Boolean;
    fMethode:TMethodeAmort;
    MethodePrime: string;
    DureePrime : integer;
    BasePrime : double;
    ModifDuree : boolean;
    SavPrime : double;
    SavDuree : integer;
    Sbv : double;
    MSG1 : string;

    Procedure OnbFermeClick(Sender: TObject);
    procedure ANTKeyDown(Sender: TObject; var Key: Word;
                                 Shift: TShiftState);
    Procedure CalculSuramortissement;
    Function  ControleDateOpe  :Boolean;
    Function  ControleDuree :Boolean;
    Function  ControleAnterieur :Boolean;
    Function  ControleBase : Boolean;
    Procedure OnExitMntPrime(Sender: TObject);
    Procedure OnExitDuree(Sender: TObject);

  end ;


Function AMLanceFiche_AMPRIME(stCode: string):Integer;

Implementation
Const
  HM: array[0..14] of string =
  ('Cette immobilisation a déjà fait l''objet d''une prime d''équipement.',
   'La saisie de la date est obligatoire.',
   'Le montant de la prime doit être positif et non nul.',
   'Opération impossible car l''immobilisation est totalement amortie sur le plan fiscal.',
   'Opération impossible car l''immobilisation est totalement amortie sur le plan économique.',
   'La date de l''opération doit étre supérieure ou égale à la plus haute date d''opération déjà enregistrée.',
   'La date de l''opération doit appartenir à l''exercice en cours.',
   'Voulez-vous enregistrer cette opération ?',
   'Le montant du suramortissement déjà pratiqué doit être positif et ne peut étre supérieur à la moitié de la prime.',
   'L''enregistrement de la prime d''équipement a échoué.',
   'Durée incorrecte : la date de fin d''amortissement est antérieure à la date de début d''exercice en cours.',
   'Opération impossible car ' +
        'la date de début d''amortissement prise en compte pour la prime est postérieure à l''exercice en cours.',
   'Opération impossible car ' +
        'le bien a déjà fait l''objet d''une subvention d''investissement égale à la base d''amortissement.',
   'Le montant maximum de la prime doit être inférieur ou égal à la base d''amortissement ' +
        'diminuée du montant de la subvention d''investissement déjà octroyée soit ',
   'Le montant de la prime doit être inférieur ou égal à la base d''amortissement.');

Function AMLanceFiche_AMPRIME(Stcode: string):Integer;
var
   Qplan:tQuery;
   fPlan:TplanAmort;
   NAM:boolean;
   fDateDeb:TDateTime;
   Sbv : double;
   BaseControle : double;

begin
   result := 0;

   // Ctl présence immo avec une prime déjà saisie
   if ExisteSQL ('SELECT IL_IMMO FROM IMMOLOG WHERE IL_IMMO="'+
    Stcode+'" AND IL_TYPEOP="PRI"') then
      PGIBox(HM[0])
   else
   begin
       NAM := false;

       fPlan:=TPlanAmort.Create(true) ;
       QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+Stcode+'"', FALSE) ;
       fPlan.Charge(QPlan);
       fPlan.Recupere(stCode,QPlan.FindField('I_PLANACTIF').AsString);

       Sbv := QPlan.FindField('I_SBVMT').AsFloat;

       if fPlan.Fiscal = true then
       begin
          fDateDeb := fPlan.DateDebFis;
          BaseControle := fPlan.AmortFisc.Base;
       end else
       begin
          if fPlan.AmortEco.Methode = 'NAM' then
          begin
             BaseControle := fPlan.ValeurAchat + fPlan.ValeurTVArecuperable -
                             fPlan.ValeurTVArecuperee;
             fDateDeb := fPlan.DateMiseEnService;
             NAM := true;
          end else
          begin
             BaseControle := fPlan.AmortEco.Base;
             fDateDeb := fPlan.DateDebEco;
          end;
       end;

    // si immo déjà subventionnée et subvnetion = base d'amortissement ---> saisie prime non autorisée
     if Sbv = BaseControle then
     begin
        PGIBox(HM[12]);
        fPlan.Free ;
        Ferme(QPlan);
     end else
     begin
       // si la date début amort > exercice en cours ---> saisie non autorisée
       if fDateDeb > VHImmo^.Encours.Fin then
       begin
           PGIBox(HM[11]);
           fPlan.Free ;
           Ferme(QPlan);
       end else
       begin
          // si l'immo est totalement amortie ---> saisie prime non autorisée
          if (not NAM) then
          begin
             if (fPlan.Fiscal) and (VHImmo^.Encours.deb > fPlan.AmortFisc.DateFinAmort) then
             begin
               PGIBox(HM[3]);
               fPlan.Free ;
               Ferme(QPlan);
             end else
             begin
               if (not (fPlan.Fiscal)) and (VHImmo^.Encours.deb > fPlan.AmortEco.DateFinAmort) then
               begin
                   PGIBox(HM[4]);
                   fPlan.Free ;
                   Ferme(QPlan);
               end else
               begin
                   fPlan.Free ;
                   Ferme(QPlan);
                   result:=StrToInt(AGLLanceFiche('AM', 'AMPRIME','' , '', Stcode));
               end;
             end;
          end else
          begin
             fPlan.Free ;
             Ferme(QPlan);
             result:=StrToInt(AGLLanceFiche('AM', 'AMPRIME','' , '', Stcode));
          end;
        end;
     end;
   end;
end;

procedure TOF_AMPRIME.OnArgument (S : String ) ;
var QPlan : TQuery;
  fMethodeRetenue : string;
begin
  Inherited ;
    fMethodeRetenue := ''; 
    // Initialisation
    TToolbarButton97(GetControl('BFerme')).Onclick := OnbFermeClick;
    TFVIERGE(ECRAN).retour := IntToStr(MrNo);
    //TFVIERGE(ECRAN).OnAfterFormShow := OnAfterFormShowEcran;

    THEdit(GetControl('PE_ANT')).OnKeyDown := ANTkeyDown;
    ThEdit(GetControl ('PE_PRIME')).OnExit := OnExitMntPrime;
    ThEdit(GetControl ('PE_DUREE')).OnExit := OnExitDuree;
    // Assignation du bloc note
    fBLOCNOTE := THRichEditOLE(GetControl('PE_BLOCNOTE'));
    ModifDuree := false;
    // Lecture de l'immo en cours
    fPlan:=TPlanAmort.Create(true) ;
    fCode_immo := S;
    QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+fCode_immo+'"', FALSE) ;
    fPlan.Charge(QPlan);
    fPlan.Recupere(fCode_immo,QPlan.FindField('I_PLANACTIF').AsString);

    SetControlProperty('PE_VALEURHT', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
    SetControlProperty('PE_PRIME', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
    SetControlProperty('PE_ANT', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));

    SetControlCaption('PE_CODE',fPlan.Codeimmo);
    SetControlCaption('PE_DESI',fPlan.libelleImmo);
    SetControlCaption('PE_ACHAT', DateToStr(fPlan.dateachat));
    SetControlCaption('PE_VALEURHT',StrFMontant(fPlan.valeurht,15,V_PGI.OkDecV,'',True));

    NAM := false;
    Sbv := QPlan.FindField('I_SBVMT').AsFloat;
    DureePrime := 0;

    if fPlan.Fiscal = true then
    begin
        fMethodeRetenue := 'FISC';
        okFiscal := true;
        fDateDeb := fPlan.DateDebFis;
        fMethode := fPlan.AmortFisc;
        MethodePrime := fPlan.AmortFisc.Methode;
        DureePrime:= fPlan.AmortFisc.Duree;
        BasePrime := fPlan.AmortFisc.Base;

    end else
    begin
        okFiscal := false;
        if fPlan.AmortEco.Methode = 'NAM' then
        begin
           fMethodeRetenue := 'NAM';
           fDateDeb := fPlan.DateMiseEnService;
           NAM := true;
           MethodePrime := 'LIN';
           BasePrime := fPlan.ValeurAchat + fPlan.ValeurTVArecuperable -
                        fPlan.ValeurTVArecuperee;
           // si saisie d'une subvention, la zone i_corvrcedde est alimentée
           if QPlan.FindField('I_SBVMT').AsFloat <> 0 then
           begin
              SetControlProperty('PE_DUREE', 'MinValue', QPlan.FindField('I_CORVRCEDDE').AsFloat);
              SetControlProperty('PE_DUREE', 'MaxValue', QPlan.FindField('I_CORVRCEDDE').AsFloat);
              SetControlText('PE_DUREE',FloatToStr(QPlan.FindField('I_CORVRCEDDE').AsFloat));
              DureePrime := StrToInt(GetControlText('PE_DUREE'));
              SetControlEnabled('PE_DUREE', false);
              ModifDuree := true;
           end else
           begin
              SetControlText ('PE_DUREE', '120');
           end;
        end else
        begin
           fMethodeRetenue := 'ECO';
           fmethode := fPlan.AmortEco;
           fDateDeb := fPlan.DateDebEco;
           MethodePrime := fPlan.AmortEco.Methode;
           DureePrime:= fPlan.AmortEco.Duree;
           BasePrime:= fplan.AmortEco.Base;
        end;
   end;

   if (not NAM) then
   begin
      SetControlVisible('PE_DUREE', false);
      SetControlVisible('PE_LIMOIS', false);
      SetControlVisible('PE_LIDUREE', false);

   end else
   begin
      SetControlVisible('PE_DUREE', true);
      SetControlVisible('PE_LIMOIS', true);
      SetControlVisible('PE_LIDUREE', true);

   end;

  // pas de suramortissement déjà pratiqué si c'est une acquisition
  if fDateDeb >= VHImmo^.Encours.deb then
  begin
     SetControlVisible('PE_ANT', false);
     SetControlVisible('PE_LIANT', false);
  end;

  SetControlCaption('PE_DEBAMORT', DateToStr(fDateDeb));

  IF NAM then
  begin
     SetControlCaption('PE_DATEDEB', 'Date de mise en service');
     SetControlCaption('PE_BASELIB', 'Base');
  end else if fMethodeRetenue = 'FISC' then
  begin
     SetControlCaption('PE_DATEDEB', 'Date début d''amortissement fiscal');
     SetControlCaption('PE_BASELIB', 'Base d''amortissement fiscal');
  end else
  begin
     SetControlCaption('PE_DATEDEB', 'Date début d''amortissement économique');
     SetControlCaption('PE_BASELIB', 'Base d''amortissement économique');
  end;

  SetControlCaption('PE_BASEAMORT', StrFMontant(BasePrime,15,V_PGI.OkDecV,'',True));

  BasePrime := BasePrime - Sbv;

end;


{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 31/08/2006
Modifié le ... :   /  /
Description .. : Touche F6 sur la zone suramortissement =
Suite ........ : récup du montant déjà amorti calculé
Mots clefs ... :
*****************************************************************}
procedure TOF_AMPRIME.ANTKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  inherited;
  if Key=VK_F6 then CalculSuramortissement;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 08/09/2006
Modifié le ... :   /  /
Description .. : en sortie de saisie du montant de la prime, si le controle du
Suite ........ : montant de la prime est ok, on calcule les antérieurs
Mots clefs ... :
*****************************************************************}
procedure TOF_AMPRIME.OnExitMntPrime(Sender:TObject);

begin
 if (savPrime <> Valeur(GetControlText('PE_PRIME'))) or (savPrime = 0) then
 begin

   // le montant de la prime doit être inférieur ou égal à la base
   if (fDateDeb < VHImmo^.Encours.Deb) then
   begin
       if (not NAM) then
       begin
         CalculSurAmortissement;
         savPrime := Valeur(GetControlText('PE_PRIME'));
       end else
       begin
          CalculSurAmortissement;
          savPrime := Valeur(GetControlText('PE_PRIME'));
       end;
   end;
 end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 08/09/2006
Modifié le ... :   /  /
Description .. : en sortie de saisie de la durée on calcule les antérieurs si
Suite ........ : nécessaire
Mots clefs ... :
*****************************************************************}
procedure TOF_AMPRIME.OnExitDuree(Sender:TObject);
begin

 if (SavDuree <> StrToInt(GetControlText('PE_DUREE'))) then
 begin
   ModifDuree := true;

   if (fDateDeb < VHImmo^.Encours.Deb) then
   begin
     CalculSurAmortissement;
     savduree := StrToInt(GetControlText('PE_DUREE'));
   end;
 end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 31/08/2006
Modifié le ... :   /  /
Description .. : module calculant les antérieurs déjà amortis compte tenu
Suite ........ : du montant de la prime et des datedeb, mode, durÚe
Suite ........ : adéquats
Mots clefs ... :
*****************************************************************}
Procedure TOF_AMPRIME.CalculSuramortissement;
var
TmpPlan: TPlanAmort;
Q : TQuery;
PresenceSBV : boolean;
begin
   // le principe = renseigner la méthode AmortPri avec les éléments nécessaires au calcul
    TmpPlan := TPlanAmort.Create(true);
    try
      Q:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+fCode_immo+'"', FALSE) ;
      TmpPlan.Charge(TQuery(Q));

      PresenceSBV := (Q.FindField('I_SBVMT').AsFloat <> 0);

      TmpPlan.AmortPri.base := (Valeur(GetControlText('PE_PRIME')))/2;
      tmpPlan.PRIDateDeb := fDateDeb;
      tmpPlan.AmortPRI.Creation := true;

      if PresenceSBV then tmpPlan.AmortSBV.Creation := true;

      if TmpPlan.AmortEco.Methode <> 'NAM' then
      begin
         tmpplan.AmortPri.duree:=fMethode.duree;
         tmpplan.AmortPri.taux := fMethode.taux;
      end else
      begin
         tmpplan.AmortPri.duree:=StrToInt(GetControlText ('PE_DUREE'));
         tmpplan.AmortPri.taux := GetTaux(MethodePrime,
                                          fDateDeb, fDateDeb,StrToInt(GetControlText ('PE_DUREE')));
      end;
      TmpPlan.CalculDateFinAmortissement(tmpplan.AmortPri);
      TmpPlan.CalculReprises;
      SetControlText('PE_ANT',StrFMontant(TmpPlan.AmortPri.Reprise,15,V_PGI.OkDecV,'',True));
    finally
      TmpPlan.free;
    end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 31/08/2006
Modifié le ... : 08/09/2006
Description .. : le montant de la prime ne doit pas être supérieure à la base - montantSubvention:
Suite ........ : en NAM,  base = valeur achat + TVA recupérée - TVA
Suite ........ : récupérable
Suite ........ : en Eco, base = baseeco
Suite ........ : en Fiscal, base = basefisc
Mots clefs ... :
*****************************************************************}
Function TOF_AMPRIME.ControleBase: boolean;

begin
  result := False;

  if Valeur(GetcontrolText('PE_PRIME'))<= 0 then
  begin
      PGIBox(HM[2]);
      result := True;
      SetFocusControl('PE_PRIME');
  end else
  begin
  if (Sbv <> 0) and (Valeur(GetcontrolText('PE_PRIME')) > BasePrime) then
    begin
      msg1 := StrFMontant(BasePrime,15,V_PGI.OkDecV,'',True);
      PGIBox(HM[13]+ msg1);
      result := True;
      SetFocusControl('PE_PRIME');
    end else
    begin
      if (Valeur(GetcontrolText('PE_PRIME')) > BasePrime) then
      begin
        PGIBox(HM[14]);
        result := True;
        SetFocusControl('PE_PRIME');
      end;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 31/08/2006
Modifié le ... :   /  /
Description .. : la date de l'opération doit être comprise dans l'exercice
Suite ........ : Elle doit être supérieure à la dernière opération enregistrée
Mots clefs ... :
*****************************************************************}

Function TOF_AMPRIME.ControleDateOpe: boolean;

var date_saisie: TdateTime;
begin
  result := False;
if TRIM(GetControlText('PE_DATOP')) = '/  /' then
begin
  PGIBox(HM[1]);
  result := true;
  SetFocusControl('PE_DATOP');

end else
begin

  date_saisie := StrToDate(GetControlText('PE_DATOP'));

  // la date doit appartenir à l'exercice en cours
  if (date_saisie >  VHImmo^.Encours.Fin) OR ( date_saisie < VHImmo^.Encours.Deb) then
  begin
      PGIBox(HM[6]);
      result := True;
      SetControlText('PE_DATOP', '');
      SetFocusControl('PE_DATOP');
  end
  // date de l'opération doit être > date dernière opération
  else if ExisteSQL('SELECT IL_DATEOP FROM IMMOLOG WHERE IL_IMMO="'+fCode_immo+'" AND IL_DATEOP>"'+USDATETIME(date_saisie)+'"') then
     begin
        PGIBox(HM[5]);
        result := True;
        SetFocusControl('PE_DATOP');
     end;
end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 31/08/2006
Modifié le ... :   /  /
Description .. : le montant du suramortissement pratiqué ne peut être
Suite ........ : supérieur à la prime
Mots clefs ... :
*****************************************************************}
Function TOF_AMPRIME.ControleAnterieur: boolean;

begin
  result := False;

  // la date doit appartenir à l'exercice en cours
  if (fDateDeb < VHImmo^.Encours.Deb) then
  begin
      if (Valeur(GetControlText('PE_ANT'))< 0) or (Valeur(GetControlText('PE_ANT'))> Valeur(GetControlText('PE_PRIME'))) then
      begin
         PGIBox(HM[8]);
         result := True;
         //SetControlText('PE_ANT', '0.00');
         SetFocusControl('PE_ANT');
      end;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 31/08/2006
Modifié le ... :   /  /
Description .. : Pour les immos non amortissables : controle que
Suite ........ : date de mise en service + durée d'inaliénabilité > date deb en cours
Mots clefs ... :
*****************************************************************}
Function TOF_AMPRIME.ControleDuree: boolean;

var
TmpPlan: TPlanAmort;
Q : TQuery;

begin
  result := false;
  // le principe = renseigner la méthode AmortPri avec les éléments nécessaires au calcul de la date de fin
  // en fonction de la durée
  TmpPlan := TPlanAmort.Create(true);
  try
     Q:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+fCode_immo+'"', FALSE) ;
     TmpPlan.Charge(TQuery(Q));

     TmpPlan.AmortPri.base := (Valeur(GetControlText('PE_PRIME')))/2;
     TmpPlan.PRIDateDeb := fDateDeb;

     if (NAM) then
     begin
        TmpPlan.AmortPri.duree:=StrToInt(GetControlText ('PE_DUREE'));
        TmpPlan.AmortPri.taux := GetTaux(MethodePrime, fDateDeb, fDateDeb,
                                       StrToInt(GetControlText ('PE_DUREE')));

     end else
     begin
        TmpPlan.AmortPri.Duree:= DureePrime;
        tmpPlan.AmortPri.taux := GetTaux(MethodePrime, fDateDeb, fDateDeb,
                                         DureePrime);
     end;

     TmpPlan.CalculDateFinAmortissement(tmpplan.AmortPri);
     if TmpPlan.AmortPri.DateFinAmort < VHImmo^.Encours.Deb then
     begin
        PGIBox(HM[10]);
        result := True;
        SetControlText('PE_DUREE', '');
        SetFocusControl('PE_DUREE');
     end else
     begin
        if (fDateDeb < VHImmo^.Encours.Deb) and (SavDuree <> StrToInt(GetControlText('PE_DUREE'))) then
           CalculSurAmortissement;
     end;
  finally
    TmpPlan.free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 31/08/2006
Modifié le ... :   /  /
Description .. : Sortie par la croix ou annuler
Mots clefs ... :
*****************************************************************}
procedure TOF_AMPRIME.OnbFermeCLick(Sender: TObject);
Var
    question : TmodalResult;
begin
  Inherited ;

  // Action Echap ou Fermer par la croix
  Question := PGIAskCancel(HM[7],Ecran.Caption);
  If Question = mrYes then
     // controle validation
     OnUpdate
  else If Question = mrNo then
     // Sortir
     Ecran.ModalResult := mrNo
  else
     // Cancel
     Ecran.ModalResult := mrNone;


end ;

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 31/08/2006
Modifié le ... :   /  /
Description .. : création d'une nouvelle méthode = AmortPrime
Suite ........ : qui reprend :
Suite ........ : 1) les éléments du fiscal (durée - méthode - date début
Suite ........ :    d'amortissement) si présence d'un fiscal
Suite ........ :
Suite ........ : 2) si pas de fiscal et éco <> NAM, reprend les élements de
Suite ........ :     l'éco
Suite ........ :
Suite ........ : 3) si pas de fiscal et éco = NAM, reprend durée
Suite ........ :     d'inaliénabilité, date de mise en service et linéaire
Mots clefs ... :
*****************************************************************}
procedure TOF_AMPRIME.OnUpdate ;
var Ordre : integer;
    TLog : TOB;
    lib :String;
    erreur : Boolean;
    QPlan : TQuery;
    PlanNew : TPlanAmort;

begin
  Inherited ;
  ecran.ModalResult := mrYes;

   // Controle de la date de l'opération
   erreur := ControleDateOpe;
   IF erreur=True then
   Begin
       ecran.ModalResult :=mrNone;
       SetFocusControl('PE_DATOP');
       exit;
   end;

   erreur := ControleBase;
   IF erreur=True then
   Begin
       ecran.ModalResult :=mrNone;
       SetFocusControl('PE_PRIME');
       exit;
   end;

   // controle de la durée d'inaliénabilité pour les NAM
   if (NAM) then
   begin
     erreur := ControleDuree;
     IF erreur=True then
     Begin
        ecran.ModalResult :=mrNone;
        SetFocusControl('PE_DUREE');
        exit;
     end;
   end;

   // controle du suramortissement déjà pratiqué
   erreur := ControleAnterieur;
   if erreur=True then
   Begin
        ecran.ModalResult := mrNone;
        SetFocusControl('PE_ANT');
        exit;
   end;

   // Mise à jour des dates les tables IMMO,IMMOAMOR et IMMOLOG
   BEGINTRANS ;
   try

      // Mise à jour de IMMO
      QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+fCode_immo+'"', FALSE) ;

      // Maj de la fiche immo
      QPlan.Edit;
      QPlan.FindField('I_OPERATION').AsString:= 'X' ;
      QPlan.FindField('I_SBVPRI').AsFloat:= Valeur(GetControlText('PE_PRIME'));
      if (NAM) and (QPlan.FindField('I_CORVRCEDDE').AsFloat = 0) then
      begin
         //Duree := Valeur(GetControlText('PE_DUREE'));
         //StrDuree := IntToStr(Duree);
         QPlan.FIndField('I_CORVRCEDDE').AsFloat:= StrToFloat(GetControlText('PE_DUREE'));
      end;
      QPlan.FIndField('I_REPRISEUO').AsFloat:= Valeur(GetControlText('PE_ANT'));
      QPlan.Post;

      // Maj du plan d'amortissement IMMOAMOR
      QPlan.Edit;
      PlanNew:=TPlanAmort.Create(true) ;
      try
         PlanNew.Charge(QPlan);
         PlanNew.Recupere(fCode_immo,QPlan.FindField('I_PLANACTIF').AsString);
         fPlanActifAv := PlanNew.NumSeq;
         PlanNew.SetTypeOpe('PRI');
         PlanNew.Calcul(Qplan, idate1900);
         PlanNew.Sauve;

         // Attribuer à l'immo le n° du nouveau plan d'amortissement
         Qplan.FindField('I_PLANACTIF').AsInteger := PlanNew.NumSeq;

         Qplan.Post;
         fPlanActifAp := PlanNew.NumSeq;

      finally
         PlanNew.Free;
         Ferme(Qplan);
      end;


      // Mise à jour de IMMOLOG
      // ATTENTION : stockage du montant de la prime dans IL_MONTANTEXC

      lib:= 'Saisie d''une prime d''équipement';
      Ordre := TrouveNumeroOrdreLogSuivant(fPlan.CodeImmo);

      TLog := TOB.Create ('IMMOLOG',nil,-1);
      try
         TLog.PutValue('IL_IMMO',fCode_immo);
         TLog.PutValue('IL_LIBELLE',lib);
         TLog.PutValue('IL_TYPEMODIF',AffecteCommentaireOperation('PRI'));

         TLog.PutValue('IL_DATEOP', StrToDate(GetControlText('PE_DATOP')));
         TLog.PutValue('IL_METHODEECO', MethodePrime);

         // on stocke la durée dans la zone il_dureeco (même si c'est le plan fiscal qui est pris en cpte)
         if (NAM) then
              TLog.PutValue('IL_DUREEECO', StrToInt(GetControlText('PE_DUREE')))
         else
              TLog.PutValue('IL_DUREEECO', DureePrime);

         TLog.PutValue('IL_MONTANTEXC', GetControlText('PE_PRIME'));

         TLog.PutValue('IL_TYPEOP', 'PRI');
         TLog.PutValue('IL_ORDRE', Ordre);
         TLog.PutValue('IL_OrdreSerie', -1);
         TLog.PutValue('IL_PLANACTIFAV', fPlanActifAv);
         TLog.PutValue('IL_PLANACTIFAP', fPlanActifAp);
         TLog.PutValue('IL_BLOCNOTE', RichToString (fBlocnote));
         TLog.InsertDB(nil);
      finally
         TLog.Free;
      end;

      COMMITTRANS ;

      // 05/06 FQ 18119 Positionner les indicateurs de modif
      VHImmo^.ChargeOBImmo := True;
      ImMarquerPublifi (True);

   except
      PGIBox(HM[9]);
      ROLLBACK ;
   end ;

   TFVIERGE(ECRAN).retour := IntToStr(MrYes);
end ;

procedure TOF_AMPRIME.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AMPRIME.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AMPRIME.OnLoad ;
begin
  Inherited ;
end ;


procedure TOF_AMPRIME.OnDisplay () ;
begin
  Inherited ;
end ;



Initialization
  registerclasses ( [ TOF_AMPRIME ] ) ;
end.


