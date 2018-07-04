{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 27/09/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AMSUBVENTION ()
Suite ........ : FQ 19694 - mbo - 18.06.2007 - pouvoir sortir de la création de comptes
Suite ........ : mbo - 10/07/2007 - si un seul cpte éligible en sbv et pas de création d'autres cptes
                                    on ne pouvait plus valider la grille
Mots clefs ... : TOF;AMSUBVENTION
*****************************************************************}
Unit AMSUBVENTION_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Forms,
     uTob,
     Windows,
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
     imOutGen,
     contcpte,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
{$IFDEF SERIE1}
     S1util,
{$ELSE}
     ent1,
{$ENDIF}
     HMsgBox,
     implan,
     outils,
     HRichOle,
     ParamSoc,
     ImPlanMeth,
     HTB97,
     AGLInit,
     Vierge,
     LookUp,
     {$IFDEF MODENT1}
     CPTypeCons,
     {$ENDIF MODENT1}
     UTOF ;

Type
  TOF_AMSUBVENTION = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

private
    { Variables }
    fCode_immo : String;
    fPlan:TPlanAmort;
    fPlanActifAv: integer;
    fPlanActifAp: integer;
    fBLOCNOTE: THRichEditOLE;
    fDateDeb: TDateTime;    // date début amort éco ou fiscal ou date mise en service suivant le cas
    NAM : Boolean;          // true si non amortissable
    fMethodeRetenue : string;  // prise en compte de la méthode éco ou de la fiscale
    fMethode:TMethodeAmort;
    MethodeSBV: string;     // mode d'amortissement de la subvention
    DureeSbv : integer;     // durée éco ou fiscale ou inaliénabilité suivant le cas
    BaseControle : double;  // base éco ou fiscale ou valeur d'achat suivant le cas
    ModifDuree : boolean;
    BorneMiniSub : string;
    BorneMaxiSub : string;
    BorneQP      : string;
    RadMaxiSub   : string;
    RadMiniSub   : string;
    BorneRP      : string;
    Msg1         : string;
    SavDuree     : integer;
    SavMnt       : double;
    PRIME        : double;

    Procedure OnbFermeClick(Sender: TObject);
    procedure ANTKeyDown(Sender: TObject; var Key: Word;
                                 Shift: TShiftState);
    Procedure CalculReprise;    //calcul de la reprise
    Function  ControleDateOpe  :Boolean;
    Function  ControleDuree :Boolean;
    Function  ControleAnterieur :Boolean;
    Function  ControleBase : Boolean;
    procedure AjouteUnCompteSBV(var L:TList; Compte:string);
    Function  ControleCompte : Boolean;
    Procedure OnExitMntSub(Sender: TObject);
    Procedure OnExitDuree(Sender: TObject);
    Function  ControleDateSub : boolean;
    Procedure OnExitCompteSub(Sender : TObject);
    Procedure OnCompteSUBelipsisClick(Sender: TObject);
    procedure BourreBorneCompte (var St : String ; Bourre : Char );
    function  CompteQuotePart(CompteSub : string): string;
    function  CompteReprise(CompteSub : string): string;
  end ;

Function AMLanceFiche_AMSUBVENTION(stCode: string):Integer;

Implementation

{$IFNDEF SERIE1}
uses
  CPGeneraux_TOM;
{$ENDIF}

Const
  HM: array[0..24] of string =
  ('Cette immobilisation a déjà fait l''objet d''une subvention d''équipement.',
   'La saisie de la date est obligatoire.',
   'Le montant de la subvention doit être positif et non nul.',
   'Opération impossible car l''immobilisation est totalement amortie sur le plan fiscal.',
   'Opération impossible car l''immobilisation est totalement amortie sur le plan économique.',
   'La date de l''opération doit étre supérieure ou égale à la plus haute date d''opération déjà enregistrée.',
   'La date de l''opération doit appartenir à l''exercice en cours.',
   'Voulez-vous enregistrer cette opération ?',
   'Le montant de la reprise de subvention doit être positif et ne peut étre supérieur au montant de la subvention.',
   'L''enregistrement de la subvention d''équipement a échoué.',
   'Durée incorrecte : la date de fin d''amortissement est antérieure à la date de début d''exercice en cours.',
   'Opération impossible car ' +
        'la date de début d''amortissement prise en compte pour la subvention est postérieure à l''exercice en cours.',
   'Le compte de subvention n''est pas renseigné correctement (radical 1311 à 1318 ou 138).',
   'Le compte de reprise de subvention est lié au compte de subvention et ne peut être modifié.',
   'La date d''octroi de la subvention doit être inférieure ou égale à la date de fin d''exercice en cours.',
   'La date d''octroi de la subvention doit être supérieure ou égale à la date de début d''amortissement économique.',
   'La date d''octroi de la subvention doit être supérieure ou égale à la date de début d''amortissement fiscal.',
   'La date d''octroi de la subvention doit être supérieure ou égale à la date de mise en service.',
   'La saisie de la date d''octroi de la subvention est obligatoire.',
   'La saisie du numéro de compte est obligatoire.',
   'Le compte de quote-part est lié au compte de subvention et ne peut être modifié.',
   'Ce compte n''existe. Veuillez le créer.',
   'Opération impossible car ' +
        'le bien a déjà fait l''objet d''une prime d''équipement égale à la base d''amortissement.',
   'Le montant maximum de la subvention doit être inférieur ou égal à la base d''amortissement ' +
        'diminuée du montant de la prime d''équipement déjà allouée soit ',
   'Le montant de la subvention doit être inférieur ou égal à la base d''amortissement.');


Function AMLanceFiche_AMSUBVENTION(Stcode: string):Integer;
var
   Qplan:tQuery;
   fPlan:TplanAmort;
   NAM:boolean;
   fDateDeb:TDateTime;
   Prime : double;
   BaseControle : double;
begin
   result := 0;

   // Ctl présence immo avec une subvention déjà saisie
   if ExisteSQL ('SELECT IL_IMMO FROM IMMOLOG WHERE IL_IMMO="'+
    Stcode+'" AND IL_TYPEOP="SUB"') then
      PGIBox(HM[0])
   else
   begin
       NAM := false;

       fPlan:=TPlanAmort.Create(true) ;
       QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+Stcode+'"', FALSE) ;
       fPlan.Charge(QPlan);
       fPlan.Recupere(stCode,QPlan.FindField('I_PLANACTIF').AsString);
       Prime := QPlan.FindField('I_SBVPRI').AsFloat;

       if fPlan.Fiscal = true then
       begin
          BaseControle := fPlan.AmortFisc.Base;
          fDateDeb := fPlan.DateDebFis;
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
             BaseControle:= fplan.AmortEco.Base;
             fDateDeb := fPlan.DateDebEco;
          end;
       end;

     // si immo déjà prime = base d'amortissement ---> saisie non autorisée
     if Prime = BaseControle then
     begin
        PGIBox(HM[22]);
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
         // si l'immo est totalement amortie ---> saisie subvention non autorisée
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
                  result:=StrToInt(AGLLanceFiche('AM', 'AMSUBVENTION','' , '', Stcode));
              end;
            end;
         end else
         begin
            fPlan.Free ;
            Ferme(QPlan);
            result:=StrToInt(AGLLanceFiche('AM', 'AMSUBVENTION','' , '', Stcode));
         end;
       end;
     end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 28/09/2006
Modifié le ... :   /  /
Description .. : Alimentation des champs de la fiche
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSUBVENTION.OnArgument (S : String ) ;
var QPlan : TQuery;
begin
  Inherited ;
    // Initialisation
    TToolbarButton97(GetControl('BFerme')).Onclick := OnbFermeClick;
    TFVIERGE(ECRAN).retour := IntToStr(MrNo);

    THEdit(GetControl('SB_REPRISE')).OnKeyDown := ANTkeyDown;
    ThEdit(GetControl ('SB_MNTSUB')).OnExit := OnExitMntSub;
    ThEdit(GetControl ('SB_DUREE')).OnExit := OnExitDuree;
    ThEdit(GetControl ('SB_CPTSUB')).OnExit := OnExitCompteSub;

    THEdit(GetControl('SB_CPTSUB')).OnElipsisClick := OnCompteSUBelipsisClick;

    SetControlProperty('SB_CPTSUB', 'MaxLength', VHImmo^.Cpta[fbGene].Lg) ;
    SetControlProperty('SB_REPSUB', 'MaxLength', VHImmo^.Cpta[fbGene].Lg) ;
    SetControlProperty('SB_QUOTEPART', 'MaxLength', VHImmo^.Cpta[fbGene].Lg) ;

    // Assignation du bloc note
    fBLOCNOTE := THRichEditOLE(GetControl('SB_BLOCNOTE'));
    ModifDuree := false;
    // Lecture de l'immo en cours
    fPlan:=TPlanAmort.Create(true) ;
    fCode_immo := S;
    QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+fCode_immo+'"', FALSE) ;
    fPlan.Charge(QPlan);
    fPlan.Recupere(fCode_immo,QPlan.FindField('I_PLANACTIF').AsString);

    SetControlProperty('SB_VALEURHT', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
    SetControlProperty('SB_MNTSUB', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
    SetControlProperty('SB_REPRISE', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));

    SetControlCaption('SB_CODE',fPlan.Codeimmo);
    SetControlCaption('SB_DESI',fPlan.libelleImmo);
    SetControlCaption('SB_ACHAT', DateToStr(fPlan.dateachat));
    SetControlCaption('SB_VALEURHT',StrFMontant(fPlan.valeurht,15,V_PGI.OkDecV,'',True));

    NAM := false;
    Prime := QPlan.FindField('I_SBVPRI').AsFloat;

    DureeSBV := 0;

    if fPlan.Fiscal = true then
    begin
        fDateDeb := fPlan.DateDebFis;
        fMethode := fPlan.AmortFisc;
        MethodeSbv := fPlan.AmortFisc.Methode;
        DureeSbv:= fPlan.AmortFisc.Duree;
        BaseControle := fPlan.AmortFisc.Base;
        fMethodeRetenue := 'FISC';
    end else
    begin
        if fPlan.AmortEco.Methode = 'NAM' then
        begin
           fDateDeb := fPlan.DateMiseEnService;
           NAM := true;
           fMethodeRetenue := 'NAM';
           MethodeSbv := 'LIN';
           BaseControle := fPlan.ValeurAchat + fPlan.ValeurTVArecuperable -
                        fPlan.ValeurTVArecuperee;
           // si saisie d'une prime avant, la zone i_corvrcedde est alimentée
           if Prime <> 0 then
           begin
              SetControlProperty('SB_DUREE', 'MinValue', QPlan.FindField('I_CORVRCEDDE').AsFloat);
              SetControlProperty('SB_DUREE', 'MaxValue', QPlan.FindField('I_CORVRCEDDE').AsFloat);
              SetControlText('SB_DUREE',FloatToStr(QPlan.FindField('I_CORVRCEDDE').AsFloat));
              DureeSBV := StrToInt(GetControlText('SB_DUREE'));
              SetControlEnabled('SB_DUREE', false);
              ModifDuree := true;
           end else
           begin
          SetControlText ('SB_DUREE', '120');
           end;
        end else
        begin
           fmethode := fPlan.AmortEco;
           fDateDeb := fPlan.DateDebEco;
           MethodeSbv := fPlan.AmortEco.Methode;
           DureeSbv:= fPlan.AmortEco.Duree;
           BaseControle:= fplan.AmortEco.Base;
           fMethodeRetenue := 'ECO';
        end;
   end;

   if (not NAM) then
   begin
      SetControlVisible('SB_DUREE', false);
      SetControlVisible('LI_MOIS', false);
      SetControlVisible('LI_DUREE', false);
   end else
   begin
       SetControlVisible('SB_DUREE', true);
       SetControlVisible('LI_MOIS', true);
       SetControlVisible('LI_DUREE', true);
   end;

  // pas de reprise si c'est une acquisition
  if fDateDeb >= VHImmo^.Encours.deb then
  begin
     SetControlVisible('SB_REPRISE', false);
     SetControlVisible('LI_REPRISE', false);

     SetControlCaption('SB_DATESUB', DateToStr(fDateDeb));

  end;

  SetControlCaption('SB_DEBAMORT', DateToStr(fDateDeb));

  IF fMethodeRetenue = 'NAM' then
  begin
     SetControlCaption('LI_DATEDEB', 'Date de mise en service');
     SetControlCaption('LI_BASEAMORT', 'Base');
  end else if fMethodeRetenue = 'FISC' then
  begin
     SetControlCaption('LI_DATEDEB', 'Date début d''amortissement fiscal');
     SetControlCaption('LI_BASEAMORT', 'Base d''amortissement fiscal');
  end else
  begin
     SetControlCaption('LI_DATEDEB', 'Date début d''amortissement économique');
     SetControlCaption('LI_BASEAMORT', 'Base d''amortissement économique');
  end;

  SetControlCaption('SB_BASE', StrFMontant(BaseControle,15,V_PGI.OkDecV,'',True));

  BaseControle := BaseControle - Prime;

  borneminiSub := '1311';
  BornemaxiSub := '1318';
  BourreBorneCompte(BorneMiniSub, '0' );
  BourreBorneCompte(BorneMaxiSub, '9' );
  RadMiniSub := '138';
  RadMaxiSub := '138';
  BourreBorneCompte(RadMiniSub, '0');
  BourreBorneCompte(RadMaxiSub, '9');

end ;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 27/09/2006
Modifié le ... :   /  /
Description .. : touche F6 = calcul des antérieurs théoriques
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSUBVENTION.ANTKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  inherited;
  if Key=VK_F6 then CalculReprise;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 28/09/2006
Modifié le ... : 28/09/2006
Description .. : en sortie du montant sub on calcule les antérieurs si besoin
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSUBVENTION.OnExitMntSub(Sender:TObject);

begin
  if (savMnt <> Valeur(GetControlText('SB_MNTSUB'))) or (savMnt = 0) then
  begin
    if (fDateDeb < VHImmo^.Encours.Deb) then
    begin
       CalculReprise;
       savMnt := Valeur(GetControlText('SB_MNTSUB'));
    end;
  end;
end;
{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 02/10/2006
Modifié le ... :   /  /
Description .. : en sortie de saisie de la durée on calcule les antérieurs si
Suite ........ : nécessaire
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSUBVENTION.OnExitDuree(Sender:TObject);
begin
 if (SavDuree <> StrToInt(GetControlText('SB_DUREE'))) then
 begin
   ModifDuree := true;
   IF (fDateDeb < VHImmo^.Encours.Deb) then
   begin
       CalculReprise;
       savduree := StrToInt(GetControlText('SB_DUREE'));
   end;
 end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 28/09/2006
Modifié le ... :   /  /
Description .. : Controle de la date d'opération :
Suite ........ : appartenir à l'exercice et supérieure ou égale à la dernière
Suite ........ : opération effectuée sur l'immo
Mots clefs ... :
*****************************************************************}
Function TOF_AMSUBVENTION.ControleDateOpe: boolean;

var date_saisie: TdateTime;
begin
result := False;
if TRIM(GetControlText('SB_DATEOP')) = '/  /' then
begin
  PGIBox(HM[1]);
  result := true;
  SetFocusControl('SB_DATEOP');
end else
begin
  date_saisie := StrToDate(GetControlText('SB_DATEOP'));

  // la date doit appartenir à l'exercice en cours
  if (date_saisie >  VHImmo^.Encours.Fin) OR ( date_saisie < VHImmo^.Encours.Deb) then
  begin
      PGIBox(HM[6]);
      result := True;
      SetFocusControl('SB_DATEOP');
  end
  // date de l'opération doit être > date dernière opération
  else if ExisteSQL('SELECT IL_DATEOP FROM IMMOLOG WHERE IL_IMMO="'+fCode_immo+'" AND IL_DATEOP>"'+USDATETIME(date_saisie)+'"') then
  begin
     PGIBox(HM[5]);
     result := True;
     SetFocusControl('SB_DATEOP');
  end;
end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 29/09/2006
Modifié le ... :   /  /
Description .. : Controle de la date d'octroi de la subvention
Suite ........ : Si plan fiscal : controle par rapport à datedebfis
Suite ........ : Si plan éco non NAM : controle par rapport à datedebeco
Suite ........ : Si plan éco NAM : contrôle par rapport à la date mise en
Suite ........ : service
Mots clefs ... :
*****************************************************************}
Function TOF_AMSUBVENTION.ControleDateSub: boolean;

var date_saisie: TdateTime;
begin
  result := False;
if TRIM(GetControlText('SB_DATESUB')) = '/  /' then
begin
  PGIBox(HM[1]);
  result := true;
  SetFocusControl('SB_DATESUB');

end else
begin
  date_saisie := StrToDate(GetControlText('SB_DATESUB'));

  // la date ne doit pas être au-delà de l'exercice en cours
  if (date_saisie >  VHImmo^.Encours.Fin) then
  begin
      PGIBox(HM[14]);
      result := True;
      SetFocusControl('SB_DATESUB');
  end else
  begin
      if (date_saisie < fDateDeb) then
      begin
        if fMethodeRetenue = 'ECO' then
        begin
          PGIBox(HM[15]);
          result := True;
          SetFocusControl('SB_DATESUB');
        end
        else if fMethodeRetenue = 'FISC' then
        begin
          PGIBox(HM[16]);
          result := True;
          SetFocusControl('SB_DATESUB');
        end else
        begin
          PGIBox(HM[17]);
          result := True;
          SetFocusControl('SB_DATESUB');
        end;
      end;
  end;
end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 28/09/2006
Modifié le ... :   /  /
Description .. : Le montant de la reprise ne peut être supérieur au montant
Suite ........ : de la subvention
Mots clefs ... :
*****************************************************************}
Function TOF_AMSUBVENTION.ControleAnterieur: boolean;
begin
  result := False;

  // la date doit appartenir à l'exercice en cours
  if (fDateDeb < VHImmo^.Encours.Deb) then
  begin
      if (Valeur(GetControlText('SB_REPRISE'))<0) or (Valeur(GetControlText('SB_REPRISE'))> Valeur(GetControlText('SB_MNTSUB'))) then
      begin
         PGIBox(HM[8]);
         result := True;
         SetFocusControl('SB_REPRISE');
      end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 28/09/2006
Modifié le ... :   /  /
Description .. : Pour les immos non amortissables : controle que
Suite ........ : date de mise en service + durée d'inaliénabilité > date deb en cours
Mots clefs ... :
*****************************************************************}
Function TOF_AMSUBVENTION.ControleDuree: boolean;

var
TmpPlan: TPlanAmort;
Q : TQuery;

begin
  result := false;

  // le principe = renseigner la méthode AmortSbv avec les éléments nécessaires au calcul de la date de fin
  // en fonction de la durée
  TmpPlan := TPlanAmort.Create(true);
  try
     Q:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+fCode_immo+'"', FALSE) ;
     TmpPlan.Charge(TQuery(Q));

     TmpPlan.AmortSbv.base := (Valeur(GetControlText('SB_MNTSUB')));
     TmpPlan.SBVDateDeb := fDateDeb;

     if (NAM) then
     begin
        TmpPlan.AmortSbv.duree:=StrToInt(GetControlText ('SB_DUREE'));
        TmpPlan.AmortSbv.taux := GetTaux(MethodeSBV, fDateDeb, fDateDeb,
                                       StrToInt(GetControlText ('SB_DUREE')));

     end else
     begin
        TmpPlan.AmortSbv.Duree:= DureeSbv;
        tmpPlan.AmortSbv.taux := GetTaux(MethodeSBV, fDateDeb, fDateDeb,
                                         DureeSbv);
     end;

     TmpPlan.CalculDateFinAmortissement(tmpplan.AmortSbv);
     if TmpPlan.AmortSbv.DateFinAmort < VHImmo^.Encours.Deb then
     begin
        PGIBox(HM[10]);
        result := True;
        SetFocusControl('SB_DUREE');
     end;
  finally
    TmpPlan.free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 02/10/2006
Modifié le ... :   /  /
Description .. : Controle que le montant de la subvention n'excède pas le
Suite ........ : montant de la base amortissable
Mots clefs ... :
*****************************************************************}
Function TOF_AMSUBVENTION.ControleBase: boolean;

begin
  result := False;

  if (Valeur(GetcontrolText('SB_MNTSUB'))<= 0.00) then
  begin
    PGIBox(HM[2]);
    result := True;
    SetFocusControl('SB_MNTSUB');
  end else
  begin
    if (Prime <> 0) and (Valeur(GetcontrolText('SB_MNTSUB')) > BaseControle) then
    begin
      msg1 := StrFMontant(BaseControle,15,V_PGI.OkDecV,'',True);
      PGIBox(HM[23]+ msg1);
      result := True;
      SetFocusControl('SB_MNTSUB');
    end else
    begin
      if (Valeur(GetcontrolText('SB_MNTSUB')) > BaseControle) then
      begin
        PGIBox(HM[24]);
        result := True;
        SetFocusControl('SB_MNTSUB');
      end;
    end;
  end;
end;
{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 02/10/2006
Modifié le ... :   /  /
Description .. : calcul des antérieurs théoriques compte tenu des différents
Suite ........ : paramètres.
Suite ........ : Cette fonction est également activée par la touche F6
Mots clefs ... :
*****************************************************************}
Procedure TOF_AMSUBVENTION.CalculReprise;
var
TmpPlan: TPlanAmort;
Q : TQuery;
PresencePRI : boolean;
begin
   // le principe = renseigner la méthode AmortSbv avec les éléments nécessaires au calcul
    TmpPlan := TPlanAmort.Create(true);
    try
      Q:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+fCode_immo+'"', FALSE) ;
      TmpPlan.Charge(TQuery(Q));

      PresencePRI := (Q.FindField('I_SBVPRI').AsFloat <> 0);

      TmpPlan.AmortSbv.base := (Valeur(GetControlText('SB_MNTSUB')));
      tmpPlan.SBVDateDeb := fDateDeb;

      if TmpPlan.AmortEco.Methode <> 'NAM' then
      begin
         tmpplan.AmortSbv.duree:=fMethode.duree;
         tmpplan.AmortSbv.taux := fMethode.taux;

         if fMethodeRetenue = 'ECO' then
            tmpplan.AmortSbv.PlanRef := 1
         else
            tmpplan.AmortSbv.PlanRef := 2;

      end else
      begin
         tmpplan.AmortSbv.PlanRef := 0;
         tmpplan.AmortSbv.duree:=StrToInt(GetControlText ('SB_DUREE'));
         tmpplan.AmortSbv.taux := GetTaux(MethodeSbv,
                                          fDateDeb, fDateDeb,StrToInt(GetControlText ('SB_DUREE')));
      end;

      if PresencePRI then TmpPlan.AmortPri.Creation := true;

      TmpPlan.AmortSbv.Creation := true;
      TmpPlan.CalculDateFinAmortissement(tmpplan.AmortSbv);
      TmpPlan.CalculReprises;
      SetControlText('SB_REPRISE',StrFMontant(TmpPlan.AmortSbv.Reprise,15,V_PGI.OkDecV,'',True));
    finally
      TmpPlan.free;
    end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 02/10/2006
Modifié le ... :   /  /
Description .. : bourre les comptes de début et fin de fourchette en
Suite ........ : fonction de la longueur de compte en compta et
Suite ........ : avec des 0 ou des 9
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSUBVENTION.BourreBorneCompte (var St : String ; Bourre : Char );
var Lg,ll,i : Integer ;
begin
  Lg:=VHImmo^.Cpta[fbGene].Lg ;
  If Length(St)>Lg Then
  begin
    St:=Trim(Copy(St,1,Lg)) ;
    exit;
  end;
  ll:=Length(St) ;
  If ll<Lg then
  BEGIN
    for i:=ll+1 to Lg do St:=St+Bourre ;
  END;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 02/10/2006
Modifié le ... :   /  /
Description .. : en sortie de zone compte : on alimente le cpte quotepart
Mots clefs ... :
*****************************************************************}

Procedure TOF_AMSUBVENTION.OnExitCompteSub(Sender: TObject);
var compte, compteQP, CompteRP : string;
begin
  compte := TRIM(GetControlText ('SB_CPTSUB'));
  BourreBorneCompte(Compte, '0');
  SetControlText('SB_CPTSUB',Compte);

  CompteRP := CompteReprise(GetControlText ('SB_CPTSUB'));
  BourreBorneCompte(CompteRP, '0');
  SetControlText ('SB_REPSUB', CompteRP);

  CompteQP := CompteQuotePart(GetControlText ('SB_CPTSUB'));
  BourreBorneCompte(CompteQP, '0');
  SetControlText ('SB_QUOTEPART', compteQP);

end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 02/10/2006
Modifié le ... :   /  /
Description .. : controle la validité du radical saisi et si le compte existe.
Mots clefs ... :
*****************************************************************}

Function TOF_AMSUBVENTION.ControleCompte : boolean;
var Compte :string;

  L : TList;

begin
  result := false;

  Compte := GetControlText('SB_CPTSUB');
  BourreBorneCompte(Compte, '0');
  SetControlText ('SB_CPTSUB', compte);
  Compte := TRIM(GetControlText ('SB_CPTSUB'));

  if ((Compte< BorneMiniSub) or (Compte > BorneMaxiSub)) and (copy(compte,1,3) <> '138') then
  begin
      PGIBox(HM[12]);
      result := True;
      SetControlText('SB_CPTSUB','');
      SetControlText('SB_REPSUB', '');
      SetControlText('SB_QUOTEPART', '');
      SetFocusControl('SB_CPTSUB');
      exit;
  end;

  if result = false then
  begin
    L := TList.Create;

    if (presence('GENERAUX','G_GENERAL',Compte) = false) then
        AjouteUnCompteSBV(L, Compte);

    if (presence('GENERAUX','G_GENERAL',BorneRP) = false) then
       AjouteUnCompteSBV(L, BorneRP);

    if (presence('GENERAUX','G_GENERAL',BorneRP) = false) then
       AjouteUnCompteSBV(L, BorneQP);

    {on ne pouvait pas sortir tant que tous les comptes n'avaient pas été créés
    While (L <> nil) and (L.count <> 0) do
    begin
      CreationComptesImmo (L);
      // la création désalloue les records
    end; }

    // ajout mbo 10.07.07 car si pas de cpte à créer on sortait toujours faux
    if L.count = 0 then L := nil;

    // mbo fq 19694 - 18.06.07 -
    if (L <> nil) and (L.count <> 0) then
    begin
      CreationComptesImmo (L);
      // la création désalloue les records
    end;

    if (L<> nil) then
       result := true;

  end;
end;
{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 02/10/2006
Modifié le ... :   /  /
Description .. : aide ref et création des comptes de subvention
Suite ........ : (radical 131 à 138)
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSUBVENTION.OnCompteSUBelipsisClick(Sender: TObject);
var
  stWhere: string;

begin
  stWhere :=  '(G_GENERAL>= "' + borneminiSub + '" AND G_GENERAL<= "' + bornemaxiSub +'")';
  stWhere :=  stWhere + ' OR  (G_GENERAL>= "' + RadMiniSub + '" AND G_GENERAL<= "' + RadMaxiSub +'")';

  {$IFDEF SERIE1}
    LookupList(TControl(Sender), TraduireMemoire('Comptes'), 'GENERAUX',
      'G_GENERAL', 'G_LIBELLE', stWhere, 'G_GENERAL', True, 3);
  {$ELSE}
    LookupList(TControl(Sender), TraduireMemoire('Comptes'), 'GENERAUX',
      'G_GENERAL', 'G_LIBELLE', stWhere, 'G_GENERAL', True, 1);
  {$ENDIF}

end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 03/10/2006
Modifié le ... :   /  /
Description .. : à partir du compte de subvention on reconstitue le compte
Suite ........ : quote part à savoir 777 + radical du compte de subvention
Suite ........ : à partir de la 3ème position
Mots clefs ... :
*****************************************************************}
Function TOF_AMSUBVENTION.CompteQuotePart(CompteSub : string): string;
begin
  if copy(CompteSub, 1, 3) = '131' then
     BorneQP := '777' + COPY(CompteSub, 3, (length (CompteSub))-3)
  else
     BorneQP := '777' + Copy(CompteSub, 4, (length (CompteSub))-3);

  Result := BorneQP;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 23/10/2006
Modifié le ... :   /  /
Description .. : à partir du compte de subvention on reconstitue le compte
Suite ........ : de reprise à savoir 139 + radical du compte de subvention
Suite ........ : à partir de la 3ème position
Mots clefs ... :
*****************************************************************}
Function TOF_AMSUBVENTION.CompteReprise(CompteSub : string): string;
begin
  if Copy(CompteSub, 1, 3) = '131' then
     BorneRP := '139' + COPY(CompteSub, 3, (length (CompteSub))-3)
  else
     BorneRP := '139' + copy(CompteSub, 4, (length (CompteSub))-3);

  Result := BorneRP;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 02/10/2006
Modifié le ... :   /  /
Description .. : sortie par la croix ou annuler : question pour prise en compte
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSUBVENTION.OnbFermeCLick(Sender: TObject);
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


procedure TOF_AMSUBVENTION.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AMSUBVENTION.OnDelete ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 03/10/2006
Modifié le ... :   /  /
Description .. : Création de la méthode SBV
Suite ........ : Mise à jour des tables
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSUBVENTION.OnUpdate ;
var Ordre : integer;
    TLog : TOB;
    lib :String;
    erreur : Boolean;
    QPlan : TQuery;
    PlanNew : TPlanAmort;
    PremReprise : string;

begin
  Inherited ;
  ecran.ModalResult := mrYes;

   // Controle de la date de l'opération
   erreur := ControleDateOpe;
   IF erreur=True then
   Begin
       ecran.ModalResult :=mrNone;
       SetFocusControl('SB_DATEOP');
       exit;
   end;

   // Controle de la date d'octroi de la subvention
   erreur := ControleDateSub;
   IF erreur=True then
   Begin
       ecran.ModalResult :=mrNone;
       SetFocusControl('SB_DATESUB');
       exit;
   end;

   // Controle du montant de la subvention
   erreur := ControleBase;
   IF erreur=True then
   Begin
       ecran.ModalResult :=mrNone;
       SetFocusControl('SB_MNTSUB');
       exit;
   end;

   // controle de la durée d'inaliénabilité pour les NAM
   if (NAM) then
   begin
      erreur := ControleDuree;
      IF erreur=True then
      Begin
          ecran.ModalResult :=mrNone;
          SetFocusControl('SB_DUREE');
          exit;
      end;
   end;
   // controle du suramortissement déjà pratiqué
   erreur := ControleAnterieur;
   if erreur=True then
   Begin
        ecran.ModalResult := mrNone;
        SetFocusControl('SB_REPRISE');
        exit;
   end;

   // controle des comptes et création si nécessaire
   erreur := ControleCompte;
   if erreur=True then
   Begin
      ecran.ModalResult := mrNone;
      SetFocusControl('SB_CPTSUB');
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
      QPlan.FindField('I_SBVMT').AsFloat:= Valeur(GetControlText('SB_MNTSUB'));

      if (NAM) and (QPlan.FindField('I_CORVRCEDDE').AsFloat = 0) then
      begin
         QPlan.FIndField('I_CORVRCEDDE').AsFloat:= StrToFloat(GetControlText('SB_DUREE'));
      end;

      QPlan.FIndField('I_CORRECTIONVR').AsFloat:= Valeur(GetControlText('SB_REPRISE'));
      // Enregistrement des comptes liés à la subvention
      
      // MVG 24/11/2006 avant asfloat.
      QPlan.FIndField('I_CPTSBVB').AsString:= GetControlText('SB_CPTSUB');
      // mbo - 03.11.06 - on ne stocke plus le compte de reprise - il sera retrouvé à partir du compte de sbv
      //QPlan.FIndField('I_CPTSBVR').AsFloat:= Valeur(GetControlText('SB_REPSUB'));

      Qplan.FindField('I_SBVDATE').AsDateTime :=  StrToDate(GetControlText('SB_DATESUB'));

      if (fDateDeb < VHImmo^.Encours.Deb) and
         (StrToDate(GetControlText('SB_DATESUB'))>fDateDeb) and
         (StrToDate(GetControlText('SB_DATESUB')) >= VHImmo^.Encours.Deb) then
      begin
         Qplan.FindField('I_DPIEC').AsString := 'X';
         PremReprise := 'X';
      end else
      begin
         Qplan.FindField('I_DPIEC').AsString := '-';
         PremReprise := '-';
      end;
      QPlan.Post;

      // Maj du plan d'amortissement IMMOAMOR
      QPlan.Edit;
      PlanNew:=TPlanAmort.Create(true) ;
      try
         PlanNew.Charge(QPlan);
         PlanNew.Recupere(fCode_immo,QPlan.FindField('I_PLANACTIF').AsString);
         fPlanActifAv := PlanNew.NumSeq;
         PlanNew.SetTypeOpe('SBV');

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
      // ATTENTION : stockage du montant de la subvention dans IL_MONTANTEXC

      lib:= 'Subvention octroyée le ' + GetControlText('SB_DATESUB');
      Ordre := TrouveNumeroOrdreLogSuivant(fPlan.CodeImmo);

      TLog := TOB.Create ('IMMOLOG',nil,-1);
      try
         TLog.PutValue('IL_IMMO',fCode_immo);
         TLog.PutValue('IL_LIBELLE',lib);
         TLog.PutValue('IL_TYPEMODIF',AffecteCommentaireOperation('SBV'));

         TLog.PutValue('IL_DATEOP', StrToDate(GetControlText('SB_DATEOP')));
         TLog.PutValue('IL_METHODEECO', MethodeSBV);

         if (NAM) then
              TLog.PutValue('IL_DUREEECO', StrToInt(GetControlText('SB_DUREE')))
         else
              TLog.PutValue('IL_DUREEECO', DureeSBV);

         TLog.PutValue('IL_MONTANTEXC', GetControlText('SB_MNTSUB'));

         TLog.PutValue('IL_TYPEOP', 'SBV');
         TLog.PutValue('IL_ORDRE', Ordre);
         TLog.PutValue('IL_OrdreSerie', -1);
         TLog.PutValue('IL_PLANACTIFAV', fPlanActifAv);
         TLog.PutValue('IL_PLANACTIFAP', fPlanActifAp);
         TLog.PutValue('IL_BLOCNOTE', RichToString (fBlocnote));
         // ajout mbo 10.11.06 pour stocker le fait que l'on est en reprise s/ 1ère dotation
         TLog.PutValue('IL_CODECB', PremReprise);

         TLog.InsertDB(nil);
      finally
         TLog.Free;
      end;

      COMMITTRANS ;

      VHImmo^.ChargeOBImmo := True;
      ImMarquerPublifi (True);

   except
      PGIBox(HM[9]);
      ROLLBACK ;
   end ;

   TFVIERGE(ECRAN).retour := IntToStr(MrYes);

end ;

procedure TOF_AMSUBVENTION.AjouteUnCompteSBV(var L:TList; Compte:string);
var Arecord : ^TDefCompte;
begin
    System.New (ARecord);
    ARecord^.Compte := Compte;
    ARecord^.Libelle := '';
    L.Add(ARecord);
end;

procedure TOF_AMSUBVENTION.OnLoad ;
begin
  Inherited ;
end ;


procedure TOF_AMSUBVENTION.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AMSUBVENTION.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AMSUBVENTION.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_AMSUBVENTION ] ) ;
end.
