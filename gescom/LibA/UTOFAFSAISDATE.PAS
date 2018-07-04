{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 24/04/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : AFSAISDATE ()
Mots clefs ... : TOF;AFSAISDATE
*****************************************************************}
Unit UTOFAFSAISDATE ;

Interface

Uses StdCtrls, Controls, Classes,  forms, sysutils,  ComCtrls,
{$IFDEF EAGLCLIENT}
       Maineagl,
{$ELSE}
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db,Fe_Main ,
{$ENDIF}
     HCtrls, HEnt1, HMsgBox, UTOF,
     utofafpiecepro_mul,vierge,ParamSoc,UtilCutOff,dicobtp,AffaireUtil,UtofAFFormule ;

Type
  TOF_AFSAISDATE = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (StArgument  : String ) ; override ;
    Procedure OnClose                  ; override ;
  Private
    Erreur:Boolean;
    zori,DateOri,Test : string;          // stock origine appel fiche + date passée en paramètre si existe +
    bMemlookup : boolean;
    function DernierJourMois(sDate:string):string;
  end ;

Function AFLanceFiche_SaisieDate(Argument:string):variant;
Procedure AFLanceFiche_AfGenCutOffGlobale;
Procedure AFLanceFiche_AfGenCutCompta;

Implementation


{***********A.G.L.***********************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 13/08/2001
Modifié le ... :   /  /
Description .. : Appel de la fct avec les différents paramètres possible:
Suite ........ : Zori : a pour appréciaiton, arrete pour arrété période, cutoff
Suite ........ : pour cutoff, CUTOFFAFF pour le cutoff par affaire
Suite ........ : ZZdate : date de début sélection
Suite ........ : Titre : titre de la fiche
Suite ........ : Libelle : libellé à mettre avant la saisie de la date
               : Test : X si on veut faire un test de cotrl par rapport à la date passée
Mots clefs ... : GIGA;SAISIE DATE
*****************************************************************}
procedure TOF_AFSAISDATE.OnArgument (StArgument : String ) ;
  Var
  Critere, Champ, valeur  : String;
  x : integer;
  zDateF : string;
begin
  Inherited ;
  bMemlookup := V_PGI.LookUpLocate;
  zDateF := '';
   // Recup des critères
  Erreur := False;
  Critere := (Trim (ReadTokenSt (stArgument)));
  While (Critere <> '') do
    BEGIN
      if Critere <> '' then
          BEGIN
          X := pos (':', Critere);
          if x <> 0 then
             begin
              Champ := copy (Critere, 1, X - 1);
              Valeur := Copy (Critere, X + 1, length (Critere) - X);
             end;
             // A pour appréciation, ARRETE
          if Champ = 'ZORI' then Zori := Valeur;
          // PL le 22/04/03 : la date de debut est invisible par defaut sur l'ecran, mais on en a besoin en cut off
          if Champ = 'ZZDATEDEB' then
            Begin
              V_PGI.LookUpLocate := true;
              SetControlText ('ZZDATEDEB', DateTostr(FindeMois(PlusMois(strtodate(Valeur), -1))));
              SetControlVisible ('TZZDATEDEB', True);
              SetControlVisible ('ZZDATEDEB', True);
              SetControlVisible ('LISTECUTOFF', True);
              SetControlProperty('LISTECUTOFF', 'Plus', 'GROUP BY ACU_DATE ORDER BY ACU_DATE DESC');
//              if (GetParamSoc('SO_AFCLIENT') = cInClientAmyot) then
              SetControlVisible ('AVECCUTOFF', True);
            end;
          // fin PL le 22/04/03
          if Champ = 'ZZDATE' then Begin zDateF := Valeur; DateOri := Valeur; end;
          if Champ = 'TITRE' then begin Ecran.caption := Valeur; UpdateCaption (Ecran); end;
          if Champ = 'LIBELLE' then SetControlText ('TZZDATE', valeur);
          if Champ = 'TEST' then Test := valeur;
          END;
      Critere := (Trim (ReadTokenSt (stArgument)));
    END;

  if (zori = 'A') and (zDateF = '') then
    Begin  // recup date fin de mois en cours
      zDateF := DateToStr (FInDeMois (PlusDate (V_PGI.DateEntree, (-1), 'M')));
    End;

  if ((zDateF = '') and (getcontroltext ('ZZDATE') <> '')) then
    zDateF :=  getcontroltext ('ZZDATE');

  if (zDateF = '') then
    zDateF := DateToStr (V_PGI.DateEntree);

  SetControlText ('ZZDATE', zDateF);

  if (zori = 'A') then
    Begin
      Ecran.Caption := 'Saisie date d''appréciation';
      UpdateCaption (Ecran);
      SetControlText ('TZZDATE', 'Les missions seront appréciées à la date du ');
    End;

  SetControlVisible ('BFERME', True);

end ;



procedure TOF_AFSAISDATE.OnUpdate;
Var stF, stD, Msg :String;
    ret : integer;
begin
  inherited ;
  ret := 0;
  stD := GetControlText('ZZDATEDEB');
  stF := GetControlText('ZZDATE');

     // dans le cas d'arrété, je ne veux pas de controle par rapport à un exercice comptable !!!
  If (Zori = 'ARRETE') or (Zori = 'CUTOFF') or (Zori = 'CUTOFFAFF') then
    begin
      if Not IsValidDate (StF) then
        ret := 1
      else
      if (Zori = 'CUTOFF') or (Zori = 'CUTOFFAFF') then
        if (strtodate(stD) >= strtodate(stF)) then
          ret := 4;

    end
  else
    ret := CtrlDate (stF);

  if (ret = 1) then
    Begin
      PGIInfo ('la date saisie n''est pas valide', titre);
      Erreur := true;
      Lasterror := 1;
      exit;
    End
  else if (ret = 2)  then
    Begin
      PGIInfo ('la date saisie n''est pas dans un exercice ouvert', titre);
      Erreur := true;
      Lasterror := 1;
      exit;
    End
  else if (ret = 3)  then
    Begin
      PGIInfo ('la date saisie est inférieure à l''arrêté de période', titre);
      Erreur := true;
      Lasterror := 1;
      exit;
    End
  else if (ret = 4)  then
    Begin
      PGIInfo ('l''intervalle de dates saisi n''est pas valide', titre);
      Erreur := true;
      Lasterror := 1;
      exit;
    End;

  If (Test = 'X') then
    begin
           // on contrôle que la date saisie n'est pas inférieure à la date de départ
      if (Zori <> 'CUTOFFCPTA') and (StrTodate (stF) < StrToDAte (DateOri)) then
        begin
            PGIInfoAf ('la date saisie est inférieure à la date permise', titre);
            Erreur := true;
            Lasterror := 1;
            SetControlText ('ZZDATE', DateOri);
            exit;
        End
      else
        begin
          If Zori = 'ARRETE' then
            begin
              // traitement de l'arrêté
              Msg := Msg + 'Confirmez vous l''arrêté de période au '+ stF +' ?'+ chr (10)+ chr (10);
              Msg := Msg + 'Attention, à l''issu les opérations suivantes ne seront plus réalisable :' + chr (10);
              Msg := Msg + '- Validation de factures d''une date antérieure à la date de cut off' + chr (10);
              Msg := Msg + '- Appréciation à une date antérieure à la date de cut off' + chr (10);
              Msg := Msg + '- Saisie ou modification d''activité à une date antérieure à la date de cut off' + chr (10);
              If PGIAskAF (Msg, 'Arrêté période') <> mrYes then exit;
              // PL le 04/11/02 : l'arrêté de période est stocké dans la date de début d'activité qui commence
              // le jour suivant, d'où le +1
              // SetParamsoc('SO_AFDATEDEBUTACT', StrTodate(stF));
              SetParamsoc ('SO_AFDATEDEBUTACT', StrTodate (stF) + 1);
              ////////////////////////////
              PGIAsk ('Opération terminée ', 'Arrêté période') ;
              SetControlText ('ZZDATE', '01/01/1900'); // pour sortir
            end;
          If Zori='CUTOFF' then
            begin
              StF := DernierJourMois (stF);
              SetControlText ('ZZDATE', StF);
              if Not ToutSeulAff  then   // quelqu'un d'autre travaille sur la base
                begin
                  try
                    Msg := 'Confirmez vous le Cut Off au '+ stF  +' ?';
                    // Arrêté de période en liaison
                    if (ctxScot  in V_PGI.PGIContexte) then
                      // Dans Scot uniquement pour l'instant
                      // on est en cutoff général
                      begin
                        Msg := Msg + chr (10) + chr (10);
                        Msg := Msg + 'Attention, à l''issu un arrêté de période est automatiquement effectué à cette date' + chr (10);
                        Msg := Msg + 'et les opérations suivantes ne seront plus réalisable :' + chr (10);
                        Msg := Msg + '- Validation de factures d''une date antérieure à la date de cut off' + chr (10);
                        Msg := Msg + '- Appréciation à une date antérieure à la date de cut off' + chr (10);
                        Msg := Msg + '- Saisie ou modification d''activité à une date antérieure à la date de cut off' + chr (10);
                      end;

                    If PGIAskAF (Msg, 'Cut Off Global') <> mrYes then
                      exit;
                    // traitement Du Cut Off
                    Gener_Cutoff (StrtoDate (stD), StrtoDate (stF), (GetCheckBoxState('AVECCUTOFF') = cbChecked));
                  finally
                    Bloqueur ('AffToutSeul', False);  // PL le 14/10/02
                  end;
                end;
  //            SetControlText('ZZDATE','01/01/1900'); // pour sortir
              end;
            If Zori = 'CUTOFFAFF' then
              begin
                StD := DernierJourMois (stD);
                StF := DernierJourMois (stF);
                SetControlText ('ZZDATEDEB', StD);
                SetControlText ('ZZDATE', StF);
                            // traitement Du Cut Off par affaire
                If PGIAskAF ('Confirmez vous le Cut Off au '+ stF ,'Cut Off par affaire') <> mrYes then
                  erreur := true;
                // SetControlText('ZZDATE','01/01/1900'); // pour sortir
              end
            else
              if Zori = 'CUTOFFCPTA' then
                begin
                  if (trim (stF) <> trim (DateOri)) then StF := DernierJourMois (stF);
                  SetControlText ('ZZDATE', StF);
                  // traitement de la génération Compta des cut off
                  If PGIAskAF ('Confirmez vous la date du Cut Off au ' + stF ,'Génération Cut Off en Compta') = mrYes then
//                    AFGenerComptaCutOff (StrToDate (StF))
                  else
                    erreur := true ;
                end;
            end;
        end;
end;


function TOF_AFSAISDATE.DernierJourMois (sDate : string) : string;
  var
  Jour, Mois, Annee : word;
  dDate : TDateTime;
begin
  Result := '01/01/1900';
  dDate := strtodate (sDate);
  DecodeDate (dDate, Annee, Mois, Jour);
  if (Mois = 12) then
    begin
      Result := DateToStr (EncodeDate (Annee, 12, 31));
    end
  else
    begin
      dDate := EncodeDate (Annee, Mois + 1, 1);
      Result := DatetoStr (dDate - 1);
    end;
end;

Procedure TOF_AFSAISDATE.OnClose;
begin
  inherited;

  If Erreur then
    begin
      LastError := (-1);
      LastErrorMsg := '';
      Erreur := false;
      exit;     // il y eu une erreur, on reste sur la même fiche
    end;

  if (ctrldate (getcontroltext ('ZZDATE')) = 0)  then
    begin
      TfVierge(Ecran).retour :=  getcontroltext ('ZZDATE');
      If Zori = 'CUTOFFAFF' then  // Seulement le cut off par affaire car le cut off global est appelé directement ici, alors que là,
                                  // on renvoit une valeur de retour pour traiter le cut off dans la fiche appelante
        begin
          TfVierge(Ecran).retour :=  TfVierge(Ecran).retour + ';' + getcontroltext ('ZZDATEDEB');
          if (GetCheckBoxState('AVECCUTOFF') <> cbChecked) then
            TfVierge(Ecran).retour :=  TfVierge(Ecran).retour + ';SANSCUTOFF';
        end;
    end
  else
    TfVierge(Ecran).retour :=  '0';


  V_PGI.LookUpLocate := bMemlookup;
End;

Function AFLanceFiche_SaisieDAte (Argument : string) : Variant;
begin
  result := AGLLanceFiche ('AFF', 'AFSAISDATE', '', '', Argument);
end;

Procedure AFLanceFiche_AfGenCutOffGlobale;
begin
  V_PGI.ZoomOLE := True;   //pour passer la fiche en modal
  if (trim (GetParamSoc ('SO_AFFORMULCUTOFF')) = '') then
    begin
      PGIInfoAf ('Pour faire une génération, vous devez saisir la formule de calcul du Cut Off.', 'Génération Globale Cut Off');
      AFLanceFiche_FormuleCutOff ('');
    end;
  if (trim (GetParamSoc ('SO_AFFORMULCUTOFF')) <> '') then
  AFLanceFiche_SaisieDate ('TEST:X;ZORI:CUTOFF;TITRE:Génération Globale Cut Off;LIBELLE:Date du cut off;TEST:X;ZZDATE:'
                            + format ('%s', [DateToStr (GetParamsoc ('SO_DATECUTOFF'))]) + ';ZZDATEDEB:'
//                            + format ('%s', ['01/01/1900']));
                            + format ('%s', [DateToStr (GetParamsoc ('SO_AFDATEDEBCAB'))]));
  V_PGI.ZoomOLE := False;
end;

Procedure AFLanceFiche_AfGenCutCompta;
var vStArg:string;
begin
  vStArg := 'ZORI:CUTOFFCPTA;TEST:X;TITRE:Génération en Compta du Cut Off;LIBELLE:Date du cut off;ZZDATE:'
             + format ('%s', [DateToStr (GetParamsoc ('SO_DATECUTOFF'))]);
  V_PGI.ZoomOLE := True;
  AGLLanceFiche ('AFF', 'AFSAISDATE', '', '', vStArg);
  V_PGI.ZoomOLE := False;
end;


Initialization
  registerclasses ([ TOF_AFSAISDATE ]);
end.
