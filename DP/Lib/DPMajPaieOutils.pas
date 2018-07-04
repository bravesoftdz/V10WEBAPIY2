{***********UNITE*************************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 16/06/2008
Modifié le ... :   /  /
Suite  ....... : FQ 12293 La personne peut avoir une autre fonction que gérant (GRT)
Mots clefs ... :
*****************************************************************}
unit DPMajPaieOutils;
/////////////////////////////////////////////////////////////////
interface
/////////////////////////////////////////////////////////////////
uses classes,
     {$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} // TQuery
     {$ENDIF}
     HCtrls,
     HEnt1,                                 // begintrans
     UTob,
     CLASSHistorisation,
     DpJurOutils,
     PgOutils2,
     Variants,
     hmsgbox,
     Controls,
     ParamSoc,
{utobdebug,     // pour consulter les tob en debug}
     SysUtils;

/////////////////////////////////////////////////////////////////
procedure DPMajPersonne_Paie (Dossier_p, Salarie_p, GuidPer_p : string; ExternePaie_p : Boolean; PFonction_p : string);  // FQ 12293
/////////////////////////////////////////////////////////////////
function  DPSSUniqueSalarie (Dossier_p, NoSS_p : string; ExternePaie_p : Boolean; var MsgErr, Salarie : string) : Boolean;
/////////////////////////////////////////////////////////////////
procedure DPMajDP_Paie (DD, DF : TDateTime; Dossier_p, Fonction_p : string);
/////////////////////////////////////////////////////////////////
function  DPGerantUniqueDP (Dossier_p, NoSS_p, FormeJ_p : string; var GuidPer, PFonction : string) : Boolean; // FQ 12293
/////////////////////////////////////////////////////////////////
function  DPAutorisationMajDP_Paie (OBParam : TOB; Dossier_p : string) : Boolean;
/////////////////////////////////////////////////////////////////
function  DPMajRemuneration_Paie (OBParam_p : TOB; ExternePaie : Boolean; Fonction_p : string) : Boolean;
/////////////////////////////////////////////////////////////////////
procedure DPMajPaieDP (Dossier_p, GuidPer_d, GuidPer_p , noSS : string; PFonction_p : string );     // FQ 12293
/////////////////////////////////////////////////////////////////////
procedure DPMajSocial_Paie (Dossier_p, Salarie_p : string; ExternePaie_p : Boolean);
/////////////////////////////////////////////////////////////////////


implementation
/////////////////////////////////////////////////////////////////
// Mise à jour fiche Personne du gérant du dossier à partir de celle de la Paie

procedure DPMajPersonne_Paie (Dossier_p, Salarie_p, GuidPer_p : string; ExternePaie_p : Boolean; PFonction_p : string);
/////////////////////////////////////////////////////////////////
   function ConvertPGSituationFamil (Situation : string) : string;
   begin
      if Situation = 'CON'      then result := 'CCB'
      else if Situation = 'CEL' then result := 'CEL'
      else if Situation = 'DIV' then result := 'DIV'
      else if Situation = 'MAR' then result := 'MAR'
      else if Situation = 'PAC' then result := 'PAC'
      else if Situation = 'VEU' then result := 'VEU'
      else if Situation = '90'  then result := ' '
      {else if Situation = 'SEP'  then result := ??? }
      else                           result := ' ';
   end;
/////////////////////////////////////////////////////////////////
   function RechDomTTCivilite (Code, Prefixedbo : string) : string;
   var wsql : string;
       Q: TQuery;
   begin
      result := '';
      // pas de RechDom, la tablette TTCIVILITE est dans CHOIXCOD dans le dossier
      wsql := 'SELECT CC_LIBELLE FROM ' + Prefixedbo + 'CHOIXCOD'
            + ' WHERE CC_TYPE="CIV" AND CC_CODE="' + Code + '"';
      Q := OpenSQL(wsql, True);
      try
         if not (Q = nil) and not (Q.eof) then
            result := Q.Findfield ('CC_LIBELLE').AsString;;
      finally
         Ferme (Q);
      end;
   end;
/////////////////////////////////////////////////////////////////
   procedure ConvertEnJUTypeCivil (Civilite : string; var Typeciv, Cv, Cvl, Cva : string);
   var wsql : string;
       Q : TQuery;
       cleCivilite : string;
   begin
      // En entrée code tablette YYCIVILITE de la table COMMUN de la DB0
      // En sorties infos tablette JUTYPECIVIL <=> champs de table JUTYPECIVIL (même préfixe)

      if Civilite = 'MLE'       then cleCivilite := 'MELLE'      //tcMle
      else if Civilite = 'MME'  then cleCivilite := 'MADAME'     //tcMme
      else if Civilite = 'MR'   then cleCivilite := 'MONSIEUR'   //tcMr
      else                           cleCivilite := ' ';
      Typeciv := cleCivilite;

      Cv  := '';
      Cva := '';
      Cvl := '';
      wsql := 'SELECT JTC_CIVCOURTE,JTC_CIVABREGE,JTC_CIVCOUR FROM JUTYPECIVIL'
            + '  WHERE JTC_TYPECIV="' + cleCivilite + '"';
      Q := OpenSQL (wsql, True);
      if not Q.EOF then
      begin
         Cv  := Q.FindField('JTC_CIVCOURTE').AsString;
         Cvl :=Q.FindField('JTC_CIVCOUR').AsString;
         Cva := Q.FindField('JTC_CIVABREGE').AsString;
      end;
      Ferme(Q);
   end;
/////////////////////////////////////////////////////////////////
// IDE 06/2008 - FQ 12176
//
   function RechEnfMin (Salarie, Prefixedbo : string) : integer;
   var wsql : string;
       Q: TQuery;
   begin
      result := 0;
      wsql := 'SELECT count(PEF_SALARIE) ENFMIN FROM ' + Prefixedbo + 'ENFANTSALARIE'
            + ' WHERE PEF_ACHARGE="X" AND PEF_TYPEPARENTAL="001" AND PEF_SALARIE="'
            + salarie + '"';
      Q := OpenSQL(wsql, True);
      try
         if not (Q = nil) and not (Q.eof) then
            result := Q.Findfield ('ENFMIN').AsInteger;
      finally
         Ferme (Q);
      end;
   end;

/////////////////////////////////////////////////////////////////
// IDE 06/2008 - FQ 12176
//
   function RechAlsaMosel (Salarie, Prefixedbo : string) : string;

   var wsql : string;
       Q: TQuery;
   begin
      result := '-';

      wsql := 'SELECT ETB_REGIMEALSACE FROM ' + Prefixedbo + 'SALARIES LEFT JOIN ' + Prefixedbo + 'ETABCOMPL'
            + ' ON ETB_ETABLISSEMENT = PSA_ETABLISSEMENT'
            + ' WHERE PSA_SALARIE="' + salarie + '"';
      Q := OpenSQL(wsql, True);
      try
         if not (Q = nil) and not (Q.eof) then
            result := Q.Findfield ('ETB_REGIMEALSACE').AsString;
      finally
         Ferme (Q);
      end;
   end;
/////////////////////////////////////////////////////////////////


var OBSalarie, OBAnn, OBDPperso, OBAnl : TOB;
    Prefixedbo, Salaries,GuidPerDos, wsql : string;
    Typeciv, Cv, Cvl, Cva : string;
    Q , Q2 , Q3 : TQuery;
/////////////////////////////////////////////////////////////////
begin
     if ExternePaie_p then  Prefixedbo := 'DB' + Dossier_p + '.dbo.'
     else                   Prefixedbo := '';
     Salaries := Prefixedbo + 'SALARIES';

     OBSalarie := TOB.Create(Salaries, nil, -1);
     // pas de SelectDB car ne fonctionnerait pas sur table préfixée dossier ??
     wsql :=  'SELECT * FROM ' + Salaries + ' WHERE PSA_SALARIE="' + Salarie_p + '"';
     OBSalarie.LoadDetailDBFromSQL('', wsql, False);

     wsql :=  'SELECT * FROM ANNUAIRE WHERE ANN_GUIDPER="' + GuidPer_p + '"';
     Q := OpenSQL(wsql, True);

     // DPPERSO - IDE 06/2008 - FQ 12176
     wsql :=  'SELECT * FROM DPPERSO WHERE DPP_GUIDPER="' + GuidPer_p + '"';
     Q2 := OpenSQL(wsql, True);

     // ANNULIEN -  IDE 06/2008 - FQ 12176
     GuidPerDos := GetColonneSQL('DOSSIER','DOS_GUIDPER','DOS_NODOSSIER="' + Dossier_p + '"');
     wsql :=  'SELECT * FROM ANNULIEN WHERE ANL_GUIDPERDOS="' + GuidPerDos + '"'
                                  + ' AND ANL_GUIDPER="' + GuidPer_p + '"'
                                  + ' AND ANL_TYPEDOS="STE"'
                                  + ' AND ANL_FONCTION="' + PFonction_p + '"'  {"GRT"'}
                                  + ' AND ANL_NOORDRE=1' ;
     Q3 := OpenSQL(wsql, True);

     //
     try
       OBAnn := TOB.Create('LE DP ANNUAIRE', nil, -1);
       // pas de selectDB car bloque sur la clé GUID sur les caractères tirets !!!
       // pas de loaddetailDBfromsql sinon insertorupdatedb ne fonctionne pas
       OBAnn.LoadDetailDB('ANNUAIRE', '','', Q, False);

       // DPPERSO - IDE 06/2008 - FQ 12176
       OBDPperso := TOB.Create('LE DP PERSO', nil, -1);
       // pas de selectDB car bloque sur la clé GUID sur les caractères tirets !!!
       // pas de loaddetailDBfromsql sinon insertorupdatedb ne fonctionne pas
       OBDPperso.LoadDetailDB('DPPERSO', '','', Q2, False);

       // ANNULIEN -  IDE 06/2008 - FQ 12176
       OBAnl := TOB.Create('LE DP ANNULIEN', nil, -1);
       // pas de selectDB car bloque sur la clé GUID sur les caractères tirets !!!
       // pas de loaddetailDBfromsql sinon insertorupdatedb ne fonctionne pas
       OBAnl.LoadDetailDB('ANNULIEN', '','', Q3, False);


       // ANNUAIRE

       if OBAnn.Detail.Count = 1 then
       begin
          with OBAnn.Detail[0] do    // 1 seul enreg car requête faite sur la clé
          begin
            PutValue ('ANN_APNOM',      OBSalarie.Detail[0].GetValue('PSA_CIVILITE') + ' ' +
                                        OBSalarie.Detail[0].GetValue('PSA_LIBELLE')  + ' ' +
                                        OBSalarie.Detail[0].GetValue('PSA_PRENOM') );
            PutValue ('ANN_NOM1',       OBSalarie.Detail[0].GetValue('PSA_LIBELLE'));
            PutValue ('ANN_NOM2',       OBSalarie.Detail[0].GetValue('PSA_PRENOM'));
            PutValue ('ANN_NOM3',       OBSalarie.Detail[0].GetValue('PSA_NOMJF'));

            // RAZ Voie/Btq
            PutValue ('ANN_VOIENO',     ' ');
            PutValue ('ANN_VOIENOCOMPL',' ');
            PutValue ('ANN_VOIETYPE',   ' ');
            PutValue ('ANN_VOIENOM',    ' ');
            PutValue ('ANN_APRUE1',     OBSalarie.Detail[0].GetValue('PSA_ADRESSE1'));
            PutValue ('ANN_APRUE2',     OBSalarie.Detail[0].GetValue('PSA_ADRESSE2'));
            PutValue ('ANN_APRUE3',     OBSalarie.Detail[0].GetValue('PSA_ADRESSE3'));
            PutValue ('ANN_APCPVILLE',  OBSalarie.Detail[0].GetValue('PSA_CODEPOSTAL') + ' ' +
                                        OBSalarie.Detail[0].GetValue('PSA_VILLE'));
            PutValue ('ANN_APPAYS',     RechDom('TTPAYS', OBSalarie.Detail[0].GetValue('PSA_PAYS'), False));

            PutValue ('ANN_ALRUE1',     OBSalarie.Detail[0].GetValue('PSA_ADRESSE1'));
            PutValue ('ANN_ALRUE2',     OBSalarie.Detail[0].GetValue('PSA_ADRESSE2'));
            PutValue ('ANN_ALRUE3',     OBSalarie.Detail[0].GetValue('PSA_ADRESSE3'));
            PutValue ('ANN_ALCP',       OBSalarie.Detail[0].GetValue('PSA_CODEPOSTAL'));
            PutValue ('ANN_ALVILLE',    OBSalarie.Detail[0].GetValue('PSA_VILLE'));
            PutValue ('ANN_PAYS',       OBSalarie.Detail[0].GetValue('PSA_PAYS'));
            PutValue ('ANN_TEL1',       OBSalarie.Detail[0].GetValue('PSA_TELEPHONE'));
            PutValue ('ANN_TEL2',       OBSalarie.Detail[0].GetValue('PSA_PORTABLE'));

            PutValue ('ANN_SITUFAM',    ConvertPGSituationFamil(OBSalarie.Detail[0].GetValue('PSA_SITUATIONFAMIL')));
            PutValue ('ANN_NATIONALITE',OBSalarie.Detail[0].GetValue('PSA_NATIONALITE'));
            PutValue ('ANN_SEXE',       OBSalarie.Detail[0].GetValue('PSA_SEXE'));
            PutValue ('ANN_DATENAIS',   OBSalarie.Detail[0].GetValue('PSA_DATENAISSANCE'));
            PutValue ('ANN_VILLENAIS',  OBSalarie.Detail[0].GetValue('PSA_COMMUNENAISS'));
            PutValue ('ANN_NODEPTNAIS', OBSalarie.Detail[0].GetValue('PSA_DEPTNAISSANCE'));
            PutValue ('ANN_PAYSNAIS',   RechDom('TTPAYS', OBSalarie.Detail[0].GetValue('PSA_PAYSNAISSANCE'), False));

            if not VarIsNull (OBAnn.Detail[0].GetValue('ANN_TYPECIV')) then
               if TRIM(OBAnn.Detail[0].GetValue('ANN_TYPECIV')) = '' then
               begin
                  // Civilite := RechDomTTCivilite (OBSalarie.Detail[0].GetValue('PSA_CIVILITE'), Prefixedbo);
                  // PSA_CIVILITE vient de la tablette YYCIVILITE et non pas de TTCIVILITE
                  // PutValue ('ANN_CV',      RechDom('YYCIVILITE', OBSalarie.Detail[0].GetValue('PSA_CIVILITE'), False));
                  // PutValue ('ANN_CVL',     RechDom('YYCIVILITE', OBSalarie.Detail[0].GetValue('PSA_CIVILITE'), False));
                  // PutValue ('ANN_CVA',     OBSalarie.Detail[0].GetValue('PSA_CIVILITE'));
                  ConvertEnJUTypeCivil(OBSalarie.Detail[0].GetValue('PSA_CIVILITE'), Typeciv, Cv, Cvl, Cva);
                  PutValue ('ANN_TYPECIV', Typeciv);
                  PutValue ('ANN_CV',      Cv);
                  PutValue ('ANN_CVL',     Cvl);
                  PutValue ('ANN_CVA',     Cva);
               end;
          end;
          OBAnn.InsertOrUpdateDB (False);
       end;
       OBAnn.Free;

       // DPPERSO - IDE 06/2008 - FQ 12176

       if OBDPperso.Detail.Count = 1 then
       begin
          with OBDPperso.Detail[0] do    // 1 seul enreg car requête faite sur la clé
          begin
            PutValue ('DPP_VOIENO',       ' ');
            PutValue ('DPP_VOIENOCOMPL',  ' ');
            PutValue ('DPP_VOIETYPE',     ' ');
            PutValue ('DPP_VOIENOM',      ' ');

            PutValue ('DPP_ADRESSE1',   OBSalarie.Detail[0].GetValue('PSA_ADRESSE1'));
            PutValue ('DPP_ADRESSE2',   OBSalarie.Detail[0].GetValue('PSA_ADRESSE2'));
            PutValue ('DPP_ADRESSE3',   OBSalarie.Detail[0].GetValue('PSA_ADRESSE3'));
            PutValue ('DPP_CODEPOSTAL', OBSalarie.Detail[0].GetValue('PSA_CODEPOSTAL'));
            PutValue ('DPP_VILLE',      OBSalarie.Detail[0].GetValue('PSA_VILLE'));
            PutValue ('DPP_PAYS',       OBSalarie.Detail[0].GetValue('PSA_PAYS'));

            PutValue ('DPP_ENFMIN', RechEnfMin (Salarie_p, Prefixedbo ));

            if OBDPperso.Detail[0].GetValue('DPP_RETRAITE') = 0 then
               PutValue ('DPP_RETRAITE', GetParamSocSecur ('SO_AGERETRAITE', '', True) );

          end;
          OBDPperso.InsertOrUpdateDB (False);
       end;
       OBDPperso.Free;

       // ANNULIEN -  IDE 06/2008 - FQ 12176

       if OBAnl.Detail.Count = 1 then
       begin
          with OBAnl.Detail[0] do    // 1 seul enreg car requête faite sur la clé
          begin
            PutValue ('ANL_ESTNONCADRE', 'X');
            PutValue ('ANL_ALSACEMOSELLE', RechAlsaMosel (Salarie_p, Prefixedbo)  );
          end;
          OBAnl.InsertOrUpdateDB (False);
       end;
       OBAnl.Free;
       //

       // DPSOCIAL

       DPMajSocial_Paie (Dossier_p, Salarie_p, ExternePaie_p);

     finally
       Ferme(Q);
       Ferme(Q2);
       Ferme(Q3);
     end;
     OBSalarie.Free;

end;
/////////////////////////////////////////////////////////////////
function DPSSUniqueSalarie (Dossier_p, NoSS_p : string; ExternePaie_p : Boolean; var MsgErr, Salarie : string) : Boolean;
/////////////////////////////////////////////////////////////////
var TabErr : array[0..2] of string;
    Salaries, wsql : string;
    Q : TQuery;
begin
  TabErr [0] := TraduireMemoire('Le numéro de Sécurité Sociale n''a été attribué à aucun salarié dirigeant.');
  TabErr [1] := '';
  TabErr [2] := TraduireMemoire('Le numéro de Sécurité Sociale est attribué à plusieurs salariés dirigeants.');

  result := False;
  MsgErr := TabErr[0];
  Salarie := '';
  if ExternePaie_p then Salaries := 'DB' + Dossier_p + '.dbo.SALARIES'
  else                  Salaries := 'SALARIES';

  // NON combien de salariés dirigeants possèdent le no SS dans la Paie
  //wsql := 'SELECT count(PSA_SALARIE) FROM ' + Salaries
  //      + ' WHERE PSA_DADSPROF       = "13"'
  //      + ' AND   PSA_DADSCAT        = "01"'
  //      + ' AND   trim(PSA_NUMEROSS) = "' + NoSS_p + '"';
  // combien de salariés possèdent le no SS dans la Paie
  wsql := 'SELECT count(PSA_SALARIE) FROM ' + Salaries
        + ' WHERE trim(PSA_NUMEROSS) = "' + NoSS_p + '"';
  Q := OpenSQL(wsql, True);

  try
    if not Q.eof then
       case Q.Fields[0].AsInteger of
       0 : begin
             result := False;
             MsgErr := TabErr[0];
           end;
       1 : begin
             // un seul salarié
             result := True;
             MsgErr := TabErr[1];
             Salarie := GetColonneSQL(Salaries,'PSA_SALARIE','trim(PSA_NUMEROSS) = "' + NoSS_p + '"');
           end;
       else
           result := False;
           MsgErr := TabErr[2];
       end;

  finally
    Ferme(Q);
  end;
end;
/////////////////////////////////////////////////////////////////
// Gestion des màj rémunérations des gérants du dossier à partir de la Paie

procedure DPMajDP_Paie (DD, DF : TDateTime; Dossier_p, Fonction_p : string);
/////////////////////////////////////////////////////////////////
var OBParam : TOB;
    i: integer;
begin
   OBParam := TOB.Create('', nil, -1) ;
   if DPAutorisationMajDP_Paie (OBParam, Dossier_p) then
   begin
      for i:=0 to OBParam.Detail.count-1 do
      begin
//pgibox('ob.detail.count='+inttostr(OBParam.Detail.count)+ 'fonctioni='+OBParam.Detail[i].GetValue('FONCTION')+'Fonction_p='+Fonction_p,'DPMAJPAIE_DP: i='+ inttostr(i));
        OBParam.Detail[i].AddChampSupValeur('DATEDEBUT',  DD);
        OBParam.Detail[i].AddChampSupValeur('DATEFIN',    DF);
        DPMajRemuneration_Paie (OBParam.Detail[i], False, Fonction_p);
        if Fonction_p = 'CLOTURE' then
           DPMajSocial_Paie (Dossier_p, OBParam.Detail[i].GetValue('SALARIE'), False);
      end;
   end;
   OBParam.Free;
end;
/////////////////////////////////////////////////////////////////
function  DPGerantUniqueDP (Dossier_p, NoSS_p, FormeJ_p : string; var GuidPer, PFonction : string) : Boolean;
/////////////////////////////////////////////////////////////////
var GuidPerDos, wsql, sDirigeant : string;
    Q1, Q2 : TQuery;
begin
   result := False;
   GuidPerDos := GetColonneSQL('DOSSIER','DOS_GUIDPER','DOS_NODOSSIER="' + Dossier_p + '"');
   GuidPer := '';
   PFonction := 'GRT';

   // combien de personnes possèdent le no SS dans le DP
   wsql := 'SELECT count(ANN_GUIDPER) FROM ANNUAIRE'
         + ' WHERE ANN_NUMEROSS = "' + NoSS_p + '"'
         + ' AND   ANN_PPPM = "PP"';
   Q1 := OpenSQL(wsql, True);

//pgibox ('dossier='+dossier_p+'guidperdos='+guidperdos,'1: dpgerantunique: formej='+Formej_p);
   try
     if not Q1.eof then
//pgibox ('Q1 nbenrg annuaire='+inttostr(Q1.Fields[0].AsInteger),'2: dpgerantunique: formej='+Formej_p);

        if Q1.Fields[0].AsInteger = 1 then
        begin
            // une seule personne correspond => récup clé
            GuidPer := GetColonneSQL('ANNUAIRE', 'ANN_GUIDPER',
                       'ANN_NUMEROSS = "' + NoSS_p + '"' +
                       ' AND ANN_PPPM = "PP"');

            // test la personne est un gérant (SARL) ou un dirigeant non gérant (SA/SAS/SASU)
            // requête faite sur la forme juridique => plusieurs résultats
            wsql   := 'SELECT ANL_FONCTION FROM ANNULIEN'
                     + ' WHERE ANL_GUIDPERDOS = "' + GuidPerDos + '"'
                     + ' AND ANL_GUIDPER      = "' + GuidPer + '"'
                     + ' AND ANL_TYPEDOS      = "STE"'
                     + ' AND ANL_FORME        = "' + FormeJ_p + '"'
                     + ' AND ANL_NOORDRE      = 1';

            Q2 := OpenSQL(wsql, True);

            try
              // FQ 12293 if not (Q2 = nil) and not (Q2.eof) then
              //             result := True;
              while not Q2.eof do
              begin
                if FormeJ_p = 'SARL' then
                begin
//pgibox ('annulien:anl_fonction='+Q2.FindField('ANL_FONCTION').AsString,'3: SARL');
                   if Q2.FindField('ANL_FONCTION').AsString = 'GRT' then
                   begin
                      result := True;
                      PFonction := Q2.FindField('ANL_FONCTION').AsString;
                   end;
                end else
                begin
                   sDirigeant := GetColonneSQL('JUTYPEFONCT','JTF_DIRIGEANT','JTF_FONCTION="' + Q2.FindField ('ANL_FONCTION').AsString + '"');
//pgibox ('annulien:anl_fonction='+Q2.FindField('ANL_FONCTION').AsString+' sdirig='+sDirigeant,'3: SA');
                   if (Q2.FindField('ANL_FONCTION').AsString <> 'GRT')
                   and (sDirigeant = 'X') then
                   begin
                      result := True;
                      PFonction := Q2.FindField('ANL_FONCTION').AsString;
                   end;
                end;
                Q2.Next;
              end;
            finally
               Ferme(Q2);
            end;
        end;

   finally
     Ferme(Q1);
   end;

end;
/////////////////////////////////////////////////////////////////
function DPAutorisationMajDP_Paie (OBParam : TOB; Dossier_p : string) : Boolean;
/////////////////////////////////////////////////////////////////
var GuidPerDos, GuidPer, FormeJ, wsql, msg, salarie : string;
    TSalarie, TFille : TOB;
    i : integer;
    Fonction : string;
begin
  result := False;
  GuidPerDos := GetColonneSQL('DOSSIER','DOS_GUIDPER','DOS_NODOSSIER="' + Dossier_p + '"');
  Fonction := 'GRT';

  // Uniquement si SARL

  FormeJ := GetColonneSQL('JURIDIQUE','JUR_FORME','JUR_GUIDPERDOS="' + GuidPerDos + '"');
  if (FormeJ <> 'SARL') and (FormeJ <> 'SA') and (FormeJ <> 'SAS') and (FormeJ <> 'SASU') then exit; // FQ 12293

  // Uniquement les salariés dirigeants ayant un no SS

  wsql := 'SELECT PSA_SALARIE, PSA_NUMEROSS FROM SALARIES'
        + ' WHERE PSA_DADSPROF = "13"'
        + ' AND   PSA_DADSCAT = "01"'
        + ' AND   trim(PSA_NUMEROSS) <> ""';
  TSalarie := TOB.Create('', nil, -1);
  TSalarie.LoadDetailDBFromSQL('', wsql, False);
{tobdebug(Tsalarie);}

  for i:= 0 to TSalarie.Detail.count-1 do
  begin
//pgibox('TSalarie.Detail.count='+inttostr(TSalarie.Detail.count)+'psa_sal='+TSalarie.Detail[i].GetValue('PSA_SALARIE'),'DPAutorisationMajDP_Paie: i='+inttostr(i));
    // vérif unicité du n° de sécurité sociale dans la Paie
    if DPSSUniqueSalarie (Dossier_p, TSalarie.Detail[i].GetValue('PSA_NUMEROSS'), False, msg, salarie) then

       // vérif salarié correspond à une personne unique dans ANNUAIRE et à un GERANT ou DIRIGEANT
       if DPGerantUniqueDP (Dossier_p, TSalarie.Detail[i].GetValue('PSA_NUMEROSS'),FormeJ, GuidPer, Fonction) then  // FQ 12293
       begin
          TFille := TOB.Create ('', OBParam, -1);
          TFille.AddChampSupValeur ('DOSSIER',    Dossier_p);
          TFille.AddChampSupValeur ('GUIDPERDOS', GuidPerDos);
          TFille.AddChampSupValeur ('GUIDPER',    GuidPer);
          TFille.AddChampSupValeur ('SALARIE',    TSalarie.Detail[i].GetValue('PSA_SALARIE'));
          TFille.AddChampSupValeur ('FONCTION',   Fonction);  // FQ 12293
       end;
  end;

  TSalarie.Free;
  result := True;
end;
/////////////////////////////////////////////////////////////////
// Mise à jour rémunérations du gérant du dossier à partir de celles de la Paie sur une période

function DPMajRemuneration_Paie (OBParam_p : TOB; ExternePaie : Boolean ; fonction_p : string) : Boolean;
/////////////////////////////////////////////////////////////////
//  EXEMPLE D'APPEL DE DPMajRemuneration_Paie A PARTIR DU DP
//
//  Récupérer la clé salarié par DPSSUniqueSalarie
//  OBParam := TOB.Create('', nil, -1) ;
//  TobParam.AddChampSupValeur ('GUIDPERDOS', GuidPerDos);  clé dossier
//  TobParam.AddChampSupValeur ('GUIDPER',    GuidPer);     clé gérant
//  TobParam.AddChampSupValeur ('DOSSIER',    NoDossier);   code dossier
//  TobParam.AddChampSupValeur ('SALARIE',    no_salarie);  clé salarié
//  TobParam.AddChampSupValeur ('DATEDEBUT',  iDate1900);   date à zéro
//  TobParam.AddChampSupValeur ('DATEFIN',    iDate1900);   date à zéro
//
//  while not Fini do
//  begin
//     DecodeDate (DateDebut, wa, wm, wj);
//     DateFin := FinDeMois(EncodeDate(wa, wm, 1));
//     if DateFin > fGerant.DateFinPrec then
//        Fini := True
//     else
//     begin
//          TobParam.PutValue ('DATEDEBUT',  DateDebut);
//          TobParam.PutValue ('DATEFIN',    DateFin);
//          DPMajRemuneration_Paie (TobParam, True);
//          DateDebut := DateFin + 1;
//     end;
//  end;
//  OBParam.Free;
/////////////////////////////////////////////////////////////////

   function ThemeRem (Theme : string) : string;
   begin
      if Theme = '001'      then result := 'SAL'
      else if Theme = '002' then result := 'INI'
      else if Theme = '004' then result := 'AVT'
      else                       result := 'SAL';
   end;
/////////////////////////////////////////////////////////////////
   function GetMtrem (TobParam : TOB;  TypeRem : string; ExternePaie:Boolean; var Avantage : string): double;
   var HistoBulletin, wsql : string;
       TMt : TOB;
       montant : double;
       i : integer;
   begin
     montant := 0;
     Avantage := '';
     if ExternePaie then HistoBulletin := 'DB'+TobParam.GetValue('DOSSIER')+'.dbo.HISTOBULLETIN'
     else                HistoBulletin := 'HISTOBULLETIN';

     if TypeRem = '004' then
     begin
        // avantages en nature
        wsql  := 'SELECT SUM(PHB_MTREM) MTREM, PRM_AVANTAGENAT FROM '+ HistoBulletin
                {+ ' DB' + TobParam.GetValue('DOSSIER') + '.dbo.HISTOBULLETIN' }
                + ' LEFT JOIN REMUNERATION'
                + ' ON PHB_RUBRIQUE=PRM_RUBRIQUE'
                + ' WHERE PRM_THEMEREM="' + ThemeRem(TypeRem) + '"'
                + ' AND PHB_SALARIE="' + TobParam.GetValue('SALARIE') + '"'
                + ' AND PHB_DATEDEBUT >="' + USDateTime(TobParam.GetValue('DATEDEBUT')) + '"'
                + ' AND PHB_DATEFIN   <="' + USDateTime(TobParam.GetValue('DATEFIN')) + '"'
                + ' AND PHB_NATURERUB="AAA"'
                + ' GROUP BY PRM_AVANTAGENAT';
     end else
     begin
        // salaire, prime
        wsql  := 'SELECT SUM(PHB_MTREM) MTREM FROM ' + HistoBulletin
                {+ 'DB' + TobParam.GetValue('DOSSIER') + '.dbo.HISTOBULLETIN' }
                + ' LEFT JOIN REMUNERATION'
                + ' ON PHB_RUBRIQUE=PRM_RUBRIQUE'
                + ' WHERE PRM_THEMEREM = "' + ThemeRem(TypeRem) + '"'
                + ' AND PHB_SALARIE    = "' + TobParam.GetValue('SALARIE') + '"'
                + ' AND PHB_DATEDEBUT >="' + USDateTime(TobParam.GetValue('DATEDEBUT')) + '"'
                + ' AND PHB_DATEFIN   <="' + USDateTime(TobParam.GetValue('DATEFIN')) + '"'
                + ' AND PHB_NATURERUB  = "AAA"';
     end;

     TMt := TOB.Create('', nil, -1);
     TMt.LoadDetailDBFromSQL('', wsql, False);
     //tobdebug(TMt);

     if TypeRem = '004' then
     begin
       for i:= 0 to TMt.Detail.Count-1 do
       begin
          if varType(TMt.Detail[i].GetValue ('MTREM')) <> varString then // protège champ virtuel vide
             montant  := montant  + TMt.Detail[i].GetValue ('MTREM');
          Avantage := Avantage + TMt.Detail[i].GetValue ('PRM_AVANTAGENAT')+ ';';
       end;
     end else
     begin
       if TMt.Detail.Count = 1 then
          if varType(TMt.Detail[0].GetValue ('MTREM')) <> varString then // protège champ virtuel vide
             montant := TMt.Detail[0].GetValue ('MTREM');
     end;
     TMt.Free;
     result := montant;
   end;
/////////////////////////////////////////////////////////////////
   procedure EnregistreRemuneration (TobParam:TOB;  TypeRem:string; Mtrem:double; Avantage:string='');
   var wsql, sCle : string;
       iNoordre : integer;
       OBJumvtrem, OBHistoRem : TOB;
   begin
     iNoOrdre := MaxChampX ('JUMVTREM', 'JMR_NOORDRE',
                            'JMR_GUIDPERDOS = "' + TobParam.GetValue('GUIDPERDOS') + '" AND ' +
                            'JMR_FONCTION   = "' + TobParam.GetValue('FONCTION') + '" AND ' + {"GRT" AND ' +}
                            'JMR_TYPEREM    = "' + TypeRem + '" AND ' +
                            'JMR_GUIDPER    = "' + TobParam.GetValue('GUIDPER') + '"') + 1;

     // Création dans JUMVTREM

     OBJumvtrem := TOB.Create('JUMVTREM', nil, -1);
     with OBJumvtrem do
     begin
        PutValue ('JMR_GUIDPERDOS', TobParam.GetValue('GUIDPERDOS'));
        PutValue ('JMR_GUIDPER',    TobParam.GetValue('GUIDPER'));
        PutValue ('JMR_FONCTION',   TobParam.GetValue('FONCTION')); //'GRT');
        PutValue ('JMR_TYPEREM',    TypeRem);
        PutValue ('JMR_DATEDEB',    TobParam.GetValue('DATEDEBUT'));
        PutValue ('JMR_DATEFIN',    TobParam.GetValue('DATEFIN'));
        PutValue ('JMR_NOORDRE',    iNoOrdre);
        PutValue ('JMR_MONTANT',    Mtrem);
        if TypeRem = '004' then
           PutValue ('JMR_AVANTAGE', Avantage )
        else
           PutValue ('JMR_PERIODEREM', '001');           // brut mensuel
     end;
     OBJumvtrem.InsertDB(nil);

{     if TypeRem = '004' then
     begin
     // avantages en nature
        wsql:= 'INSERT INTO JUMVTREM '
             + '(JMR_GUIDPERDOS,JMR_GUIDPER,JMR_FONCTION,JMR_TYPEREM,JMR_DATEDEB,JMR_DATEFIN,'
             + 'JMR_NOORDRE,JMR_PERIODEREM,JMR_MONTANT,JMR_AVANTAGE) '
             + 'VALUES("'
             + TobParam.GetValue('GUIDPERDOS') + '","' + TobParam.GetValue('GUIDPER')
             + '","GRT","' + TypeRem + '","'
             + USDateTime(TobParam.GetValue('DATEDEBUT')) + '","'
             + USDateTime(TobParam.GetValue('DATEFIN'))  + '",'
             + IntToStr(iNoOrdre) + ',"001",' + StrfPoint(Mtrem)
             + ',"' + Avantage + '")';
             // strfpoint vu dans \compta\lib4\cumuls
     end else
     begin
        // salaire, prime
        wsql:= 'INSERT INTO JUMVTREM '
             + '(JMR_GUIDPERDOS,JMR_GUIDPER,JMR_FONCTION,JMR_TYPEREM,JMR_DATEDEB,JMR_DATEFIN,'
             + 'JMR_NOORDRE,JMR_PERIODEREM,JMR_MONTANT) '
             + 'VALUES("'
             + TobParam.GetValue('GUIDPERDOS') + '","' + TobParam.GetValue('GUIDPER')
             + '","GRT","' + TypeRem + '","'
             + USDateTime(TobParam.GetValue('DATEDEBUT')) + '","'
             + USDateTime(TobParam.GetValue('DATEFIN'))  + '",'
             + IntToStr(iNoOrdre) + ',"001",' + StrfPoint(MtRem) + ')';
             // strfpoint vu dans \compta\lib4\cumuls
     end;

     ExecuteSQL(wsql); }

     // Relecture enreg inséré

     wsql :=  'SELECT * FROM JUMVTREM WHERE JMR_GUIDPERDOS = "' + TobParam.GetValue('GUIDPERDOS') + '"'
             + ' AND JMR_GUIDPER  = "' + TobParam.GetValue('GUIDPER') + '"'
             + ' AND JMR_FONCTION = "' + TobParam.GetValue('FONCTION') + '"'  //"GRT"'
             + ' AND JMR_TYPEREM  = "' + TypeRem + '"'
             + ' AND JMR_NOORDRE  = '  + IntToStr(iNoOrdre) ;

     OBHistoRem := TOB.Create('', nil, -1);
     OBHistoRem.LoadDetailDBFromSQL('JUMVTREM', wsql, False);
     // loaddetaildb ne fonctionne pas (commit pas encore été fait ?)
     // on relit l'enreg qu'on vient d'insérer => wsql sur la clé => tob OBHistoRem réelle en fait

     sCle := 'JHM_GUIDPERDOS = "' + OBHistoRem.Detail[0].GetValue('JMR_GUIDPERDOS') + '" AND ' +
             'JHM_GUIDPER    = "' + OBHistoRem.Detail[0].GetValue('JMR_GUIDPER') + '" AND ' +
             'JHM_FONCTION   = "' + OBHistoRem.Detail[0].GetValue('JMR_FONCTION') + '" AND ' +
             'JHM_TYPEREM    = "' + OBHistoRem.Detail[0].GetValue('JMR_TYPEREM') + '" AND ' +
             'JHM_NOORDRE    = '  + IntToStr(OBHistoRem.Detail[0].GetValue('JMR_NOORDRE'));

//tobdebug(OBHistoRem.Detail[0]);
     // pour rendre utilisables les propriétés TOB et accès base de la fonction Historisation
     // OBHistoRem.Detail[0].VirtuelleToReelle ('JUMVTREM');   // inutile en principe, tob réelle

     // Création dans JUHISTOREM

     Historisation (OBHistoRem.Detail[0], 'JUMVTREM', 'JUHISTOREM', sCle,
                    ' JHM_NOHISTO DESC', 'JHM_NOHISTO', 'Modification');
     OBHistoRem.Free;
   end;
/////////////////////////////////////////////////////////////////
var wsql : string;
    MtRem : double;
    Avantage : string;
/////////////////////////////////////////////////////////////////
begin
  result := False;
  Avantage := '';

  {if AdrTParam <> 0 then
  begin
     TobParam := TOB (AdrTParam);
  end
  else
  begin
     exit;
  end; }


  BeginTrans;
  try
     /// ----------------------------------------------------------------
     /// Suppression des 2 rémunérations brutes mensuelles + avantages, sur la période
     /// ----------------------------------------------------------------

     wsql :=  'DELETE FROM JUMVTREM WHERE JMR_GUIDPERDOS = "' + OBParam_p.GetValue('GUIDPERDOS') + '"'
             + ' AND JMR_GUIDPER    = "'  + OBParam_p.GetValue('GUIDPER') + '"'
             + ' AND JMR_FONCTION   = "' + OBParam_p.GetValue('FONCTION') + '"'  {"GRT"'}
             + ' AND (((JMR_TYPEREM = "001" OR JMR_TYPEREM = "002") AND (JMR_PERIODEREM = "001"))'
             + ' OR (JMR_TYPEREM = "004"))'
             + ' AND (((JMR_DATEDEB >= "' + UsDateTime (OBParam_p.GetValue('DATEDEBUT')) + '")'
             + '     AND (JMR_DATEDEB <= "' + UsDateTime (OBParam_p.GetValue('DATEFIN')) + '"))'
             + ' OR ((JMR_DATEFIN >= "' + UsDateTime (OBParam_p.GetValue('DATEDEBUT')) + '")'
             + '     AND (JMR_DATEFIN <= "' + UsDateTime (OBParam_p.GetValue('DATEFIN')) + '")))';
     ExecuteSQL(wsql);

     /// ----------------------------------------------------------------------------
     /// Création / historisation salaire brut mensuel issu de la PAIE sur la période
     /// ----------------------------------------------------------------------------

     if fonction_p<>'DECLOTURE' then
     // MVG 30/06/2008 En déclôture, on ne fait que supprimer
     begin
       MtRem := GetMtrem (OBParam_p, '001', ExternePaie, Avantage);
       if MtRem <> 0 then
          EnregistreRemuneration (OBParam_p, '001', Mtrem);

       /// ------------------------------------------------------------------------------
       /// Création / historisation prime brute mensuelle issue de la PAIE sur la période
       /// ------------------------------------------------------------------------------

       MtRem := GetMtrem (OBParam_p, '002', ExternePaie, Avantage);
       if MtRem <> 0 then
          EnregistreRemuneration (OBParam_p, '002', Mtrem);

      /// ------------------------------------------------------------------------------
      /// Création / historisation avantage brut mensuel issus de la PAIE sur la période
      /// ------------------------------------------------------------------------------

      MtRem := GetMtrem (OBParam_p, '004', ExternePaie, Avantage);
      if MtRem <> 0 then
         EnregistreRemuneration (OBParam_p, '004', Mtrem, Avantage);

     end;
     CommitTrans;   // déplacé ici
     result := True;

  except
     RollBack;
  end;

end;

/////////////////////////////////////////////////////////////////
// IDE 27/06/08  - FQ 12176
// Mise à jour du DP par rapport à l'etat actuel du dossier paie
//
procedure DPMajPaieDP (Dossier_p, GuidPer_d, GuidPer_p , noSS : string; PFonction_p : string );
//
/////////////////////////////////////////////////////////////////
var msg , wsql : string;
    salarie : string;
    HistoBulletin : string;
    DateDeb , DateFin , DateFinPer: TDateTime ;
    iAnneeEntree, iMoisEntree, iJourEntree: word;
    TobParam : TOB;
    fini : boolean;
    wA, wM, wJ : word;
    Q: TQuery;


begin

  // datedeb =  1er jour de l'année precedant la date d'entrée dans le bureau
  // si 30/06/2008 -> 01/01/2007
  DecodeDate (V_PGI.DateEntree, iAnneeEntree, iMoisEntree, iJourEntree);
  DateFin := EncodeDate(iAnneeEntree, 1, 1) - 1;   // 1 jour avant le 1er jour de l'année d'entrée
  DecodeDate (DateFin,iAnneeEntree, iMoisEntree, iJourEntree);
  DateDeb := EncodeDate (iAnneeEntree, 1, 1);

  HistoBulletin := 'DB'+Dossier_p+'.dbo.HISTOBULLETIN' ;

  // on verifie qu'il y a un et un seul salarie avec le n° de securité sociale
  if DPSSUniqueSalarie (Dossier_p, noSS , true , msg , salarie) then
  begin

      // un seul salarié
      DateFin := Idate1900;
      wsql  := 'SELECT MAX(PHB_DATEFIN) DATEMAX FROM ' + HistoBulletin
                + ' LEFT JOIN REMUNERATION'
                + ' ON PHB_RUBRIQUE=PRM_RUBRIQUE'
                + ' WHERE PRM_THEMEREM = "SAL" '
                + ' AND PHB_SALARIE    = "' + salarie + '"'
                + ' AND PHB_NATURERUB  = "AAA"';
      Q := OpenSql(wsql, True);
      if (not Q.EOF) then
           DateFin := Q.FindField('DATEMAX').AsDateTime ;
      Ferme(Q);

      if (DateFin <> Idate1900) then
      begin

        msg :=  'Voulez-vous reprendre les salaires, primes et avantages en nature #10#13' +
                'du dossier de Paie du ' + DateToStr(DateDeb) +  ' au ' + DateToStr(DateFin) + ' ?'  ;

        if PGIAsk(msg, 'Rémunérations du salarié : ' + GetNomCompPer(GuidPer_p)) = mrYes then
           begin

             // MAJ Annuaire, dpperso et annulien
             DPMajPersonne_Paie (Dossier_p, Salarie, GuidPer_p , True, PFonction_p);  // FQ 12293

             // MAJ éléments de rémunérations

             Fini := False;
             TobParam := TOB.Create ('', nil, -1);
             TobParam.AddChampSupValeur ('GUIDPERDOS', GuidPer_d);
             TobParam.AddChampSupValeur ('GUIDPER',    GuidPer_p);
             TobParam.AddChampSupValeur ('DOSSIER',    Dossier_p);
             TobParam.AddChampSupValeur ('SALARIE',    Salarie);
             TobParam.AddChampSupValeur ('DATEDEBUT',  iDate1900);
             TobParam.AddChampSupValeur ('DATEFIN',    iDate1900);
             TobParam.AddChampSupValeur ('FONCTION',   PFonction_p);  // FQ 12293

             // boucle du 01/01 de l'année civile précédant la date d'entrée ( DateDeb) dans le bureau
             // à la plus haute date clôturéee dans le paie (DateFin)
             while not Fini do
             begin
               DecodeDate (DateDeb, wa, wm, wj);
               DateFinPer := FinDeMois(EncodeDate(wa, wm, 1));
               if DateFinPer > DateFin then
                  Fini := True
               else
               begin
                  TobParam.PutValue ('DATEDEBUT',  DateDeb);
                  TobParam.PutValue ('DATEFIN',    DateFinPer);
//tobdebug (TobParam);
                  DPMajRemuneration_Paie (TobParam, True,'HISTORIQUE');
                  DateDeb := DateFinPer + 1;
               end;
             end;
             TobParam.Free;

             msg := 'La reprise des rémunérations du salarié ' + GetNomCompPer(GuidPer_p) + ' est effectuée.' ;
             PGIInfo ( msg , 'Rémunérations du salarié :' );
           end;
      end;

  end else
  begin
      // soit aucun salairie , soit plusieurs
      PGIBox(msg, 'Lien d''une personne : ' + GetNomCompPer(GuidPer_p));
  end;


end;

/////////////////////////////////////////////////////////////////
// Màj paramètres sociaux du dossier à partir de ceux du salarié dans la Paie

procedure DPMajSocial_Paie (Dossier_p, Salarie_p : string; ExternePaie_p : Boolean);
var Prefixedbo, GuidPerDos, wsql : string;
    Q, Qw : TQuery;
    OBDPSocial, TFille : TOB;
    TauxAcc, TauxTransp : double;
begin
   if ExternePaie_p then  Prefixedbo := 'DB' + Dossier_p + '.dbo.'
   else                   Prefixedbo := '';
   GuidPerDos := GetColonneSQL ('DOSSIER','DOS_GUIDPER','DOS_NODOSSIER="' + Dossier_p + '"');

   // Récup taux cotisation accident référencé dans l'établissement du salarié
   TauxAcc := 0;
   wsql := 'SELECT PAT_DATEVALIDITE,PAT_TAUXAT FROM ' + Prefixedbo + 'TAUXAT'
             + ' LEFT JOIN ' + Prefixedbo + 'SALARIES'
             + ' ON PAT_ETABLISSEMENT = PSA_ETABLISSEMENT AND PAT_ORDREAT = PSA_ORDREAT'
             + ' WHERE PSA_SALARIE="' + Salarie_p + '"'
             + ' ORDER BY PAT_DATEVALIDITE DESC';
   Qw := OpenSQL(wsql, True);

   try
     if not Qw.eof then
        TauxAcc := Qw.FindField ('PAT_TAUXAT').AsFloat;   // 1er enreg = le plus récent
   finally
     Ferme(Qw);
   end;

   // Récup taux cotisation transport référencé dans l'établissement du salarié
   TauxTransp := 0;
   wsql := 'SELECT ETB_TAUXVERSTRANS FROM ' + Prefixedbo + 'ETABCOMPL'
             + ' LEFT JOIN ' + Prefixedbo + 'SALARIES'
             + ' ON ETB_ETABLISSEMENT = PSA_ETABLISSEMENT'
             + ' WHERE PSA_SALARIE="' + Salarie_p + '"';
   Qw := OpenSQL(wsql, True);

   try
     if not Qw.eof then
        TauxTransp := Qw.FindField ('ETB_TAUXVERSTRANS').AsFloat;
   finally
     Ferme(Qw);
   end;

   // DPSOCIAL

   OBDPSocial := TOB.Create('LE DP SOCIAL', nil, -1);
   wsql :=  'SELECT * FROM DPSOCIAL WHERE DSO_GUIDPER="' + GuidPerDos + '"';
   Q := OpenSQL(wsql, True);

   try
     if not Q.eof then
     begin
        OBDPSocial.LoadDetailDB('DPSOCIAL', '', '', Q, False);
        // 1 seul enreg en fait, car requête faite sur la clé
        TFille := OBDPSocial.FindFirst(['DSO_GUIDPER'], [GuidPerDos], FALSE);
     end else
     begin
        // pas d'enreg dans DPSOCIAL => le créer
        TFille :=  TOB.Create('DPSOCIAL', OBDPSocial, -1);
     end;

     //tobdebug(OBDPSocial);

     if (TFille <> nil) then
     begin
        with TFille do
        begin
          PutValue ('DSO_GUIDPER', GuidPerDos);
          if TauxAcc = 0 then PutValue ('DSO_ACCIDENT', '-')
          else                PutValue ('DSO_ACCIDENT', 'X');
          PutValue ('DSO_TXACCIDENT', TauxAcc);
          if TauxTransp = 0 then PutValue ('DSO_TRANSPORT', '-')
          else                   PutValue ('DSO_TRANSPORT', 'X');
          PutValue ('DSO_TXTRANSPORT', TauxTransp);
        end;
        OBDPSocial.InsertOrUpdateDB (False);
     end;

   finally
     OBDPSocial.Free;
     Ferme (Q);
   end;

end;

/////////////////////////////////////////////////////////////////
end.
/////////////////////////////////////////////////////////////////


