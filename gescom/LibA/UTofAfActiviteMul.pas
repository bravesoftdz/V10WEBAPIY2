unit UTofAfActiviteMul;

interface
uses  StdCtrls,Controls,Classes,Forms,sysUtils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF, AffaireUtil,Grids,EntGC,
      DicoBTP,Saisutil,Hstatus,M3FP,UTofAfBaseCodeAffaire,AfActivite,
      ActiviteUtil,
      AfUtilArticle,TraducAffaire, HQry, HTB97, FactUtil,UtilMulTrt,LicUtil, utilressource,
{$IFDEF EAGLCLIENT}
      eMul,MaineAGL,
{$ELSE}
       HDB, mul, db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ENDIF}
      UTob,ParamSoc,UtofAfActiviteGLoEcla, Confidentaffaire,
{$IFDEF BTP}
      CalcOleGenericBTP;
{$ENDIF}
Type
     TOF_AfActiviteMul = Class (TOF_AFBASECODEAFFAIRE)
     public
        procedure OnLoad ; override ;
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnClose ; override ;
        procedure TypeAppelChargeCle(var AvActionFiche:TActionFiche);override;
        procedure AfTraitementsActivite(AcsTypeTrait:string);
        procedure TraiteLesLignes(TobActivite:TOB);
        procedure TraiteToutesLesLignes(TobActivite:TOB);
     private
        GcsNewVisa:string;
        GcsNewStatut:string;
        GcbAffReadOnly:boolean;
        titre : string;
        znumapprec : Integer;
        TWC : String;
        bLigneALigne, bVisa, bStatut, bApprec : boolean;
       procedure ModifActivite(zaction ,zvisaf: string) ;
//       procedure ChargeTobComplete(Q : tquery;TobComplete: Tob);
       procedure MajLignesAct (TobAct : TOB;zaction,zvisaf : string) ;
//       procedure SetVISAStatut;
//       procedure SetAllVISAStatut;
    END ;

Type
     TOF_AfActiviteCon = Class (TOF_AFBASECODEAFFAIRE)
     public
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnUpdate ; override ;
     private
        titre : string;
     END ;

     const
	// libellés des messages de la TOF  AfActiviteMul
	TexteMsgAffaire: array[1..2] of string 	= (
          {1}        'Aucune modification n''a été demandée'
          {2}        ,'Code Prestation invalide'
                     );

Procedure AFLanceFiche_Mul_Activite      (Argument:string);
Procedure AFLanceFiche_Consult_Activite  (strArgument:string='');               // $$$JP 18/06/2003: passage argument possible
Procedure AFLanceFiche_Modif_Activite    (Argument:string);

implementation

{ TOF_AfActiviteMul}


procedure TOF_AfActiviteMul.OnLoad;
Begin
     inherited;

     TWC := RecupWhereCritere(TPageControl(TFMul(Ecran).Pages));
End;

procedure TOF_AfActiviteMul.OnClose;
Begin
 Inherited;
End;

procedure TOF_AfActiviteMul.TypeAppelChargeCle(var AvActionFiche:TActionFiche);
begin

if GcbAffReadOnly then AvActionFiche := taConsult else  AvActionFiche := taCreat;
end;

procedure TOF_AfActiviteMul.OnArgument(stArgument : String );
Var
    Critere, Champ, val           :string;
    x                             :integer;
    zaff,zdateD,zdateF,zrep       :string;
    zcombo,zvisaf,ztypact,zvisaa  :string;
    toppre,topfou,topfra          :boolean;
    ComboTypeArticle              :THMultiValComboBox;
begin
     // Pierre le 25/06/02 : on rend invisible le numapprec dans le cas ou l'on ne le passe pas en parametre
     SetControlVisible('ACT_NUMAPPREC',false);
     SetControlVisible('TACT_NUMAPPREC',false);
     ///////////////////////

     bApprec:=false;
     zaff:=''; zrep:=''; zvisaf:=''; znumapprec:=0; zvisaa:='';
     zdated := '';
     zdatef := '';
     ztypact := '';   //562
     titre := Ecran.caption;
     GcbAffReadOnly:=true;

     // Recup des critères
     Critere:=(Trim(ReadTokenSt(stArgument)));
     While (Critere <>'') do
     begin
          if Critere<>'' then
          begin
               X:=pos(':',Critere);
               if x<>0 then
               begin
                    Champ:=copy(Critere,1,X-1);
                    Val:=Copy (Critere,X+1,length(Critere)-X);
               end;

               //if Champ = 'AFA_AFFAIRE' then Zaff := Val;
               if Champ = 'AFA_AFFAIRE' then Zaff := Val;
               if Champ = 'AFA_REPRISEACTIV'then Zrep := Val;
               if Champ = 'ACT_DATEACTIVITE'then ZDateD := Val;
               if Champ = 'ACT_DATEACTIVITE_'then ZDateF := Val;
               if Champ = 'ACT_ETATVISAFAC'then ZVisaf := Val;
               if Champ = 'ACT_ETATVISA'then ZVisaa := Val;
               if champ = 'MONOSEL' then
                {$IFNDEF EAGLCLIENT}
                TFMul(Ecran).Fliste.Multiselection := false;
                {$ELSE}
                TFMul(Ecran).Fliste.MultiSelect := false;
                {$ENDIF}
               if champ = 'NUMAPPREC' then
               begin
                    bApprec := true;
                    znumapprec := StrToInt(Val);
                    SetControlVisible('ACT_NUMAPPREC',true);
                    SetControlVisible('TACT_NUMAPPREC',true);
               end;
               if Champ = 'ACT_TYPEACTIVITE'then Ztypact := Val;     // 562
          end;

          Critere := (Trim(ReadTokenSt(stArgument)));
     end;

     if zAff='' then GcbAffReadOnly:=false;
     SetControlText('ACT_AFFAIRE',zaff);

     // Init du code affaire dans la tof ancêtre
     inherited;

     if ((zdated = '') and (getcontroltext('ACT_DATEACTIVITE') <> '')) then
        zdated :=  getcontroltext('ACT_DATEACTIVITE');

     if ((zdatef = '') and(getcontroltext('ACT_DATEACTIVITE_') <> '')) then
        zdatef :=  getcontroltext('ACT_DATEACTIVITE_');

     if (ZDateD = '') then
        zDateD := DateToStr(V_PGI.DateEntree);

     if (ZDateF = '') then
        zDateF := DateToStr(V_PGI.DateEntree);

     SetControlText('ACT_DATEACTIVITE',zdateD);
     SetControlText('ACT_DATEACTIVITE_',zdateF);
     SetControlText('ACT_ETATVISAFAC',zvisaf);
     SetControlText('ACT_ETATVISA',zvisaa);
     if (ztypact = 'TOU') then SetControlText('ACT_TYPEACTIVITE','');

     // alimentation du mul_combo type ligne activite en fct de l'activite reprise
     toppre := false; topfra := false; topfou := false;
     zcombo := '';
     Trt_Activiterepris(zrep,zcombo,toppre,topfou,topfra);

     if not(toppre) then
     begin
          SetControlEnabled('ZPROD_PRES',false);
          SetControlEnabled('TZPROD_PRES',false);
     end;
     if not(topfra) then
     begin
          SetControlEnabled('ZPROD_FRAIS',true);
          SetControlEnabled('TZPROD_FRAIS',true);
     end;
     if not(topfou) then
     begin
          SetControlEnabled('ZPROD_FOUR',false);
          SetControlEnabled('TZPROD_FOUR',false);
     end;

     SetControlText('ACT_TYPEARTICLE',zcombo);

     //mcd 05/03/02
     ComboTypeArticle:=THMultiValComboBox(GetControl('ACT_TYPEARTICLE'));
     ComboTypeArticle.plus:=PlusTypeArticle;
     if ComboTypeArticle.Text='' then ComboTypeArticle.Text:=PlusTypeArticleText;
     //RechInfos;
   If Not GetParamSoc ('So_AfAppPoint') then
     begin
     SetControlVisible ('BvisaFac',False);
     SetControlVisible ('ACT_ETATVISAFAC',False);
     SetControlVisible ('TACT_ETATVISAFAC',False);
     end;
   Tfmul(Ecran).FiltreDisabled:=true;  //mcd 23/05/03
end;

procedure TOF_AfActiviteMul.AfTraitementsActivite (AcsTypeTrait:string);
  var
  F : TFMul ;
  Critere, Traits, sSelect : string;
  x : integer;
  champ, val : string;
  TobAct : TOB;

  begin
  GcsNewVisa := '';
  GcsNewStatut := '';
  Traits := '';
  bLigneALigne := false;
  bVisa := false;
  bStatut := false;
  F := TFMul(Ecran);

  if (F.typeaction = taConsult) then exit;

  if (F.FListe.NbSelected = 0) and (not F.FListe.AllSelected) then
    begin
    PGIInfoAf ('Veuillez sélectionner les lignes à traiter', titre);
    exit;
    end;

  // PL V571
  if (PGIAskAf ('Confirmez-vous le traitement à effectuer ?' + chr(10) + 'Attention, certains traitements ne sont pas réversibles.',
        titre) <> mrYes) then exit;


  if (AcsTypeTrait = '1') then    //562
    // Traitement général avec écran de choix des traitements
    begin
    Traits := AGLLanceFiche('AFF', 'AFACTIVITEGLOCHOI', '', '', '');   //pas de tof, pas appel fct
    if (Traits = '') then exit;


    Critere := (Trim (ReadTokenSt (Traits)));
    while (Critere <> '') do
        begin
          if Critere <> '' then
            begin
              x := pos ('=', Critere);
              if (x <> 0) then
                begin
                Champ := copy (Critere, 1, x - 1);
                Val := copy (Critere, x + 1, length(Critere) - x);
                end;

              if Champ = 'LIGNE' then bLigneALigne := true;
              if Champ = 'STATUT' then
                begin
                bStatut := true;
                GcsNewStatut := Val;
                end;
              if Champ = 'VISA' then
                begin
                bVisa := true;
                GcsNewVisa := Val;
                end;
            end;
          Critere := (Trim (ReadTokenSt (Traits)));
        end; // end while
    end
  else
    if (AcsTypeTrait = '2') then
      // Traitement particulier : visa de facturation
      begin
        bVisa := true;
        bStatut := false;
        GcsNewVisa := 'VIS';
        bLigneALigne := False;
      end;

  TobAct := TOB.Create ('', nil, -1);
  Try
    if (bLigneALigne = true) then
      // PL le 31/01/03 : Le SELECT * a été traité :
      // dans le cas ou l'on traite ligne à ligne, on ne peut pas réduire les champs sélectionnés : l'éclatement
      // a besoin de l'intégralité de la ligne à éclater pour la dupliquer...
      sSelect := 'SELECT * FROM ACTIVITE'
    else
      // PL le 15/04/03 : modif clé activite
      // dans le cas où l'on traite toutes les lignes, on peut réduire
      sSelect := 'SELECT ACT_TYPEACTIVITE,ACT_AFFAIRE,ACT_NUMLIGNEUNIQUE,ACT_RESSOURCE,ACT_DATEACTIVITE,ACT_TYPEARTICLE'
                    + ',ACT_NUMLIGNE,ACT_ACTIVITEREPRIS,ACT_ETATVISAFAC,ACT_VISEURFAC,ACT_DATEVISAFAC'
                    + ',ACT_NUMAPPREC  FROM ACTIVITE';


    TraiteEnregMulTable (F, sSelect
                        ,'ACT_TYPEACTIVITE;ACT_AFFAIRE;ACT_NUMLIGNEUNIQUE'
                        , 'ACTIVITE', 'ACT_AFFAIRE', 'ACTIVITE', TobAct, True);
      // Fin PL le 15/04/03

    // PL le 03/02/03 : on repositionne les flags de modification des champs à false pour qu'au moment de l'update
    // ne soient updatés que les champs réellement modifiés (à intégrer dans TraiteEnregMulTable)
    TOBAct.SetAllModifie (False);

    if (bLigneALigne = true) then
      // on traite ligne à ligne
      begin
        // Traitement ligne à ligne
        if (TOBAct.Detail.count <> 0) then TraiteLesLignes (TOBAct);
      end
    else
      // on traite en global
      begin
        // Traitement toutes les lignes
        if (TOBAct.Detail.count <> 0) then TraiteToutesLesLignes (TOBAct);
      end;
  finally
    TOBAct.Free;
    F.ChercheClick;
  end;
(*
    if F.FListe.AllSelected then
      BEGIN
      if Transactions(SetAllVISAStatut,3) <> oeOK then
            PGIBoxAf('Impossible de traiter toutes les lignes d''activité', F.Caption);

      F.FListe.AllSelected := false;
      END
    ELSE
      BEGIN
      DejaVisa := False;
      InitMove(F.FListe.NbSelected,'');
      for i := 0 to F.FListe.NbSelected-1 do
         BEGIN
         F.FListe.GotoLeBookMark(i);
{$IFDEF EAGLCLIENT}
         F.Q.TQ.Seek(F.FListe.Row-1);
         if F.Q.TQ.FindField('ACT_ETATVISAFAC').AsString = GcsNewVisa then // Ne pas viser les lignes déjà visées
{$ELSE}
         if F.Q.FindField('ACT_ETATVISAFAC').AsString = GcsNewVisa then // Ne pas viser les lignes déjà visées
{$ENDIF}
            begin
            try
            DejaVisa := True;
            bMemoVisa := bVisa;
            bVisa := false;
            if Transactions(SetVISAStatut,3) <> oeOK then
                PGIBoxAf('Impossible de traiter une ligne d''activité', F.Caption);
            finally
            bVisa := bMemoVisa;
            end;
            end
         else
         if Transactions(SetVISAStatut,3) <> oeOK then
                PGIBoxAf('Impossible de traiter une ligne d''activité', F.Caption);

         MoveCur(False);
         END;

      F.FListe.ClearSelected;
      FiniMove;
      if DejaVisa then
            PGIInfo('Certaines lignes d''activité déjà visées n''ont pas été modifiées', F.Caption);
      END;

    F.ChercheClick;
    *)
//    end;

  TToolBarButton97 (GetControl ('bSelectAll')).Down := false;
  end;

procedure TOF_AfActiviteMul.TraiteLesLignes(TobActivite:TOB);
  var
    //Q : TQuery;
    critere, champ, val : string;
    CodeArticle, Tiers, TypeActivite, Affaire, Ressource, DateActivite, TypeArticle, Statut, MontantPV, Qte, Visa : String;
    Stat1, Mont1, Qte1, Visa1, Stat2, Mont2, Qte2, Visa2, Stat3, Mont3, Qte3, Visa3, Stat4, Mont4, Qte4, Visa4 : String;
    x, i, iDerNumLig : integer;
    Result, sLigne : string;
    TOBMere, TOBL, TOBL1, TOBL2, TOBL3, TOBL4 : TOB;
    dQte, dQteOld : double;
begin

  for i := 0 to TobActivite.Detail.Count - 1 do
    begin
      TOBMere := TOB.Create ('', nil, -1);
      try
        TOBL := TobActivite.Detail[i];
        Statut := TOBL.Getvalue ('ACT_ACTIVITEREPRIS');
        MontantPV := TOBL.Getvalue ('ACT_TOTVENTE');
        Qte := TOBL.Getvalue ('ACT_QTE');
        Visa := TOBL.Getvalue ('ACT_ETATVISAFAC');

        CodeArticle := TOBL.Getvalue ('ACT_CODEARTICLE');
        Tiers := TOBL.Getvalue ('ACT_TIERS');
        TypeActivite := TOBL.Getvalue ('ACT_TYPEACTIVITE');
        Affaire := TOBL.Getvalue ('ACT_AFFAIRE');
        Ressource := TOBL.Getvalue ('ACT_RESSOURCE');
        DateActivite := TOBL.Getvalue ('ACT_DATEACTIVITE');
        TypeArticle := TOBL.Getvalue ('ACT_TYPEARTICLE');
        sLigne := 'TIERS=' + Tiers + ';AFF=' + Affaire + ';DATE=' + DateActivite + ';ART=' + CodeArticle + ';RESS=' + Ressource
                  + ';STATUT=' + Statut + ';MONTANTPV=' + MontantPV + ';QTE=' + Qte + ';VISA=' + Visa + ';';

        // On n'éclate pas les lignes facturées ou visées (si le visa de facturation est géré)
        if (Statut = 'FAC') or ((GetParamSoc ('SO_AFAPPPOINT') = true) and (Visa = 'VIS')) then
          continue;

        // mcd 12/06/02 Result:=AGLLanceFiche('AFF','AFACTIVITEGLOECLA','','',sLigne);
        result := AFLanceFiche_Eclat_Activ (sLigne);

        // Si rien n'a été changé on continue sur la ligne suivante
        if (Result = '') then
          continue;
        // Si on interromp le traitement ligne à ligne, on sort de la boucle
        if (Result = 'ZZ') then
          break;

        Stat1 := '';
        Mont1 := '';
        Qte1 := '';
        Visa1 := '';
        Stat2 := '';
        Mont2 := '';
        Qte2 := '';
        Visa2 := '';
        Stat3 := '';
        Mont3 := '';
        Qte3 := '';
        Visa3 := '';
        Stat4 := '';
        Mont4 := '';
        Qte4 := '';
        Visa4 := '';

        // Exploitation des réponses
        Critere := (Trim (ReadTokenSt (Result)));
        While (Critere <> '') do
          begin
            if Critere <> '' then
              begin
                X := pos ('=', Critere);
                if x <> 0 then
                  begin
                    Champ := copy (Critere, 1, X - 1);
                    Val := Copy (Critere, X + 1, length (Critere) - X);
                  end;
                if Champ = 'STAT1' then Stat1 := val else
                if Champ = 'MONT1' then Mont1 := val else
                if Champ = 'QTE1' then Qte1 := val else
                if Champ = 'VISA1' then Visa1 := val else
                if Champ = 'STAT2' then Stat2 := val else
                if Champ = 'MONT2' then Mont2 := val else
                if Champ = 'QTE2' then Qte2 := val else
                if Champ = 'VISA2' then Visa2 := val else
                if Champ = 'STAT3' then Stat3 := val else
                if Champ = 'MONT3' then Mont3 := val else
                if Champ = 'QTE3' then Qte3 := val else
                if Champ = 'VISA3' then Visa3 := val else
                if Champ = 'STAT4' then Stat4 := val else
                if Champ = 'MONT4' then Mont4 := val else
                if Champ = 'QTE4' then Qte4 := val else
                if Champ = 'VISA4' then Visa4 := val;
              end;

            Critere := (Trim (ReadTokenSt (Result)));
          end;


        // Création des TOB de mise à jour de la base
        if (Stat1 <> '') or (Mont1 <> '') or (Qte1 <>'') or (Visa1 <> '') then
          begin
            TOBL1 := TOB.Create ('ACTIVITE', TOBMere, -1);
            TOBL1.Dupliquer (TOBL, false, true, false);
            if (Stat1 <> '') then TOBL1.PutValue ('ACT_ACTIVITEREPRIS', Stat1);
            if (Mont1 <> '') then
              begin
                TOBL1.PutValue ('ACT_TOTVENTE', Mont1);
                TOBL1.PutValue ('ACT_TOTVENTEDEV', Mont1);
              end;
            if (Qte1 <> '') then
              begin
                dQte := Valeur (Qte1);
                dQteOld := Valeur (TOBL1.GetValue ('ACT_QTE'));
                if (dQteOld <> 0) then
                  begin
                    TOBL1.PutValue ('ACT_TOTPR', (TOBL1.GetValue ('ACT_TOTPR') / dQteOld) * dQte);
                    TOBL1.PutValue ('ACT_TOTPRCHARGE',(TOBL1.GetValue ('ACT_TOTPRCHARGE') / dQteOld) * dQte );
                    TOBL1.PutValue ('ACT_TOTPRCHINDI',(TOBL1.GetValue ('ACT_TOTPRCHINDI') / dQteOld) * dQte );
                  end;
                TOBL1.PutValue ('ACT_QTE', Qte1);
                TOBL1.PutValue ('ACT_QTEFAC', Qte1);
                TOBL1.PutValue ('ACT_QTEUNITEREF', '0');
                if (TOBL1.GetValue ('ACT_TYPEARTICLE') = 'PRE') then
                  TOBL1.PutValue ('ACT_QTEUNITEREF', ConversionUnite (TOBL1.GetValue ('ACT_UNITE'), VH_GC.AFMESUREACTIVITE, TOBL1.GetValue ('ACT_QTE')));

              end;
            if (Visa1 <> '') then
              begin
                if Visa1 = 'VIS' then TOBL1.PutValue ('ACT_NUMAPPREC', znumapprec) else TOBL1.PutValue ('ACT_NUMAPPREC', 0);
                TOBL1.PutValue ('ACT_ETATVISAFAC', Visa1);
                TOBL1.PutValue ('ACT_VISEURFAC', V_PGI.User);
                TOBL1.PutValue ('ACT_DATEVISAFAC', NowH);
              end;
          end;

        iDerNumLig := 1;
        if (Stat2 <> '') or (Mont2 <> '') or (Qte2 <> '') or (Visa2 <> '') then
          begin
          /////////////// PL le 15/04/03 : modif clé activité
            // Recherche du dernier numéro de ligne pour cette clé
            iDerNumLig := ProchainPlusNumLigneUniqueActivite (TypeActivite, Affaire);

(*            Q := nil;
            try
              // PL A passer en EAGL - AFFAIREEAGL
              Q := OpenSQL ('SELECT MAX(ACT_NUMLIGNE) FROM ACTIVITE WHERE ACT_TYPEACTIVITE="' + TypeActivite
                          + '" AND ACT_AFFAIRE="' + Affaire
                          + '" AND ACT_RESSOURCE="' + Ressource
                          + '" AND ACT_DATEACTIVITE="' + UsDateTime (strtodate (DateActivite))
                          + '" AND ACT_FOLIO="' + Folio
                          + '" AND ACT_TYPEARTICLE="' + TypeArticle
                          + '"', true);

              if Not Q.EOF then
                  iDerNumLig := Q.Fields[0].AsInteger;

            finally
              ferme (Q);
            end;
 *)
          ////////////// fin PL le 15/04/03 : modif clé activité

            TOBL2 := TOB.Create ('ACTIVITE', TOBMere, -1);
            TOBL2.Dupliquer (TOBL, false, true, false);
            if (Stat2 <> '') then TOBL2.PutValue ('ACT_ACTIVITEREPRIS', Stat2);
            if (Mont2 <> '') then
              begin
                TOBL2.PutValue ('ACT_TOTVENTE', Mont2);
                TOBL2.PutValue ('ACT_TOTVENTEDEV', Mont2);
              end;
            if (Qte2 <> '') then
              begin
                dQte := Valeur (Qte2);
                dQteOld := Valeur (TOBL2.GetValue ('ACT_QTE'));
                if (dQteOld <> 0) then
                  begin
                    TOBL2.PutValue ('ACT_TOTPR', (TOBL2.GetValue ('ACT_TOTPR') / dQteOld) * dQte);
                    TOBL2.PutValue ('ACT_TOTPRCHARGE', (TOBL2.GetValue ('ACT_TOTPRCHARGE') / dQteOld) * dQte);
                    TOBL2.PutValue ('ACT_TOTPRCHINDI', (TOBL2.GetValue ('ACT_TOTPRCHINDI') / dQteOld) * dQte);
                  end;
                TOBL2.PutValue ('ACT_QTE', Qte2);
                TOBL2.PutValue ('ACT_QTEFAC', Qte2);
                TOBL2.PutValue ('ACT_QTEUNITEREF', '0');
                if (TOBL2.GetValue ('ACT_TYPEARTICLE') = 'PRE') then
                  TOBL2.PutValue ('ACT_QTEUNITEREF', ConversionUnite (TOBL2.GetValue ('ACT_UNITE'), VH_GC.AFMESUREACTIVITE, TOBL2.GetValue ('ACT_QTE')));
              end;
            if (Visa2 <> '') then
              begin
                if Visa2 = 'VIS' then TOBL2.PutValue ('ACT_NUMAPPREC', znumapprec) else TOBL2.PutValue ('ACT_NUMAPPREC', 0);
                TOBL2.PutValue ('ACT_ETATVISAFAC', Visa2);
                TOBL2.PutValue ('ACT_VISEURFAC', V_PGI.User);
                TOBL2.PutValue ('ACT_DATEVISAFAC', NowH);
              end;
            //iDerNumLig := iDerNumLig + 1; PL le 15/04/03 : modif clé activité
            //TOBL2.PutValue ('ACT_NUMLIGNE', iDerNumLig); // PL le 15/04/03 : modif clé activité : on lui laisse le même numligne que la ligne mère
            TOBL2.PutValue ('ACT_NUMLIGNEUNIQUE', iDerNumLig); // PL le 15/04/03 : modif clé activité
            TOBL2.SetAllModifie (true);
          end;

          if (Stat3 <> '') or (Mont3 <> '') or (Qte3 <>'') or (Visa3 <> '') then
            begin
              TOBL3 := TOB.Create ('ACTIVITE', TOBMere, -1);
              TOBL3.Dupliquer (TOBL, false, true, false);
              if (Stat3 <> '') then TOBL3.PutValue ('ACT_ACTIVITEREPRIS', Stat3);
              if (Mont3 <> '') then
                begin
                  TOBL3.PutValue ('ACT_TOTVENTE', Mont3);
                  TOBL3.PutValue ('ACT_TOTVENTEDEV', Mont3);
                end;
              if (Qte3 <> '') then
                begin
                  dQte := Valeur (Qte3);
                  dQteOld := Valeur (TOBL3.GetValue ('ACT_QTE'));
                  if (dQteOld <> 0) then
                    begin
                      TOBL3.PutValue ('ACT_TOTPR', (TOBL3.GetValue ('ACT_TOTPR') / dQteOld) * dQte);
                      TOBL3.PutValue ('ACT_TOTPRCHARGE', (TOBL3.GetValue ('ACT_TOTPRCHARGE') / dQteOld) * dQte );
                      TOBL3.PutValue ('ACT_TOTPRCHINDI', (TOBL3.GetValue ('ACT_TOTPRCHINDI') / dQteOld) * dQte );
                    end;
                  TOBL3.PutValue ('ACT_QTE', Qte3);
                  TOBL3.PutValue ('ACT_QTEFAC', Qte3);
                  TOBL3.PutValue ('ACT_QTEUNITEREF', '0');
                  if (TOBL3.GetValue ('ACT_TYPEARTICLE') = 'PRE') then
                    TOBL3.PutValue ('ACT_QTEUNITEREF', ConversionUnite (TOBL3.GetValue ('ACT_UNITE'), VH_GC.AFMESUREACTIVITE, TOBL3.GetValue ('ACT_QTE')));
                end;
              if (Visa3 <> '') then
                begin
                  if Visa3 = 'VIS' then TOBL3.PutValue ('ACT_NUMAPPREC', znumapprec) else TOBL3.PutValue ('ACT_NUMAPPREC', 0);
                  TOBL3.PutValue ('ACT_ETATVISAFAC', Visa3);
                  TOBL3.PutValue ('ACT_VISEURFAC', V_PGI.User);
                  TOBL3.PutValue ('ACT_DATEVISAFAC', NowH);
                end;
              iDerNumLig := iDerNumLig + 1;
              TOBL3.PutValue ('ACT_NUMLIGNEUNIQUE', iDerNumLig); // PL le 15/04/03 : modif clé activité
              TOBL3.SetAllModifie (true);
            end;

          if (Stat4 <> '') or (Mont4 <> '') or (Qte4 <>'') or (Visa4 <> '') then
            begin
              TOBL4 := TOB.Create ('ACTIVITE', TOBMere, -1);
              TOBL4.Dupliquer (TOBL, false, true, false);
              if (Stat4 <> '') then TOBL4.PutValue ('ACT_ACTIVITEREPRIS', Stat4);
              if (Mont4 <> '') then
                begin
                  TOBL4.PutValue('ACT_TOTVENTE', Mont4);
                  TOBL4.PutValue('ACT_TOTVENTEDEV', Mont4);
                end;
              if (Qte4 <> '') then
                begin
                  dQte := Valeur (Qte4);
                  dQteOld := Valeur (TOBL4.GetValue ('ACT_QTE'));
                  if (dQteOld <> 0) then
                    begin
                      TOBL4.PutValue ('ACT_TOTPR', (TOBL4.GetValue ('ACT_TOTPR') / dQteOld) * dQte);
                      TOBL4.PutValue ('ACT_TOTPRCHARGE', (TOBL4.GetValue ('ACT_TOTPRCHARGE') / dQteOld) * dQte );
                      TOBL4.PutValue ('ACT_TOTPRCHINDI', (TOBL4.GetValue ('ACT_TOTPRCHINDI') / dQteOld) * dQte );
                    end;
                  TOBL4.PutValue ('ACT_QTE', Qte4);
                  TOBL4.PutValue ('ACT_QTEFAC', Qte4);
                  TOBL4.PutValue ('ACT_QTEUNITEREF', '0');
                  if (TOBL4.GetValue ('ACT_TYPEARTICLE') = 'PRE') then
                    TOBL4.PutValue ('ACT_QTEUNITEREF', ConversionUnite (TOBL4.GetValue ('ACT_UNITE'), VH_GC.AFMESUREACTIVITE, TOBL4.GetValue ('ACT_QTE')));
                end;
              if (Visa4 <> '') then
                begin
                  if Visa4 = 'VIS' then TOBL4.PutValue ('ACT_NUMAPPREC', znumapprec) else TOBL4.PutValue ('ACT_NUMAPPREC', 0);
                  TOBL4.PutValue ('ACT_ETATVISAFAC', Visa4);
                  TOBL4.PutValue ('ACT_VISEURFAC', V_PGI.User);
                  TOBL4.PutValue ('ACT_DATEVISAFAC', NowH);
                end;
              iDerNumLig := iDerNumLig + 1;
              TOBL4.PutValue ('ACT_NUMLIGNEUNIQUE', iDerNumLig); // PL le 15/04/03 : modif clé activité
              TOBL4.SetAllModifie (true);
            end;

        // le insertorupdate est ici justifié, on peut ajouter ou modifier des enregistrements
        TOBMere.InsertOrUpdateDB (false);

      finally
        TOBMere.Free;
      end;
    end;

end;


procedure TOF_AfActiviteMul.TraiteToutesLesLignes(TobActivite:TOB);
  var
  bChangeVisa : boolean;
  i, numapp : integer;
  begin
    if (Not bStatut) and (Not bVisa) then exit;

    for i := 0 to TobActivite.Detail.Count - 1 do
      begin
        bChangeVisa := bVisa;
        // On ne peut pas dé-viser les lignes facturées
        if (GcsNewVisa = 'ATT') then
          begin
            if (TobActivite.Detail[i].GetValue ('ACT_ACTIVITEREPRIS') = 'FAC') then
              bChangeVisa := false
            else
              // En appréciation, on peut vouloir dé-viser
              if bApprec then bChangeVisa := true;
          end;
        if bVisa then
        if bChangeVisa or (V_PGI.PassWord = CryptageSt (DayPass (Date))) then
        if (TobActivite.Detail[i].GetValue ('ACT_ETATVISAFAC') <> GcsNewVisa) then
          begin
          if GcsNewVisa <> 'VIS' then numapp := 0 else numapp := znumapprec;

          TobActivite.Detail[i].PutValue ('ACT_ETATVISAFAC', GcsNewVisa);
          TobActivite.Detail[i].PutValue ('ACT_VISEURFAC', V_PGI.User);
          TobActivite.Detail[i].PutValue ('ACT_DATEVISAFAC', NowH);
          TobActivite.Detail[i].PutValue ('ACT_NUMAPPREC', numapp);
          end;

        if bStatut then
          // On ne change pas le statut des lignes facturées ou visées (si le visa de facturation est géré)
          if ((TobActivite.Detail[i].GetValue ('ACT_ACTIVITEREPRIS') <> 'FAC')
            and ((GetParamSoc ('SO_AFAPPPOINT') <> true) or (TobActivite.Detail[i].GetValue ('ACT_ETATVISAFAC') <> 'VIS')))
            or (V_PGI.PassWord = CryptageSt (DayPass (Date)))  then
            if (TobActivite.Detail[i].GetValue ('ACT_ACTIVITEREPRIS') <> GcsNewStatut) then
              begin
                TobActivite.Detail[i].PutValue ('ACT_ACTIVITEREPRIS', GcsNewStatut);
              end;
      end;

    TobActivite.UpdateDB;
  end;

procedure TOF_AfActiviteMul.ModifActivite(zaction,zvisaf : string) ;
var  F : TFMul ;
     St :string;
    // L:THDBGrid;
    // Q:TQuery;
     TobCOmplete : TOB;
begin
F:=TFMul(Ecran);
//L:= F.FListe;
//Q:= F.Q;
{$IFNDEF EAGLCLIENT}
if (F.Fliste.Multiselection = false) then exit;
{$ELSE}
if (F.Fliste.MultiSelect = false) then exit;
{$ENDIF}

if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
   begin
   PGIInfoAf('Aucune ligne sélectionnée',titre);
   exit;
   end;


  if (zaction = '') and (zvisaf = '') then
  Begin
    PGIInfoAf(TexteMsgAffaire[1],titre);
    exit;
  End;


    St:= 'Confirmez vous la modification de ces lignes';
    If (PGIAskAf(st,titre)<> mrYes) then exit;



    // on crée une TOB de toutes les pieces sélectionnées
TobComplete := TOB.Create ('Liste des lignes',NIL, -1);

// A FAIRE SELECT * : le nombre d'enregistrement doit être limité. on a besoin d'un certain nombre de champs qui sont assez
// variables vu les différents traitements possibles, par contre en cherchant dans tous, on doit pouvoir réduire
// la sélection...
TraiteEnregMulTable (F,'SELECT * FROM ACTIVITE',
                        'ACT_TYPEACTIVITE;ACT_AFFAIRE;ACT_NUMLIGNEUNIQUE;ACT_RESSOURCE;ACT_DATEACTIVITE;ACT_FOLIO;ACT_TYPEARTICLE;ACT_NUMLIGNE',
                        'ACTIVITE', 'ACT_AFFAIRE', 'ACTIVITE', TobComplete, True);

MajLignesAct(TObCOmplete, zaction, zvisaf);

TobComplete.free;

FiniMove;
end;

procedure TOF_AfActiviteMul.MajLignesAct (TobAct : TOB;zaction ,zvisaf: string) ;
Var wi : integer;
    TobDet : TOB;
    toto : string;
BEGIN
  for wi:= 0 to TobAct.Detail.count-1  do
  Begin
    TobDet := TobAct.detail[wi];
    toto := TobDet.GetValue('ACT_AFFAIRE');
    if (zaction <>'') then
    	TobDet.putValue('ACT_ACTIVITEREPRIS',zaction);
    if (zvisaf <>'') then
    	begin
    	TobDet.putValue('ACT_ETATVISAFAC',zvisaf);
      TobDet.putValue('ACT_VISEURFAC',V_PGI.User);
      TobDet.PutValue('ACT_DATEVISA',date);
      TobDet.PutValue('ACT_NUMAPPREC',znumapprec);
      End;
  End;

  TobAct.InsertOrUpdateDB(false);

END;



{ TOF_AfActiviteCon}

procedure TOF_AfActiviteCon.Onupdate;
begin
Inherited;
{$IFDEF EAGLCLIENT}
TraduitAFLibGridSt(TFMul(Ecran).FListe);
{$ELSE}
TraduitAFLibGridDB(TFMul(Ecran).FListe);
{$ENDIF}
end;

procedure TOF_AfActiviteCon.OnArgument(stArgument : String );
var
    ComboTypeArticle    :THMultiValComboBox;
    Critere , Part0, Part1, Part2, Part3, Avenant            :string;
    Champ, Val          :string;
    ChampCtrl           :TControl;
    PageCtrl            :TPageControl;
    i, j                :integer;
begin
     inherited;

     titre := Ecran.caption;
     if (THEdit(GetControl('ACT_ETATVISA')) <> nil) AND (GetParamSoc('so_afVisaActivite') = False) then
     begin
          SetControlVisible('ACT_ETATVISA',False);
          SetControlVisible('TACT_ETATVISA',False);
          SetControlText ('ACT_ETATVISA','ATT;VIS');
     end;

     //mcd 05/03/02
     ComboTypeArticle:=THMultiValComboBox(GetControl('ACT_TYPEARTICLE'));
     ComboTypeArticle.plus:=PlusTypeArticle;
     if ComboTypeArticle.Text='' then
        ComboTypeArticle.Text := PlusTypeArticleText;

     // $$$jp: paramètres
     Critere := Trim (ReadTokenSt (stArgument));
     while Critere <> '' do
     begin
          // JP 09/12/2002: lien OLE - on initialise les champs spécifiés
          i := Pos ('=', Critere);
          if i <> 0 then
          begin
               Champ  := Copy (Critere, 1, i-1);
               Val := Copy (Critere, i+1, Length (Critere)-i);

               // $$$JP 20/06/03: nouvelle visualisation en lien OLE ou Suivi par planning
               if (Val <> '') AND (Champ = 'XX_WHERE') then
               begin
                    // Mise à "zéro" de tous les critères disponibles
                    PageCtrl := TPageControl (GetControl ('PAGES'));
                    for i := 0 to PageCtrl.PageCount-1 do
                    begin
                         for j := 0 to PageCtrl.Pages [i].ControlCount-1 do
                         begin
                              ChampCtrl := PageCtrl.Pages [i].Controls [j];
                              if ChampCtrl is THEdit then
                              begin
                                   THEdit (ChampCtrl).Text := ''
                              end
                              else if ChampCtrl is THMultiValComboBox then
                              begin
                                   THMultiValComboBox (ChampCtrl).Value := '';
                                   THMultiValComboBox (ChampCtrl).Complete := FALSE; // pour que l'AGL ne construise pas un bout du WHERE
                              end
                              else if ChampCtrl is THValComboBox then
                              begin
                                   THValComboBox (ChampCtrl).Value := '';
                              end;
                         end;
                    end;
                    PageCtrl.Visible := FALSE;
                    TToolWindow97 (GetControl ('PFILTRES')).Visible := FALSE;

                    // Activation du XX_WHERE
                    SetControlText ('XX_WHERE', Val);
               end
               // $$$JP 20/06/03: plus besoin, on passe par l'argument "XX_WHERE=..."
{               if (Val <> '') AND ((Copy (Champ, 1, 3) = 'OLE') OR (Copy (Champ, 1, 3) = 'SUI')) then // $$$JP 18/06/2003: gestion également depuis suivi activité + si valeur vide, on fait rien
               begin
                    ChampCtrl := GetControl (Copy (Champ, 4, Length (Champ)-3));
                    if ChampCtrl <> nil then
                    begin
                         if ChampCtrl is THEdit then
                             THEdit (ChampCtrl).Text := Val
                         else
                             if ChampCtrl is THMultiValComboBox then
                                 THMultiValComboBox (ChampCtrl).Value := Val
                             else
                                 if ChampCtrl is THValComboBox then
                                    THValComboBox (ChampCtrl).Value := Val;
                    end;
               end}
                //mcd 19/06/03
            else If Champ = 'AFF' then
              begin
              SetControlText('ACT_AFFAIRE',Val);
              SetControlEnabled ('ACT_AFFAIRE1',False);
              SetControlEnabled ('ACT_AFFAIRE2',False);
              SetControlEnabled ('ACT_AFFAIRE3',False);
              SetControlEnabled ('ACT_AVENANT',False);
              SetControlEnabled ('ACT_TIERS',False);
              setControlEnabled ('BSELECTAFF1',False);
              setControlEnabled ('BEFFACEAFF1',False);
              BchangeTiers :=False;
              {$IFDEF BTP}
              BTPCodeAffaireDecoupe(GetControlText('ACT_AFFAIRE'),Part0,Part1,Part2,Part3,Avenant,taModif,false);
              {$ELSE}
              CodeAffaireDecoupe(GetControlText('ACT_AFFAIRE'),Part0,Part1,Part2,Part3,Avenant,taModif,false);
              {$ENDIF}
              SetControlText('ACT_AFFAIRE1',Part1); SetControlText('ACT_AFFAIRE2',Part2);
              SetControlText('ACT_AFFAIRE3',Part3); SetControlText('ACT_AVENANT',avenant);
              end
            else if Champ ='DATEDEB' then SetControlText ('ACT_DATEACTIVITE',Val)
            else if Champ ='DATEFIN' then SetControlText ('ACT_DATEACTIVITE_',Val);
          end;
          // JP fin 09/12/2003

          // Paramètre suivant
          Critere := Trim (ReadTokenSt (stArgument));
     end;
     // $$$jp-fin
end;


procedure AppelActivite (stargument : string);
Var Critere, Champ, val  : String;
    x, iFolio : integer;
    zaff, zdate, zcli, zres, ztyp, zFolio : string;
    TypeSaisie : T_Sai;
    dDate : TDateTime;
Begin
ZFolio := '';

Critere := (Trim (ReadTokenSt (stArgument)));
While (Critere <> '') do
    BEGIN
    if Critere<> '' then
        BEGIN
        X := pos (':', Critere);
        if x <> 0 then
           begin
           Champ := copy (Critere, 1, X - 1);
           Val := Copy (Critere, X + 1, length (Critere) - X);
           end;
        if Champ = 'TYP' then Ztyp := Val;
        if Champ = 'AFF' then Zaff := Val;
        if Champ = 'CLI' then Zcli := Val;
        if Champ = 'RES' then Zres := Val;
        if Champ = 'DAT' then ZDate := Val;
        if Champ = 'FOL' then ZFolio := Val; // PL le 20/06/06 : ajout du Folio
        END;
    Critere := (Trim (ReadTokenSt (stArgument)));
    END;

    if (ztyp = 'CLI') then
       TypeSaisie := tsaClient
    else
        TypeSaisie := tsaRess;
    dDate := StrToDate (zdate);

    iFolio := 1;
    if (ZFolio <> '') then
      iFolio := strtoint (ZFolio);

    If Not (SaisieActivitemanager) then
      begin   //mcd 17/01/2003 si saisie manager on interdit la saisie sur un utilisatuer # du sien
      if Zres <> Vh_GC.ressourceUser  then
        begin
        PgiInfo ('Vous n''avez pas le droit de modifier cette ligne');
        exit;
        end
        // sinon accès par assistant uniquement
      else  begin
          // PL le 20/06/06 : ajout du Folio
        AFCreerActiviteModale (tsaRess, tacGlobal, 'REA', zres, zaff, zcli, ddate, iFolio);
//        AFCreerActiviteModale (tsaRess, tacGlobal, 'REA', zres, zaff, zcli, ddate);
        exit;
        end;
      end;

      // PL le 20/06/06 : ajout du Folio
    AFCreerActiviteModale (TypeSaisie, tacGlobal, 'REA', zres, zaff, zcli, ddate, iFolio);
//    AFCreerActiviteModale (TypeSaisie, tacGlobal, 'REA', zres, zaff, zcli, ddate);

End;


/////////////// Procedure appellé par le bouton Validation //////////////
procedure AGLAfTraitementsActivite(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then MaTOF:=TFMul(F).LaTOF else exit;
if (MaTOF is TOF_AfActiviteMul) then TOF_AfActiviteMul(MaTOF).AfTraitementsActivite(Parms[1]) else exit;
end;

procedure AGLAfModifActivite(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then MaTOF:=TFMul(F).LaTOF else exit;
if (MaTOF is TOF_AfActiviteMul) then TOF_AfActiviteMul(MaTOF).ModifActivite(Parms[1],Parms[2]) else exit;
end;

procedure AGLAppelActivite(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then MaTOF:=TFMul(F).LaTOF else exit;
if (MaTOF is TOF_AfActiviteCon) or (MaTOF is TOF_AfActiviteMul) then AppelActivite(Parms[1]) else exit;
end;

Procedure AFLanceFiche_Mul_Activite(Argument:string);
begin
AGLLanceFiche ('AFF','AFACTIVITE_MUL','','',Argument);
end;

// $$$JP 18/06/2003: argument
procedure AFLanceFiche_Consult_Activite (strArgument:string);
begin
     AGLLanceFiche ('AFF', 'AFACTIVITECON_MUL', '', '', strArgument);
end;

//Procedure AFLanceFiche_Consult_Activite;
//begin
//     AGLLanceFiche ('AFF','AFACTIVITECON_MUL','','','');
//end;

Procedure AFLanceFiche_Modif_Activite(Argument:string);
begin
AGLLanceFiche ('AFF','AFACTIVITEGLO_MUL','','',Argument);
end;


Initialization
registerclasses([TOF_AfActiviteMul]);
registerclasses([TOF_AfActiviteCon]);
RegisterAglProc('AfModifActivite',TRUE,2,AGLAfModifActivite);
RegisterAglProc('AfTraitementsActivite',TRUE,1,AGLAfTraitementsActivite);
RegisterAglProc( 'AppelActivite', True ,1, AGLAppelActivite);
end.


