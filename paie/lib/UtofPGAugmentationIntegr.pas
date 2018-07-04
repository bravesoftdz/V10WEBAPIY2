{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 23/12/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGAUGMINTEGRATION ()
Mots clefs ... : TOF;PGAUGMINTEGRATION
*****************************************************************
PT1 22/05/2007 JL V_720 Gestion maj historique par avance
}
Unit UtofPGAugmentationIntegr ;

Interface

Uses
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     EdtREtat,
{$else}
     eMul,
     UtilEagl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     uTob,
     UTofPGMulAugmIntegration,
     EntPaie,
     Ed_Tools,
     ParamSoc,
     PGEdtEtat,
     HTB97,
     UTOF ;

Type
  TOF_PGAUGMINTEGRATION = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    WhereAugmentation : String;
    IArrondiAugm,PrecisionArrondi,PctAugmDec : Integer;
    TobParamHisto : Tob;
    Function MajSalaires (TFixe,TVar : Tob;Var TS,TAugmentation : Tob): Boolean;
    Function ConstruireTobSalairesFixe : Tob;
    Function ConstruireTobSalairesVar : Tob;
    Function RendSalaires(TypeSalaire : String;TS : Tob) : Double;
    Function ArrondiAugm(Montant : Double) : Double;
    procedure ExporterIntegration(Sender : Tobject);
    procedure CalcSalaireExport(TFixe,TVar : Tob;Var TS,TE : Tob);
    procedure MajHistoParAvance(TobParam : Tob;leSalarie,LeChamp,LaValeur : String;LaDate : TDateTime);
  end ;

Implementation

procedure TOF_PGAUGMINTEGRATION.OnUpdate ;
var i : Integer;
    LeSalarie : String;
    TobSalarie,TS,TFixe,TVariable,TobAugm,TA : Tob;
    Q : TQuery;
    Rep : Word;
    MajS : Boolean;
begin
  Inherited ;
  Rep := PGIAsk('Voulez-vous mettre à jour les fiches salariés');
  If Rep = MrNo then Exit;
  If GetParamSocSecur('SO_PGHISTOAVANCE',False) = True then//PT1
  begin
    Q := OpenSQL('SELECT DISTINCT PPP_PGINFOSMODIF,PPP_PGTYPEDONNE,PPP_PGTYPEINFOLS FROM PARAMSALARIE WHERE ##PPP_PREDEFINI## PPP_HISTORIQUE="X"',True);
    TobParamHisto := Tob.Create('MaTob',Nil,-1);
    TobParamHisto.LoadDetailDB('Table','','',Q,False);
    Ferme(Q);
  end;
  TFixe := ConstruireTobSalairesFixe;
  TVariable := ConstruireTobSalairesVar;
//  WhereAugmentation := ConvertPrefixe(WhereAugmentation,'PSA','PBG');
  Q := OpenSQL('SELECT BUDGETPAIE.* FROM BUDGETPAIE '+
  'LEFT JOIN SALARIES ON PBG_SALARIE=PSA_SALARIE '+
  WhereAugmentation,True);
  TobAugm := Tob.Create('BUDGETPAIE',Nil,-1);
  TobAugm.LoadDetailDB('BUDGETPAIE','','',Q,False);
  Ferme(Q);
 // Q := OpenSQL('SELECT * FROM SALARIES WHERE PSA_SALARIE IN (SELECT PBG_SALARIE FROM BUDGETPAIE '+WhereAugmentation+')',True);
 Q := OpenSQL('SELECT SALARIES.* FROM SALARIES LEFT JOIN BUDGETPAIE ON PBG_SALARIE=PSA_SALARIE '+WhereAugmentation,True);
  TobSalarie := Tob.Create('SALARIES',Nil,-1);
  TobSalarie.LoadDetailDB('SALARIES','','',Q,False);
  Ferme(Q);
  if ((ListeAug.nbSelected) > 0) and (not ListeAug.AllSelected) then
  begin
  InitMoveProgressForm (NIL,'Chargement des données pour les augmentations',
                    'Veuillez patienter SVP ...', ListeAug.NbSelected-1,
                    False,True);
    for i := 0 to ListeAug.NbSelected - 1 do
    begin
      ListeAug.GotoLeBookmark(i);
      {$IFDEF EAGLCLIENT}
      TFMAug.Q.TQ.Seek(ListeAug.Row - 1);
      {$ENDIF}
      LeSalarie := Q_MulAug.FindField('PSA_SALARIE').AsString;
      TS := TobSalarie.FindFirst(['PSA_SALARIE'], [LeSalarie], False);
      TA := TobAugm.FindFirst(['PBG_SALARIE'], [LeSalarie], False);
      if (TS <> nil) and (TA <> nil) then
      begin
        MajS := MajSalaires (TFixe,TVariable,TS,TA);
        if MajS then TS.UpdateDB(False);
        TA.UpdateDB(False);
      end;
      MoveCurProgressForm ('Salarié : '+
                             TS.GetValue('PSA_LIBELLE'));
    end;
    ListeAug.ClearSelected;
  end;
  if (ListeAug.AllSelected = TRUE) then
  begin
     InitMoveProgressForm (NIL,'Chargement des données pour les augmentations',
                    'Veuillez patienter SVP ...', Q_MulAug.RecordCount,
                    False,True);
                     {$IFDEF EAGLCLIENT}
   if (TFMAug.bSelectAll.Down) then
    TFMAug.Fetchlestous;
  {$ENDIF}
    Q_MulAug.First;

    while not Q_MulAug.EOF do
    begin
      LeSalarie := Q_MulAug.FindField('PSA_SALARIE').AsString;
      LeSalarie := Q_MulAug.FindField('PSA_SALARIE').AsString;
      TS := TobSalarie.FindFirst(['PSA_SALARIE'], [LeSalarie], False);
      TA := TobAugm.FindFirst(['PBG_SALARIE'], [LeSalarie], False);
      if (TS <> nil) and (TA <> nil) then
      begin
        MajS := MajSalaires (TFixe,TVariable,TS,TA);
        if MajS then TS.UpdateDB(False);
        TA.UpdateDB(False);
      end;
      MoveCurProgressForm ('Salarié : '+
                             TS.GetValue('PSA_LIBELLE'));
      Q_MulAug.Next;
    end;
    ListeAug.AllSelected := False;
  end;
  FiniMoveProgressForm;
  TobAugm.Free;
  TobSalarie.Free;
  If GetParamSocSecur('SO_PGHISTOAVANCE',False) = True then TobParamHisto.Free; //PT1
end ;

procedure TOF_PGAUGMINTEGRATION.OnArgument (S : String ) ;
var Q: TQuery;
    CodeArrondi : String;
    BExport : TToolBarButton97;
begin
  Inherited ;
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
  WhereAugmentation := ReadTokenPipe(S,';');
  SetControltext('DATEFIXE',DateToStr(Date));
  SetControltext('DATEVAR',DateToStr(Date));
  SetControlText('LIBSALAIRE1',GetParamSoc('SO_PGSALLIB1'));
  SetControlText('LIBSALAIRE2',GetParamSoc('SO_PGSALLIB2'));
  SetControlText('LIBSALAIRE3',GetParamSoc('SO_PGSALLIB3'));
  SetControlText('LIBSALAIRE4',GetParamSoc('SO_PGSALLIB4'));
  SetControlText('LIBSALAIRE5',GetParamSoc('SO_PGSALLIB5'));
  BExport := TToolBarButton97(GetControl('BEXPORTER'));
  If BExport <> Nil then BExport.OnClick := ExporterIntegration;
end ;

Function TOF_PGAUGMINTEGRATION.MajSalaires (TFixe,TVar : Tob;Var TS,TAugmentation : Tob) : Boolean;
var Champ,Etat : String;
    i : Integer;
    Montant,MontantInit,MontantTot,Pct : Double;
    MajFixe,MajVar : Double;
begin
     Result := False;
     MajFixe := 0;
     MajVar := 0;
     Etat := TAugmentation.GeTvalue('PBG_ETATINTAUGM');
     If (Etat = '004') or (Etat = '007') or (Etat = '011') or (Etat = '012') then Exit;
     MontantTot := RendSalaires('FIXE',TS);
     If GetCheckBoxState('CFIXE') = CbChecked then
     begin
          For i := 0 to TFixe.Detail.Count - 1 do
          begin
               Champ := TFixe.Detail[i].GetValue('SALAIRE');
               MontantInit := TS.GetValue(Champ);
               Montant := TAugmentation.GeTvalue('PBG_FIXEAP');
               If i < TFixe.Detail.Count - 1 then
               begin
                    If MontantTot <> 0 then Pct := MontantInit / MontantTot
                    else Pct := 0;
                    Montant := Montant * Pct;
                    Montant := ArrondiAugm(Montant);
                    MajFixe := MajFixe + Montant;
               end
               else Montant := Montant - MajFixe;
               If (Montant > 0) and (Montant <> MontantInit) then
               begin
                    TS.PutValue(Champ,Montant);
                    If GetParamSocSecur('SO_PGHISTOAVANCE',False) = True then  MajHistoParAvance(TobParamHisto,TS.GetValue('PSA_SALARIE'),Champ,FloatToStr(Montant),StrToDate(GetControlText('DATEFIXE'))); //PT1
                    Result := True;
               end;
               TAugmentation.PutValue('PBG_INTFIXE','X');
               TAugmentation.PutValue('PBG_DATEFIXE',StrToDate(GetControlText('DATEFIXE')));
               If TAugmentation.GetValue('PBG_INTVARIABLE')= 'X' then TAugmentation.PutValue('PBG_ETATINTAUGM','010')
               Else TAugmentation.PutValue('PBG_ETATINTAUGM','009');
          end;
     end;
     If GetCheckBoxState('CVARIABLE') = CbChecked then
     begin
          MontantTot := RendSalaires('VARIABLE',TS);
          For i := 0 to TVar.Detail.Count - 1 do
          begin
               Champ := TVar.Detail[i].GetValue('SALAIRE');
               MontantInit := TS.GetValue(Champ);
               Montant := TAugmentation.GeTvalue('PBG_VARIABLEAP');
               If i < TVar.Detail.Count - 1 then
               begin
                    If MontantTot <> 0 then Pct := MontantInit / MontantTot
                    else Pct := 0;
                    Montant := ArrondiAugm(Montant * Pct);
                    MajVar := MajFixe + MajVar;
               end
               else Montant := Montant - MajVar;
               If (Montant > 0) and (Montant <> MontantInit) then
               begin
                    TS.PutValue(Champ,Montant);
                    If GetParamSocSecur('SO_PGHISTOAVANCE',False) = True then  MajHistoParAvance(TobParamHisto,TS.GetValue('PSA_SALARIE'),Champ,FloatToStr(Montant),StrToDate(GetControlText('DATEVAR')));
                    Result := True;
               end;
               TAugmentation.PutValue('PBG_INTVARIABLE','X');
               TAugmentation.PutValue('PBG_DATEVAR',StrToDate(GetControlText('DATEVAR')));
               If TAugmentation.GetValue('PBG_INTFIXE')= 'X' then TAugmentation.PutValue('PBG_ETATINTAUGM','010')
               Else TAugmentation.PutValue('PBG_ETATINTAUGM','009');
          end;
     end;
end;

Function TOF_PGAUGMINTEGRATION.ConstruireTobSalairesFixe : Tob;
var LesSalaires,TempSalaire,Champ : String;
    TobSalaire,TS : tob;
    i : Integer;
begin
     LesSalaires := GetParamSoc('SO_PGAUGFIXE');
     Result := Nil;
     If LesSalaires = '' then Exit;
     TobSalaire := Tob.Create('LesSalaires',Nil,-1);
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
               TS := Tob.Create('FilleSalaire',TobSalaire,-1);
               TS.AddChampSupValeur('SALAIRE',Champ);
          end;
     end;
     Result := TobSalaire;
end;

Function TOF_PGAUGMINTEGRATION.ConstruireTobSalairesVar : Tob;
var LesSalaires,TempSalaire,Champ : String;
    TobSalaire,TS : tob;
    i : Integer;
begin
     LesSalaires := GetParamSoc('SO_PGAUGVARIABLE');
     Result := Nil;
     If LesSalaires = '' then Exit;
     TobSalaire := Tob.Create('LesSalaires',Nil,-1);
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
               TS := Tob.Create('FilleSalaire',TobSalaire,-1);
               TS.AddChampSupValeur('SALAIRE',Champ);
          end;
     end;
     Result := TobSalaire;
end;

Function TOF_PGAUGMINTEGRATION.RendSalaires(TypeSalaire : String;TS : Tob) : Double;
var LesSalaires,TempSalaire,Champ : String;
    Montant : Double;
    i : Integer;
begin
     LesSalaires := '';
     If TypeSalaire = 'FIXE' then LesSalaires := GetParamSoc('SO_PGAUGFIXE');
     If TypeSalaire = 'VARIABLE' then LesSalaires := GetParamSoc('SO_PGAUGVARIABLE');
     If LesSalaires = '' then
     begin
          Result := 0;
          Exit;
     end;
     Montant := 0;
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
               Montant := Montant + TS.GetDouble(champ);
          end;
     end;
     Result := Montant;
end;

Function TOF_PGAUGMINTEGRATION.ArrondiAugm(Montant : Double) : Double;
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

Procedure TOF_PGAUGMINTEGRATION.ExporterIntegration(Sender : Tobject);
Var Pages : TPageControl;
    StPages : String;
    i : Integer;
    LeSalarie : String;
    TobSalarie,TS,TFixe,TVariable,TobEdition,TE : Tob;
    Q : TQuery;
begin
     PGIInfo('L''export permet de visualiser les données qui seront intégré dans la fiche salarié,'+
     '#13#10mais aucune de ces données ne sera sauvegardée.',Ecran.Caption);
     Pages := TPageControl(GetControl('PAGES'));
     StPages := AglGetCriteres (Pages, FALSE);
     TFixe := ConstruireTobSalairesFixe;
  TVariable := ConstruireTobSalairesVar;
  TobEdition := Tob.Create('edition',Nil,-1);
//  WhereAugmentation := ConvertPrefixe(WhereAugmentation,'PSA','PBG');
   // Q := OpenSQL('SELECT * FROM SALARIES WHERE PSA_SALARIE IN (SELECT PBG_SALARIE FROM BUDGETPAIE '+WhereAugmentation+')',True);
 Q := OpenSQL('SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,PSA_LIBELLEEMPLOI,PSA_SALAIREMOIS1,'+
 'PSA_SALAIREMOIS2,PSA_SALAIREMOIS3,PSA_SALAIREMOIS4,PSA_SALAIREMOIS5,PSA_SALAIRANN1,'+
 'PSA_SALAIRANN2,PSA_SALAIRANN3,PSA_SALAIRANN4,PSA_SALAIRANN5,'+
 'PBG_VALIDERPAR,PBG_FIXEAV,PBG_FIXEAP,PBG_VARIABLEAV,PBG_VARIABLEAP,'+
 'PSE_RESPONSABS,PSE_RESPONSVAR FROM SALARIES LEFT JOIN BUDGETPAIE ON PBG_SALARIE=PSA_SALARIE '+
 'LEFT JOIN DEPORTSAL ON PSE_SALARIE=PSA_SALARIE '+WhereAugmentation,True);
  TobSalarie := Tob.Create('Lessalaries',Nil,-1);
  TobSalarie.LoadDetailDB('Lessalaries','','',Q,False);
  Ferme(Q);
  if ((ListeAug.nbSelected) > 0) and (not ListeAug.AllSelected) then
  begin
                    InitMoveProgressForm (NIL,'Chargement des données pour les augmentations',
                    'Veuillez patienter SVP ...', ListeAug.NbSelected-1,
                    False,True);
    for i := 0 to ListeAug.NbSelected - 1 do
    begin
      ListeAug.GotoLeBookmark(i);
      {$IFDEF EAGLCLIENT}
      TFMAug.Q.TQ.Seek(ListeAug.Row - 1);
      {$ENDIF}
      LeSalarie := Q_MulAug.FindField('PSA_SALARIE').AsString;
      TS := TobSalarie.FindFirst(['PSA_SALARIE'], [LeSalarie], False);
      if TS <> Nil then
      begin
           TE := Tob.Create('FilleEdition',TobEdition,-1);
           TE.AddChampSupValeur('PSA_SALARIE',TS.GetValue('PSA_SALARIE'),False);
           TE.AddChampSupValeur('PSA_LIBELLE',TS.GetValue('PSA_LIBELLE'),False);
           TE.AddChampSupValeur('PSA_PRENOM',TS.GetValue('PSA_PRENOM'),False);
           TE.AddChampSupValeur('PSA_ETABLISSEMENT',TS.GetValue('PSA_ETABLISSEMENT'),False);
           TE.AddChampSupValeur('PSA_LIBELLEEMPLOI',TS.GetValue('PSA_LIBELLEEMPLOI'),False);
           TE.AddChampSupValeur('PBG_VALIDERPAR',RechDom('PGRESPONSVARIABLE',TS.GetValue('PBG_VALIDERPAR'),False),False);
           If TS.GetValue('PSE_RESPONSABS') <> null then TE.AddChampSupValeur('PSE_RESPONSABS',RechDom('PGRESPONSABS',TS.GetValue('PSE_RESPONSABS'),False),False)
           else TE.AddChampSupValeur('PSE_RESPONSABS','',False);
           If TS.GetValue('PSE_RESPONSVAR') <> null then TE.AddChampSupValeur('PSE_RESPONSVAR',RechDom('PGRESPONSVARIABLE',TS.GetValue('PSE_RESPONSVAR'),False),False)
           else TE.AddChampSupValeur('PSE_RESPONSABS','',False);
           if (TS <> nil) then CalcSalaireExport (TFixe,TVariable,TS,TE);
      end;
      MoveCurProgressForm ('Salarié : '+
                             TS.GetValue('PSA_LIBELLE'));
    end;
    ListeAug.ClearSelected;
  end;
  if (ListeAug.AllSelected = TRUE) then
  begin
     InitMoveProgressForm (NIL,'Chargement des données pour les augmentations',
                    'Veuillez patienter SVP ...', Q_MulAug.RecordCount,
                    False,True);
                     {$IFDEF EAGLCLIENT}
   if (TFMAug.bSelectAll.Down) then
    TFMAug.Fetchlestous;
  {$ENDIF}
    Q_MulAug.First;

    while not Q_MulAug.EOF do
    begin
      LeSalarie := Q_MulAug.FindField('PSA_SALARIE').AsString;
      TS := TobSalarie.FindFirst(['PSA_SALARIE'], [LeSalarie], False);
      if TS <> nil then
      begin
           TE := Tob.Create('FilleEdition',TobEdition,-1);
           TE.AddChampSupValeur('PSA_SALARIE',TS.GetValue('PSA_SALARIE'),False);
           TE.AddChampSupValeur('PSA_LIBELLE',TS.GetValue('PSA_LIBELLE'),False);
           TE.AddChampSupValeur('PSA_PRENOM',TS.GetValue('PSA_PRENOM'),False);
           TE.AddChampSupValeur('PSA_ETABLISSEMENT',TS.GetValue('PSA_ETABLISSEMENT'),False);
           TE.AddChampSupValeur('PSA_LIBELLEEMPLOI',TS.GetValue('PSA_LIBELLEEMPLOI'),False);
           TE.AddChampSupValeur('PBG_VALIDERPAR',RechDom('PGRESPONSVARIABLE',TS.GetValue('PBG_VALIDERPAR'),False),False);
           If TS.GetValue('PSE_RESPONSABS') <> null then TE.AddChampSupValeur('PSE_RESPONSABS',RechDom('PGRESPONSABS',TS.GetValue('PSE_RESPONSABS'),False),False)
           else TE.AddChampSupValeur('PSE_RESPONSABS','',False);
           If TS.GetValue('PSE_RESPONSVAR') <> null then TE.AddChampSupValeur('PSE_RESPONSVAR',RechDom('PGRESPONSVARIABLE',TS.GetValue('PSE_RESPONSVAR'),False),False)
           else TE.AddChampSupValeur('PSE_RESPONSABS','',False);
           CalcSalaireExport (TFixe,TVariable,TS,TE);
      end;
      MoveCurProgressForm ('Salarié : '+
                             TS.GetValue('PSA_LIBELLE'));
      Q_MulAug.Next;
    end;
    ListeAug.AllSelected := False;
  end;
  FiniMoveProgressForm;
  TobSalarie.Free;
  LanceEtatTOB('E','PAU','PEX',TobEdition,True,True,False,Pages,'','',False,0,StPages);
  TobEdition.Free;
end;

procedure TOF_PGAUGMINTEGRATION.CalcSalaireExport(TFixe,TVar : Tob;Var TS,TE : Tob);
var Champ : String;
    i,NumSalaire : Integer;
    Montant,MontantInit,MontantTot,Pct : Double;
    MajFixe,MajVar : Double;
begin
     MajFixe := 0;
     MajVar := 0;
     NumSalaire := 1;
     MontantTot := RendSalaires('FIXE',TS);
     If GetCheckBoxState('CFIXE') = CbChecked then
     begin
          For i := 0 to TFixe.Detail.Count - 1 do
          begin
               Champ := TFixe.Detail[i].GetValue('SALAIRE');
               MontantInit := TS.GetValue(Champ);
               Montant := TS.GeTvalue('PBG_FIXEAP');
               If i < TFixe.Detail.Count - 1 then
               begin
                    If MontantTot <> 0 then Pct := MontantInit / MontantTot
                    else Pct := 0;
                    Montant := Montant * Pct;
                    Montant := ArrondiAugm(Montant);
                    MajFixe := MajFixe + Montant;
               end
               else Montant := Montant - MajFixe;
               TE.AddChampSupValeur('MTSALAIREACT'+IntToStr(NumSalaire),FloatToStr(MontantInit));
               If (Montant > 0) and (Montant <> MontantInit) then TE.AddChampSupValeur('MTSALAIREAUG'+IntToStr(NumSalaire),FloatToStr(Montant))
               else TE.AddChampSupValeur('MTSALAIREAUG'+IntToStr(NumSalaire),'');
               NumSalaire := NumSalaire + 1;
          end;
     end;
     If GetCheckBoxState('CVARIABLE') = CbChecked then
     begin
          MontantTot := RendSalaires('VARIABLE',TS);
          For i := 0 to TVar.Detail.Count - 1 do
          begin
               Champ := TVar.Detail[i].GetValue('SALAIRE');
               MontantInit := TS.GetValue(Champ);
               Montant := TS.GeTvalue('PBG_VARIABLEAP');
               If i < TVar.Detail.Count - 1 then
               begin
                    If MontantTot <> 0 then Pct := MontantInit / MontantTot
                    else Pct := 0;
                    Montant := ArrondiAugm(Montant * Pct);
                    MajVar := MajFixe + MajVar;
               end
               else Montant := Montant - MajVar;
               TE.AddChampSupValeur('MTSALAIREACT'+IntToStr(NumSalaire),FloatToStr(MontantInit));
               If (Montant > 0) and (MontantInit <> Montant) then TE.AddChampSupValeur('MTSALAIREAUG'+IntToStr(NumSalaire),FloatToStr(Montant))
               else TE.AddChampSupValeur('MTSALAIREAUG'+IntToStr(NumSalaire),'');
               NumSalaire := NumSalaire + 1;
          end;
     end;
     For i := NumSalaire to 5 do
     begin
          TE.AddChampSupValeur('MTSALAIREACT'+IntToStr(i),0);
          TE.AddChampSupValeur('MTSALAIREAUG'+IntToStr(i),0);
     end;
end;

procedure TOF_PGAUGMINTEGRATION.MajHistoParAvance(TobParam : Tob;leSalarie,LeChamp,LaValeur : String;LaDate : TDateTime);   //PT1
var TP,TH : Tob;
begin
  TP := TobParam.FindFirst(['PPP_PGINFOSMODIF'],[LeChamp],False);
  If TP = Nil then Exit;
  TH := Tob.Create('PGHISTODETAIL',Nil,-1);
  TH.PutValue('PHD_SALARIE',LeSalarie);
  TH.PutValue('PHD_ETABLISSEMENT','');
  TH.PutValue('PHD_ORDRE',0);
  TH.PutValue('PHD_GUIDHISTO',AglGetGuid());
  TH.PutValue('PHD_PGINFOSMODIF',LeChamp);
  TH.PutValue('PHD_PGTYPEHISTO','003');
  TH.PutValue('PHD_ANCVALEUR','');
  TH.PutValue('PHD_NEWVALEUR',LaValeur);            
  TH.PutValue('PHD_TYPEVALEUR',TP.GetValue('PPP_PGTYPEDONNE'));
  TH.PutValue('PHD_TABLETTE',TP.GetValue('PPP_PGTYPEINFOLS'));
  TH.PutValue('PHD_PGTYPEINFOLS','SAL');
  TH.PutValue('PHD_DATEAPPLIC',LaDate);
  TH.PutValue('PHD_TRAITEMENTOK','X');
  TH.PutValue('PHD_DATEFINVALID',IDate1900);
  TH.PutValue('PHD_TYPEBUDG','');
  TH.PutValue('PHD_NUMAUG',-1);
  TH.PutValue('PHD_ANNEE','');
  TH.InsertDB(Nil,False);
end;

Initialization
  registerclasses ( [ TOF_PGAUGMINTEGRATION ] ) ;
end.


