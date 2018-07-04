{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 23/05/2005
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PGHISTODETAIL (PGHISTODETAIL)
Mots clefs ... : TOM;PGHISTODETAIL
*****************************************************************
PT1 18/07/2007 JL V_80 Ajout gestion elipsis pour salarié (affichage nom + prénom)
PT2 21/09/2007 FC V_80 FQ 14800 Filtrer les valeurs d'un élément dynamique en fonction de la convention collective
PT3 18/10/2007 GGU V_80 FQ 14881 Saisie des éléments dynamiques impossible en CWAS
PT4 16/11/2007 FC V_80 FQ 14931 Pb filtrage des valeurs d'un élément quand plusieurs prédéfinis paramétrés
PT5 28/01/2008 FC V810 FQ 15085 Si libellé, la valeur saisie ne pointe pas sur l'élément auquel elle appartient
PT6 17/04/2008 GGU V81 FQ 15361 Factorisation du code en vue de la modification
}
Unit UTomPGHistoDetail ;

Interface

Uses StdCtrls,
     Controls,
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,hdb, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
     Fiche, 
     FichList, 
{$else}
     eFiche, 
     eFichList, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HMsgBox,
     UTOM,
     LookUp,
     SaisieList,
     uTableFiltre,
     UTob ;

Type
  TOM_PGHISTODETAIL = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    private
      InfosModif,InitSal,LaTablette,CodeTabLib,NiveauSaisie : String;
      LeSalarie,LeChamp,TypePop : String;
      InitDate : TDateTime;
      Procedure SaisieDonne;
      procedure SalarieElipsisClick(Sender : TObject);
//PT6      function GetWherePTE(CodeTable, NiveauPredefini, NiveauSaisie, EtabSalarie, ConvSalarie: String): String;
    end ;

Implementation

Uses
  PGTablesDyna; //PT6

procedure TOM_PGHISTODETAIL.OnNewRecord ;
var Q : TQuery;
    UnControl : TComponent;
begin
  Inherited ;
  SetField('PHD_DATEAPPLIC',V_PGI.DateEntree);
//PT3  SetField('PHD_DATEANNUL',IDate1900);  // : Le champs n'existe pas dans la table -> Plantage en CWAS
  If LeSalarie <> '' then
  begin
       If NiveauSaisie = 'SAL' then SetField('PHD_SALARIE',LeSalarie)
       else If NiveauSaisie = 'ETB' then SetField('PHD_ETABLISSEMENT',LeSalarie)
       else If NiveauSaisie = 'POP' then
       begin
          SetField('PHD_POPULATION',TypePop);
          SetField('PHD_CODEPOP',LeSalarie)
       end;
//PT5       SetField('PHD_PGINFOSMODIF',LeChamp);
       SetControlEnabled('PHD_SALARIE',False);
//PT5       SetControlEnabled('PHD_PGINFOSMODIF',False);
  end;
  SetField('PHD_PGINFOSMODIF',LeChamp); //PT5
  SetControlEnabled('PHD_PGINFOSMODIF',False); //PT5
  SetField('PHD_PGTYPEHISTO','003');
  Q := OpenSQL('SELECT * FROM PARAMSALARIE WHERE PPP_PGINFOSMODIF="'+LeChamp+'"',True);
  If Not Q.Eof then
  begin
    SetField('PHD_TABLETTE',Q.FindField('PPP_LIENASSOC').AsString);
    SetField('PHD_PGTYPEINFOLS',Q.FindField('PPP_PGTYPEINFOLS').AsString);
  end;
  Ferme(Q);
  UnControl := Ecran.FindComponent('PHD_NEWVALEUR');
  If unControl <> Nil then
  begin
    If unControl is TCheckbox then SetField('PHD_NEWVALEUR','-');
  end;
end ;

procedure TOM_PGHISTODETAIL.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_PGHISTODETAIL.OnUpdateRecord ;
begin
  Inherited ;
  If DS.State in [DsInsert] then
  begin
       SetField('PHD_GUIDHISTO',AglGetGuid());
  end;
  If not (Ecran is TFSaisieList) then
  begin
      If GetField('PHD_TYPEVALEUR') = 'B' then
      begin
           SetField('PHD_ANCVALEUR',GetControlText('ANCHECK'));
           SetField('PHD_NEWVALEUR',GetControlText('NEWCHECK'));
      end
      else
      begin
           SetField('PHD_ANCVALEUR',GetControlText('ANCEDIT'));
           SetField('PHD_NEWVALEUR',GetControlText('NEWEDIT'));
      end;
  end;
  SetField('PHD_CODTABL',CodeTabLib);
  if getField('PHD_DATEAPPLIC') <= V_PGI.DateEntree then SetField('PHD_TRAITEMENTOK','X');
  If ExisteSQL('SELECT PHD_NEWVALEUR FROM PGHISTODETAIL WHERE PHD_DATEAPPLIC="'+UsDateTime(GetField('PHD_DATEAPPLIC'))+'" '+
  'AND PHD_GUIDHISTO<>"'+GetField('PHD_GUIDHISTO')+'" AND PHD_SALARIE="'+GetField('PHD_SALARIE')+'" AND PHD_PGINFOSMODIF="'+GetField('PHD_PGINFOSMODIF')+'"') then
  begin
    LastError := 1;
    PGIBox('Il existe une valeur saisie pour cette date',Ecran.Caption);
    Exit;
  end;
end ;

procedure TOM_PGHISTODETAIL.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_PGHISTODETAIL.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

{//PT6
Function TOM_PGHISTODETAIL.GetWherePTE(CodeTable, NiveauPredefini, NiveauSaisie, EtabSalarie, ConvSalarie : String) : String;
begin
  result := ' WHERE PTE_CODTABL="'+CodeTable+'" AND PTE_DTVALID<="'+UsDateTime(GetField('PHD_DATEAPPLIC'))+'" ';
  if NiveauPredefini = 'CEG' then
    result := result + ' AND PTE_PREDEFINI="CEG" '
  else if NiveauPredefini = 'STD' then
    result := result + ' AND PTE_PREDEFINI="STD" '
  else if NiveauPredefini = 'DOS' then
    result := result + ' AND PTE_PREDEFINI="DOS" AND PTE_NODOSSIER="' + V_PGI.NoDossier + '" ';
  if NiveauSaisie = 'GEN' then
    result := result + ' AND PTE_NIVSAIS="GEN" '
  else if NiveauSaisie = 'ETB' then
    result := result + ' AND PTE_NIVSAIS="ETB" AND PTE_VALNIV="' + EtabSalarie + '" '
  else if NiveauSaisie = 'CON' then
    result := result + ' AND PTE_NIVSAIS="CON" AND PTE_VALNIV="' + ConvSalarie + '" ';
end;
}

procedure TOM_PGHISTODETAIL.OnLoadRecord ;
var //PT6  St,StWhere : String;
    TF: TTableFiltre;
//PT6    Q : TQuery; //PT2
//PT6    ConvSal,Etab : String; //PT2
//PT6    DateMax:TDateTime;
begin
  Inherited ;
  InfosModif := GetField('PHD_PGINFOSMODIF');
  InitSal := GetField('PHD_SALARIE');
  SaisieDonne;
  If LaTablette = 'PGCOMBOZONELIBRE' then
  begin
    InitDate := GetField('PHD_DATEAPPLIC');
    SetControlProperty('PHD_NEWVALEUR'
                      ,'Plus'
                      ,GetPlusPGCOMBOZONELIBRE(InitDate, CodeTabLib, InitSal)); //PT6
{ //PT6
    //DEB PT2
    Q := OpenSQL('SELECT PSA_CONVENTION,PSA_ETABLISSEMENT FROM SALARIES WHERE PSA_SALARIE = "' + InitSal + '"',True,1); //PT4
    if not Q.Eof then
    begin
      ConvSal := Q.FindField('PSA_CONVENTION').AsString;
      Etab := Q.FindField('PSA_ETABLISSEMENT').AsString;  //PT4
    end;
    Ferme(Q);
    //FIN PT2
    //DEB PT4
    //Vérifier s'il n'existe pas des valeurs pour l'établissement du salarié
    if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' + GetWherePTE(CodeTabLib, 'DOS', 'ETB', Etab, ConvSal)) then
    begin
      St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT'+ GetWherePTE(CodeTabLib, 'DOS', 'ETB', Etab, ConvSal);
      Q := OpenSQL(St,True,1);
      DateMax := iDate1900;
      if not Q.Eof then
        DateMax := Q.FindField('DTVALID').AsDateTime;
      Ferme(Q);
      StWhere := ' WHERE PTD_DTVALID = "' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+CodeTabLib+'"' +
        ' AND PTD_PREDEFINI="DOS" AND PTD_NODOSSIER="' + V_PGI.NoDossier + '"' +
        ' AND PTD_NIVSAIS="ETB" AND PTD_VALNIV="' + Etab + '"';
    end
    //Vérifier s'il n'existe pas des valeurs pour le dossier en général
    else if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' + GetWherePTE(CodeTabLib, 'DOS', 'GEN', Etab, ConvSal)) then
    begin
      St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT'+ GetWherePTE(CodeTabLib, 'DOS', 'GEN', Etab, ConvSal);
      Q := OpenSQL(St,True,1);
      DateMax := iDate1900;
      if not Q.Eof then
        DateMax := Q.FindField('DTVALID').AsDateTime;
      Ferme(Q);
      StWhere := ' WHERE PTD_DTVALID = "' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+CodeTabLib+'"' +
        ' AND PTD_PREDEFINI="DOS" AND PTD_NODOSSIER="' + V_PGI.NoDossier + '"' +
        ' AND PTD_NIVSAIS="GEN"';
    end
    //Vérifier s'il n'existe pas des valeurs pour STD + convention
    else if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' + GetWherePTE(CodeTabLib, 'STD', 'CON', Etab, ConvSal)) then
    begin
      St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT' + GetWherePTE(CodeTabLib, 'STD', 'CON', Etab, ConvSal);
      Q := OpenSQL(St,True,1);
      DateMax := iDate1900;
      if not Q.Eof then
        DateMax := Q.FindField('DTVALID').AsDateTime;
      Ferme(Q);
      StWhere := ' WHERE PTD_DTVALID = "' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+CodeTabLib+'"' +
        ' AND PTD_PREDEFINI="STD"' +
        ' AND PTD_NIVSAIS="CON" AND PTD_VALNIV="' + ConvSal + '"';
    end
    //Vérifier s'il n'existe pas des valeurs pour STD + Convention 000
    else if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' + GetWherePTE(CodeTabLib, 'STD', 'CON', Etab, '000')) then
    begin
      St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT' + GetWherePTE(CodeTabLib, 'STD', 'CON', Etab, '000');
      Q := OpenSQL(St,True,1);
      DateMax := iDate1900;
      if not Q.Eof then
        DateMax := Q.FindField('DTVALID').AsDateTime;
      Ferme(Q);
      StWhere := ' WHERE PTD_DTVALID = "' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+CodeTabLib+'"' +
        ' AND PTD_PREDEFINI="STD"' +
        ' AND PTD_NIVSAIS="CON" AND PTD_VALNIV="000"';
    end
    //Vérifier s'il n'existe pas des valeurs pour STD + GEN
    else if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' + GetWherePTE(CodeTabLib, 'STD', 'GEN', Etab, ConvSal)) then
    begin
      St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT' + GetWherePTE(CodeTabLib, 'STD', 'GEN', Etab, ConvSal);
      Q := OpenSQL(St,True,1);
      DateMax := iDate1900;
      if not Q.Eof then
        DateMax := Q.FindField('DTVALID').AsDateTime;
      Ferme(Q);
      StWhere := ' WHERE PTD_DTVALID = "' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+CodeTabLib+'"' +
        ' AND PTD_PREDEFINI="STD" AND PTD_NIVSAIS="GEN"';
    end
    //Vérifier s'il n'existe pas des valeurs pour CEG
    else
    begin
      St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT' + GetWherePTE(CodeTabLib, 'CEG', 'GEN', Etab, ConvSal);
      Q := OpenSQL(St,True,1);
      DateMax := iDate1900;
      if not Q.Eof then
        DateMax := Q.FindField('DTVALID').AsDateTime;
      Ferme(Q);
      StWhere := ' WHERE PTD_DTVALID = "' + USDATETIME(DateMax) + '" AND PTD_CODTABL="'+CodeTabLib+'"' +
        ' AND PTD_PREDEFINI="CEG" AND PTD_NIVSAIS="GEN"';
    end;
//      St := ' WHERE PTD_DTVALID IN (SELECT MAX(PTE_DTVALID) FROM TABLEDIMENT WHERE PTD_CODTABL=PTE_CODTABL '+
//        'AND PTE_DTVALID<="'+UsDateTime(GetField('PHD_DATEAPPLIC'))+'") AND PTD_CODTABL="'+CodeTabLib+'"' +
//        ' AND (PTD_NIVSAIS = "GEN" OR (PTD_NIVSAIS ="CON" AND PTD_VALNIV="' +ConvSal+ '")' +  //PT2
//        ' OR (PTD_NIVSAIS ="CON" AND PTD_VALNIV="000"))';  //PT2
    //FIN PT4
    SetControlProperty('PHD_NEWVALEUR','Plus',StWhere );
}
   end;
   if (Ecran <> nil) and (Ecran is TFSaisieList) then
   begin
     TF := TFSaisieList(Ecran).LeFiltre;
     If TF.DataSet.State = DsInsert then SetControlEnabled('PHD_DATEAPPLIC',True)
    else SetControlEnabled('PHD_DATEAPPLIC',False);
   end
   else
   begin
    If TFFiche(Ecran).FTypeAction = TaCreat then SetControlEnabled('PHD_DATEAPPLIC',True)
    else SetControlEnabled('PHD_DATEAPPLIC',False);
    TFFiche(Ecran).Caption := 'Historique '+RechDom('PGZONEHISTOSALGEREE',GetField('PHD_PGINFOSMODIF'),False)+ ' du salarié : '+RechDom('PGSALARIE',getField('PHD_SALARIE'),False);
    UpdateCaption(TFFiche(Ecran));
   end;
end ;

procedure TOM_PGHISTODETAIL.OnChangeField ( F: TField ) ;
//PT6 var St : String;
//PT6  Q : TQuery; //PT2
//PT6  ConvSal,Etab : String; //PT2
begin
  Inherited ;
      If F.FieldName = 'PHD_PGINFOSMODIF' then
      begin
           if InfosModif <> GetField('PHD_PGINFOSMODIF') then SaisieDonne;
           InfosModif := GetField('PHD_PGINFOSMODIF');
      end;
      If F.FieldName = 'PHD_SALARIE' then
      begin
           if InfosModif <> GetField('PHD_PGINFOSMODIF') then SaisieDonne;
           InfosModif := GetField('PHD_PGINFOSMODIF');
      end;
  If LaTablette = 'PGCOMBOZONELIBRE' then
  begin
    If GetField('PHD_DATEAPPLIC') <> InitDate then
    begin
      InitDate := GetField('PHD_DATEAPPLIC');
      SetControlProperty('PHD_NEWVALEUR'
                        ,'Plus'
                        ,GetPlusPGCOMBOZONELIBRE(InitDate, CodeTabLib, GetField('PHD_SALARIE'))); //PT6

{//PT6
      //DEB PT2
      Q := OpenSQL('SELECT PSA_CONVENTION,PSA_ETABLISSEMENT FROM SALARIES WHERE PSA_SALARIE = "' + GetField('PHD_SALARIE') + '"',True,1);
      if not Q.Eof then
      begin
        ConvSal := Q.FindField('PSA_CONVENTION').AsString;
        Etab := Q.FindField('PSA_ETABLISSEMENT').AsString; //PT4
      end;
      Ferme(Q);
      //FIN PT2
      //DEB PT4
      //Vérifier s'il n'existe pas des valeurs pour l'établissement du salarié
      if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' + GetWherePTE(CodeTabLib, 'DOS', 'ETB', Etab, ConvSal)) then
        St := ' WHERE PTD_DTVALID IN (SELECT MAX(PTE_DTVALID) FROM TABLEDIMENT '
            + GetWherePTE(CodeTabLib, 'DOS', 'ETB', Etab, ConvSal)
            + ') AND PTD_CODTABL="'+CodeTabLib+'"' +
          ' AND PTD_PREDEFINI="DOS" AND PTD_NODOSSIER="' + V_PGI.NoDossier + '"' +
          ' AND PTD_NIVSAIS="ETB" AND PTD_VALNIV="' + Etab + '"'
      //Vérifier s'il n'existe pas des valeurs pour le dossier en général
      else if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' + GetWherePTE(CodeTabLib, 'DOS', 'GEN', Etab, ConvSal)) then
        St := ' WHERE PTD_DTVALID IN (SELECT MAX(PTE_DTVALID) FROM TABLEDIMENT '
            + GetWherePTE(CodeTabLib, 'DOS', 'GEN', Etab, ConvSal)
            + ') AND PTD_CODTABL="'+CodeTabLib+'"' +
          ' AND PTD_PREDEFINI="DOS" AND PTD_NODOSSIER="' + V_PGI.NoDossier + '"' +
          ' AND PTD_NIVSAIS="GEN"'
      //Vérifier s'il n'existe pas des valeurs pour STD + convention
      else if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' + GetWherePTE(CodeTabLib, 'STD', 'CON', Etab, ConvSal)) then
        St := ' WHERE PTD_DTVALID IN (SELECT MAX(PTE_DTVALID) FROM TABLEDIMENT '
            + GetWherePTE(CodeTabLib, 'STD', 'CON', Etab, ConvSal)
            + ') AND PTD_CODTABL="'+CodeTabLib+'"' +
          ' AND PTD_PREDEFINI="STD"' +
          ' AND PTD_NIVSAIS="CON" AND PTD_VALNIV="' + ConvSal + '"'
      //Vérifier s'il n'existe pas des valeurs pour STD + Convention 000
      else if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' + GetWherePTE(CodeTabLib, 'STD', 'CON', Etab, '000')) then
        St := ' WHERE PTD_DTVALID IN (SELECT MAX(PTE_DTVALID) FROM TABLEDIMENT '
            + GetWherePTE(CodeTabLib, 'STD', 'CON', Etab, '000')
            + ') AND PTD_CODTABL="'+CodeTabLib+'"' +
          ' AND PTD_PREDEFINI="STD"' +
          ' AND PTD_NIVSAIS="CON" AND PTD_VALNIV="000"'
      //Vérifier s'il n'existe pas des valeurs pour STD + GEN
      else if ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT' + GetWherePTE(CodeTabLib, 'STD', 'GEN', Etab, ConvSal)) then
        St := ' WHERE PTD_DTVALID IN (SELECT MAX(PTE_DTVALID) FROM TABLEDIMENT '
            + GetWherePTE(CodeTabLib, 'STD', 'GEN', Etab, ConvSal)
            + ') AND PTD_CODTABL="'+CodeTabLib+'"' +
          ' AND PTD_PREDEFINI="STD" AND PTD_NIVSAIS="GEN"'
      //Vérifier s'il n'existe pas des valeurs pour CEG
      else
        St := ' WHERE PTD_DTVALID IN (SELECT MAX(PTE_DTVALID) FROM TABLEDIMENT '
            + GetWherePTE(CodeTabLib, 'CEG', 'GEN', Etab, ConvSal)
            + ') AND PTD_CODTABL="'+CodeTabLib+'"' +
          ' AND PTD_PREDEFINI="CEG" AND PTD_NIVSAIS="GEN"';
      //FIN PT4
      SetControlProperty('PHD_NEWVALEUR','Plus',St);
}
    end;
  end;
end ;

procedure TOM_PGHISTODETAIL.OnArgument ( S: String ) ;
var Q : TQuery;
  {$IFNDEF EAGLCLIENT}
  sal: THDBEdit;
  {$ELSE}
  sal: THEdit;
  {$ENDIF}
begin
  Inherited ;
  NiveauSaisie := ReadTokenPipe(S,';');
  LeSalarie := ReadTokenPipe(S,';');
  LeChamp := ReadTokenPipe(S,';');
  if LeChamp = 'PSA_SITUATIONFAMIL' then LeChamp := 'PSA_SITUATIONFAMI';
  LaTablette := ReadTokenPipe(S,';');
  TypePop := ReadTokenPipe(S,';');
  CodeTabLib := '';
  If LaTablette = 'PGCOMBOZONELIBRE' then
  begin
    Q := OpenSQL('SELECT PPP_CODTABL FROM PARAMSALARIE WHERE PPP_PGINFOSMODIF="'+LeChamp+'"',True);
    If Not Q.Eof then
    begin
     CodeTabLib := Q.FindField('PPP_CODTABL').AsString;
    end;
    Ferme(Q);
  end;
  //DEBUT PT1
 {$IFNDEF EAGLCLIENT}
  sal := THDBEdit(GetControl('PHD_SALARIE'));
  {$ELSE}
  sal := THEdit(GetControl('PHD_SALARIE'));
  {$ENDIF}
 if sal <> nil then Sal.OnElipsisClick := SalarieElipsisClick;
 //FIN PT1
end ;

procedure TOM_PGHISTODETAIL.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_PGHISTODETAIL.OnCancelRecord ;
begin
  Inherited ;
end ;

Procedure TOM_PGHISTODETAIL.SaisieDonne;
var Q : TQuery;
    LeType,Tablette,Champ : String;
begin
  If GetField('PHD_SALARIE') = '' then Exit;
  If GetField('PHD_PGINFOSMODIF') = '' then Exit;
  If DS.State in [DSInsert] then
  begin
       Champ := GetField('PHD_PGINFOSMODIF');
       Q := OpenSQL('SELECT PPP_PGTYPEDONNE,PPP_LIENASSOC FROM PARAMSALARIE WHERE PPP_PGINFOSMODIF="'+GetField('PHD_PGINFOSMODIF')+'"',True);
       If Not Q.Eof then
       begin
            LeType := Q.FindField('PPP_PGTYPEDONNE').AsString;
            Tablette := Q.FindField('PPP_LIENASSOC').AsString;
       end;
       Ferme(Q);
       SetField('PHD_TYPEVALEUR',LeType);
       SetField('PHD_TABLETTE',Tablette);
       If LeType = 'SAL' then
       begin
         Q := OpenSQL('SELECT '+GetField('PHD_PGINFOSMODIF')+' FROM SALARIES WHERE PSA_SALARIE="'+Getfield('PHD_SALARIE')+'"',True);
         If Not Q.Eof then
         begin
              If (LeType = 'B') or (LeType='S') or (LeType='T') then
              begin
                   Setfield('PHD_ANCVALEUR',Q.FindField(Champ).AsString);
              end
              else If LeType = 'D' then
              begin
                   Setfield('PHD_ANCVALEUR',DateToStr(Q.FindField(Champ).AsDateTime));
              end
              else if LeType = 'I' then
              begin
                   Setfield('PHD_ANCVALEUR',IntToStr(Q.FindField(Champ).AsInteger));
              end
              else If LeType = 'F' then
              begin
                   Setfield('PHD_ANCVALEUR',FloatToStr(Q.FindField(Champ).AsFloat));
              end;
         end;
       end;
       Ferme(Q);
  end;
  If (Ecran is TFSaisieList) then Exit;
  If GetField('PHD_TYPEVALEUR') = 'B' then
  begin
       SetControlText('NEWCHECK',GetField('PHD_NEWVALEUR'));
       SetControlText('ANCCHECK',GetField('PHD_ANCVALEUR'));
  end
  else
  begin
       SetControlText('NEWEDIT',GetField('PHD_NEWVALEUR'));
       SetControlText('ANCEDIT',GetField('PHD_ANCVALEUR'));
  end;
  SetControlVisible('LIBANC',False);
  SetControlVisible('LIBNEW',False);
  SetControlProperty('ANCEDIT','DataType','');
  SetControlProperty('NEWEDIT','DataType','');
  SetControlProperty('ANCEDIT','ElipsisButton',False);
  SetControlProperty('NEWEDIT','ElipsisButton',False);
  SetControlProperty('ANCEDIT','EditMask','');
  SetControlProperty('NEWEDIT','EditMask','');
  If GetField('PHD_TYPEVALEUR') = 'B' then
  begin
       SetControlText('NEWCHECK','-');
       SetControlVisible('ANCCHECK',True);
       SetControlVisible('NEWCHECK',True);
       SetControlVisible('ANCEDIT',False);
       SetControlVisible('NEWEDIT',False);
       SetControlVisible('TANCEDIT',False);
       SetControlVisible('TNEWEDIT',False);
       SetControlText('ANCCHECK',GetField('PHD_ANCVALEUR'));
       SetControlText('NEWCHECK',GetField('PHD_NEWVALEUR'));
  end
  else
  begin
       SetControlVisible('ANCCHECK',False);
       SetControlVisible('NEWCHECK',False);
       SetControlVisible('ANCEDIT',True);
       SetControlVisible('NEWEDIT',True);
       SetControlVisible('TANCEDIT',True);
       SetControlVisible('TNEWEDIT',True);
       If GetField('PHD_TYPEVALEUR') = 'S' then
       begin
            SetControlText('NEWEDIT','');
            SetControlText('ANCEDIT',GetField('PHD_ANCVALEUR'));
            SetControlText('NEWEDIT',GetField('PHD_NEWVALEUR'));
            SetControlProperty('ANCEDIT','OpeType',OtString);
            SetControlProperty('NEWEDIT','OpeType',OtString);
       end
       else If GetField('PHD_TYPEVALEUR') = 'F' then
       begin
            SetControlText('NEWEDIT','0');
            SetControlText('ANCEDIT',GetField('PHD_ANCVALEUR'));
            SetControlText('NEWEDIT',GetField('PHD_NEWVALEUR'));
            SetControlProperty('ANCEDIT','OpeType',OtReel);
            SetControlProperty('NEWEDIT','OpeType',OtReel);
       end
       else If GetField('PHD_TYPEVALEUR') = 'I' then
       begin
            SetControlText('NEWEDIT','0');
            SetControlText('ANCEDIT',GetField('PHD_ANCVALEUR'));
            SetControlText('NEWEDIT',GetField('PHD_NEWVALEUR'));
            SetControlProperty('ANCEDIT','OpeType',OtReel);
            SetControlProperty('NEWEDIT','OpeType',OtReel);
       end
       else If GetField('PHD_TYPEVALEUR') = 'D' then
       begin
            SetControlProperty('ANCEDIT','EditMask','!99/99/0000;1;_');
            SetControlProperty('NEWEDIT','EditMask','!99/99/0000;1;_');
            SetControlText('NEWEDIT',DateToStr(V_PGI.DateEntree));
            SetControlText('ANCEDIT',GetField('PHD_ANCVALEUR'));
            SetControlText('NEWEDIT',GetField('PHD_NEWVALEUR'));
            SetControlProperty('ANCEDIT','ElipsisButton',True);
            SetControlProperty('NEWEDIT','ElipsisButton',True);
            SetControlProperty('ANCEDIT','OpeType',OtDate);
            SetControlProperty('NEWEDIT','OpeType',OtDate);
       end
       else If GetField('PHD_TYPEVALEUR') = 'T' then
       begin

            SetControlText('ANCEDIT',GetField('PHD_ANCVALEUR'));
            SetControlText('NEWEDIT',GetField('PHD_NEWVALEUR'));
            SetControlCaption('LIBNEW',RechDom(GetField('PHD_TABLETTE'),GetControltext('NEWEDIT'),False));
            SetControlCaption('LIBANC',RechDom(GetField('PHD_TABLETTE'),GetControltext('ANCEDIT'),False));
            SetControlVisible('LIBANC',True);
            SetControlVisible('LIBNEW',True);
            SetControlProperty('ANCEDIT','ElipsisButton',True);
            SetControlProperty('NEWEDIT','ElipsisButton',True);
            SetControlProperty('ANCEDIT','DataType',GetField('PHD_TABLETTE'));
            SetControlProperty('NEWEDIT','DataType',GetField('PHD_TABLETTE'));
            SetControlProperty('ANCEDIT','OpeType',OtString);
            SetControlProperty('NEWEDIT','OpeType',OtString);
       end
       else If GetField('PHD_TYPEVALEUR') = 'C' then
       begin

            SetControlText('ANCEDIT',GetField('PHD_ANCVALEUR'));
            SetControlText('NEWEDIT',GetField('PHD_NEWVALEUR'));
            SetControlCaption('LIBNEW',RechDom(GetField('PHD_TABLETTE'),GetControltext('NEWEDIT'),False));
            SetControlCaption('LIBANC',RechDom(GetField('PHD_TABLETTE'),GetControltext('ANCEDIT'),False));
            SetControlVisible('LIBANC',True);
            SetControlVisible('LIBNEW',True);
            SetControlProperty('ANCEDIT','DataType',GetField('PHD_TABLETTE'));
            SetControlProperty('NEWEDIT','DataType',GetField('PHD_TABLETTE'));
       end;
  end;
end;

procedure TOM_PGHISTODETAIL.SalarieElipsisClick(Sender : TObject); // PT1
begin
  {$IFNDEF EAGLCLIENT}
  LookupList(THDBEdit(Sender), 'Salarié', 'SALARIES', 'PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', '', 'PSA_SALARIE', True, -1);
  {$ELSE}
  LookupList(THEdit(Sender), 'Salarié', 'SALARIES', 'PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', '', 'PSA_SALARIE', True, -1);
  {$ENDIF}
end;


Initialization
  registerclasses ( [ TOM_PGHISTODETAIL ] ) ;
end.

