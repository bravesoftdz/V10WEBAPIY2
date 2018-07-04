{***********UNITE*************************************************
Auteur  ...... :  JL
Créé le ...... : 21/12/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGAUGMENTATIONSUIVI ()
Mots clefs ... : TOF;PGAUGMENTATIONSUIVI
*****************************************************************}
Unit UTofPGAugmentationSuivi ;

Interface

Uses StdCtrls,
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul,
{$ENDIF}
     forms,
     uTob,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HSysMenu,
     EntPaie,
     ParamSoc,
     Vierge,
     Grids,
     ed_tools,
     graphics,
     UTOF ;

Type
  TOF_PGAUGMENTATIONSUIVI = Class (TOF)

    procedure OnArgument (S : String ) ; override ;
    private
    procedure AfficheHistoSal(Salarie : String);
    Function ChampSalaires(TypeSalaire : String) : String;
    Function WhereSalaires : String;
    procedure GrilleGetCellCanvas( ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
  end ;

Implementation

procedure TOF_PGAUGMENTATIONSUIVI.OnArgument (S : String ) ;
Var LeSalarie : String;
    Grille : THGrid;
begin
  Inherited ;
//        TFVierge(Ecran).Height := 486;
//        TFVierge(Ecran).Width := 593;
        LeSalarie := ReadTokenPipe(S,';');
        Ecran.Caption := 'Historique des augmentations du salarié : '+RechDom('PGSALARIE',LeSalarie,False);
        UpdateCaption(Ecran);
        Grille := THGrid(GetControl('GHISTOSAL'));
        If grille <> Nil then Grille.GetCellCanvas := GrilleGetCellCanvas;
        AfficheHistoSal(LeSalarie);
end ;

procedure TOF_PGAUGMENTATIONSUIVI.AfficheHistoSal(Salarie : String);
var Q : TQuery;
    TobGrille : Tob;
    Grille : THGrid;
    ChampFixe,ChampVar,LesChamps : String;
    WhereSalaire : String;
    i : Integer;
    FixePrec,VarPrec,BrutPrec : Double;
    Brut,Variable,Fixe : Double;
    PctV,PctF,PctB : Double;
    PctAugmDec : Integer;
    Titres : HTStringList;
    DatePrec : TDateTime;
    CalculerPct : Boolean;
    BrutTpsPlein,UniteEff,Coeff : Double;
begin
        UniteEff := 1;
        Q := OpenSQL('SELECT PSA_UNITEPRISEFF FROM SALARIES WHERE PSA_SALARIE="'+Salarie+'"',true);
        If Not Q.Eof then UniteEff := Q.findField('PSA_UNITEPRISEFF').AsFloat;
        Ferme(Q);
        Coeff := 1 / UniteEff;
        PctAugmDec := GetParamSoc('SO_PGAUGMPCTDEC');
        ChampFixe := ChampSalaires('FIXE');
        ChampVar := ChampSalaires('VARIABLE');
        WhereSalaire := WhereSalaires;
        If ChampFixe <> '' then ChampFixe := '('+ChampFixe+') LEFIXE';
        If ChampVar <> '' then ChampVar := '('+ChampVar+') LEVARIABLE';
        LesChamps := '';
        If ChampFixe <> '' then LesChamps := ','+ChampFixe;
        If ChampVar <> '' then LesChamps := LesChamps+','+ChampVar;
        FixePrec := 0;
        VarPrec := 0;
        BrutPrec := 0;
        Q := OpenSQL('SELECT PHS_DATEAPPLIC'+LesChamps+' FROM HISTOSALARIE WHERE '+
        'PHS_SALARIE="'+Salarie+'"'+WhereSalaire+' ORDER BY PHS_DATEAPPLIC',True);
        TobGrille := Tob.Create('LesAug',Nil,-1);
        TobGrille.LoadDetailDB('LesAug','','',Q,False);
        Ferme(Q);
        DatePrec := IDate1900;
        For i := 0 to TobGrille.Detail.Count - 1 do
        begin
             If DatePrec > StrToDate('01/01/2002') then  CalculerPct := True
             else CalculerPct := False;
             If not TobGrille.Detail[i].FieldExists('LEFIXE') then TobGrille.Detail[i].AddChampSupValeur('LEFIXE',0);
             If not TobGrille.Detail[i].FieldExists('LEVARIABLE') then TobGrille.Detail[i].AddChampSupValeur('LEVARIABLE',0);
             TobGrille.Detail[i].AddChampSupValeur('LETOTAL',TobGrille.Detail[i].GetDouble('LEFIXE') + TobGrille.Detail[i].GetDouble('LEVARIABLE'));
             Brut := TobGrille.Detail[i].getValue('LETOTAL');
             BrutTpsPlein := Brut * Coeff;
             Variable := TobGrille.Detail[i].getValue('LEVARIABLE');
             Fixe := TobGrille.Detail[i].getValue('LEFIXE');
             TobGrille.Detail[i].AddChampSupValeur('TPSPLEIN',BrutTpsPlein);
             If not CalculerPct then
             begin
                  TobGrille.Detail[i].AddChampSupValeur('PCTBRUT','NS');
             end
             else
             begin
                  If BrutPrec <> 0 then PctB := Arrondi(((BrutTpsPlein - BrutPrec)/BrutPrec)*100,PctAugmDec)
                  else PctB := 0;
                  TobGrille.Detail[i].AddChampSupValeur('PCTBRUT',FormatFloat('# ##0.00',PctB));
             end;
             FixePrec := Fixe;
             VarPrec := Variable;
             BrutPrec := BrutTpsPlein;
             DatePrec := TobGrille.Detail[i].getValue('PHS_DATEAPPLIC');
        end;
        Grille := THGrid(GetControl('GHISTOSAL'));
        Grille.ColCount := 7;
        Titres := HTStringList.Create;
        Titres.Insert(0, 'Date d''application');
        Titres.Insert(1, 'Salaire fixe');
        Titres.Insert(2, 'Salaire Variable');
        Titres.Insert(3, 'Brut');
        Titres.Insert(4, 'Brut tps plein');
        Titres.Insert(5, '%');
        Grille.Titres := Titres;
        Titres.free;
        Grille.ColFormats[0] := ShortDateFormat;
        Grille.ColFormats[1] := '# ##0.00';
        Grille.ColFormats[2] := '# ##0.00';
        Grille.ColFormats[3] := '# ##0.00';
        Grille.ColFormats[4] := '# ##0.00';
        Grille.ColFormats[5] := '# ##0.00';
        Grille.ColAligns[0] := taCenter;
        Grille.ColAligns[1] := taRightJustify;
        Grille.ColAligns[2] := taRightJustify;
        Grille.ColAligns[3] := taRightJustify;
        Grille.ColAligns[4] := taRightJustify;
        Grille.ColAligns[5] := taRightJustify;
        Grille.ColAligns[6] := taRightJustify;
        Grille.RowCount := TobGrille.Detail.Count + 1;
        TobGrille.PutGridDetail(Grille,False,False,'PHS_DATEAPPLIC;LEFIXE;LEVARIABLE;LETOTAL;TPSPLEIN;PCTBRUT',False);
        TobGrille.Free;
//        HMTrad.ResizeGridColumns(Grille) ;
//        Grille.ColWidths[0] := 110;
//        Grille.ColWidths[1] := 95;
//        Grille.ColWidths[2] := 95;
//        Grille.ColWidths[3] := 95;
        Grille.ColWidths[4] := -1;
//        Grille.ColWidths[5] := 50;
        Grille.ColWidths[6] := 1;

end;

Function TOF_PGAUGMENTATIONSUIVI.ChampSalaires(TypeSalaire : String) : String;
var LesSalaires,TempSalaire,Champ,Requete : String;
    i : Integer;
begin
     LesSalaires := '';
     If TypeSalaire = 'FIXE' then LesSalaires := GetParamSoc('SO_PGAUGFIXE');
     If TypeSalaire = 'VARIABLE' then LesSalaires := GetParamSoc('SO_PGAUGVARIABLE');
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
               if TempSalaire = 'AN'+IntToStr(i) then Champ := 'PHS_SALAIREANN'+IntToStr(i);
          end;
          For i := 1 to 5 do
          begin
               If TempSalaire = 'MO'+IntToStr(i) then Champ := 'PHS_SALAIREMOIS'+IntToStr(i);
          end;
          If Champ <> '' then
          begin
               If Requete <> '' then Requete := Requete + ' + '+Champ
               else Requete := Champ;
          end;
     end;
     Result := Requete;
end;

Function TOF_PGAUGMENTATIONSUIVI.WhereSalaires : String;
var LesSalairesF,LesSalairesV,TempSalaire,Champ,Requete : String;
    i : Integer;
begin
     LesSalairesF := GetParamSoc('SO_PGAUGFIXE');
     LesSalairesV := GetParamSoc('SO_PGAUGVARIABLE');
     If (LesSalairesF = '') and (LesSalairesV = '')  then
     begin
          Result := '';
          Exit;
     end;
     Requete := '';
     While LesSalairesF <> '' do
     begin
          Champ := '';
          TempSalaire := ReadTokenPipe(LesSalairesF,';');
          For i := 1 to 5 do
          begin
               if TempSalaire = 'AN'+IntToStr(i) then Champ := 'PHS_BSALAIREANN'+IntToStr(i);
          end;
          For i := 1 to 5 do
          begin
               If TempSalaire = 'MO'+IntToStr(i) then Champ := 'PHS_BSALAIREMOIS'+IntToStr(i);
          end;
          If Champ <> '' then
          begin
               If Requete <> '' then Requete := Requete + ' OR '+Champ+'="X"'
               else Requete := ' AND ('+Champ+'="X"';
          end;
     end;
     While LesSalairesV <> '' do
     begin
          Champ := '';
          TempSalaire := ReadTokenPipe(LesSalairesV,';');
          For i := 1 to 5 do
          begin
               if TempSalaire = 'AN'+IntToStr(i) then Champ := 'PHS_BSALAIREANN'+IntToStr(i);
          end;
          For i := 1 to 5 do
          begin
               If TempSalaire = 'MO'+IntToStr(i) then Champ := 'PHS_BSALAIREMOIS'+IntToStr(i);
          end;
          If Champ <> '' then
          begin
               If Requete <> '' then Requete := Requete + ' OR '+Champ+'="X"'
               else Requete := ' AND ('+Champ+'="X"';
          end;
     end;
     If Requete <> '' then Requete := Requete+')';
     Result := Requete;
end;

procedure TOF_PGAUGMENTATIONSUIVI.GrilleGetCellCanvas( ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
begin
      If ARow = 0 then Exit;
      If (ACol = 5) then Canvas.Font.Style := [fsBold];

end;

Initialization
  registerclasses ( [ TOF_PGAUGMENTATIONSUIVI ] ) ;
end.

