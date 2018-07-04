{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 01/03/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PARAMHISTOSTDDOS ()
Mots clefs ... : TOF;PARAMHISTOSTDDOS
*****************************************************************
PT1    02/10/2007 FC V_80 : Concept plan de paie
}
Unit UTofPGParamHistoStdDos ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,Fe_Main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,   MainEAGL,
{$ENDIF}
     uTob,
     forms,
     HTB97,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HSysMenu,
     graphics,
     Grids,
     EntPaie,
     Vierge,
     UTOF ;

Type
  TOF_PARAMHISTOSTDDOS = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnClose                  ; override ;
    private
    Grille : THGrid;
    TobGrille : Tob;
    ProfilParDefaut : Boolean;
    CEG, STD, DOS: Boolean;
    procedure AfficheGrille;
    procedure ClickGrille(Sender : TObject);
    procedure GrilleGetCellCanvas( ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
    Function ValiderSaisie : Boolean;
    Function RendLibelle(LeChamp : String) : String;
    procedure BAjoutClick(Sender: TObject);
    procedure BMoinsClick(Sender: TObject);
    procedure ChangeValueSelection(MyRect : TGridRect ; Value : string);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure MajProfil(TobParam : Tob;ProfilStd,ProfilDos : Boolean);
    Function VerifierLesDonnes(TobModif,TobInit : Tob) : Boolean;
  end ;

Implementation
uses
  pgOutils;

procedure TOF_PARAMHISTOSTDDOS.OnClose ;
begin
  Inherited ;
  If TobGrille <> Nil then FreeAndNil(TobGrille);
end;

procedure TOF_PARAMHISTOSTDDOS.OnUpdate ;
Var AjoutChamp : Boolean;
begin
  Inherited ;
  AjoutChamp := ValiderSaisie;
  If AjoutChamp then AglLanceFiche('PAY', 'MUL_INITHISTO', '', '', 'PARAMSOC');
end;

procedure TOF_PARAMHISTOSTDDOS.OnArgument (S : String ) ;
var Bt : TToolBarButton97;
begin
  Inherited ;
  If ReadTokenPipe(S,';') = 'X' then ProfilParDefaut := True
  else ProfilParDefaut := False;
  Grille := THGrid(GetControl('GRILLE'));
  AfficheGrille;
  Grille.GetCellCanvas := GrilleGetCellCanvas;
  Grille.OnDblClick := ClickGrille;
  TFVierge(Ecran).OnKeyDown := FormKeyDown;
  Bt := TToolBarButton97(GetControl('BPLUS'));
  If Bt <> Nil then Bt.OnClick := BAjoutClick;
  Bt := TToolBarButton97(GetControl('BMOINS'));
  If Bt <> Nil then Bt.OnClick := BMoinsClick;
  //DEB PT1
  AccesPredefini('TOUS', CEG, STD, DOS);
{$IFDEF CPS1}
   STD := TRUE;
{$ENDIF}
  //FIN PT1
end ;

procedure TOF_PARAMHISTOSTDDOS.AfficheGrille;
var Q : TQuery;
    TobChamp,T : Tob;
    i : Integer;
    Histo,Predef,Theme,ThemePrec,Libelle,Champ : String;
    HMTrad: THSystemMenu;
begin
  If TobGrille <> Nil then FreeAndNil(TobGrille);
  Grille.ColCOunt := 5;
  Grille.FixedCols := 1;
  Grille.ColDrawingModes[2]:= 'IMAGE';
  Grille.ColEditables[2] := False;
  Grille.ColDrawingModes[1]:= 'IMAGE';
  Grille.ColEditables[1] := False;
  Grille.ColAligns[0] := TaLeftJustify;
  Grille.ColWidths[3] := -1;
  Grille.ColWidths[4] := -1;
  Q:= OpenSQL('SELECT * FROM PARAMSALARIE WHERE PPP_PGTHEMESALARIE<>"PRO" AND PPP_PREDEFINI="CEG" AND PPP_HISTORIQUE="-" ORDER BY PPP_PGTHEMESALARIE',True);
  TobChamp := Tob.Create('LesChamps',Nil,-1);
  TobChamp.LoadDetailDB('LesChamps','','',Q,False);
  Ferme(Q);
  Theme := '';
  ThemePrec := '';
  TobGrille := Tob.Create('RemplirGrille',Nil,-1);
  For i := 0 to TobChamp.Detail.Count - 1 do
  begin
    //Affichage du thème
    Theme := TobChamp.Detail[i].GetValue('PPP_PGTHEMESALARIE');
    If (Theme='') and (i = 0) then
    begin
        T := Tob.Create('ligne grille',TobGrille,-1);
        T.AddChampSupValeur('CHAMP','Aucun thème');
        T.AddChampSupValeur('HISTOSTD','#ICO#00');
        T.AddChampSupValeur('HISTODOS','#ICO#00');
        T.AddChampSupValeur('CODE','THEME');
        T.AddChampSupValeur('CODETHEME','');
    end
    else If (Theme <> '') and (Theme <> ThemePrec) then
    begin
        T := Tob.Create('ligne grille',TobGrille,-1);
        T.AddChampSupValeur('CHAMP',RechDom('PGTHEMESAL',Theme,False));
        T.AddChampSupValeur('HISTOSTD','#ICO#00');
        T.AddChampSupValeur('HISTODOS','#ICO#00');
        T.AddChampSupValeur('CODE','THEME');
        T.AddChampSupValeur('CODETHEME',Theme);
    end;
    ThemePrec := Theme;
    //Recherche des standards
    Q := OpenSQL('SELECT PPP_HISTORIQUE FROM PARAMSALARIE WHERE PPP_PREDEFINI="STD" AND PPP_PGINFOSMODIF="'+TobChamp.Detail[i].GetValue('PPP_PGINFOSMODIF')+'"',True);
    If Not Q.Eof then Histo := Q.FindField('PPP_HISTORIQUE').AsString
    else Histo := '-';
    Ferme(Q);

    Libelle := TobChamp.Detail[i].GetValue('PPP_LIBELLE');
    Champ := TobChamp.Detail[i].GetValue('PPP_PGINFOSMODIF');
    If (Copy(Champ,1,12) = 'PSA_TRAVAILN') or (Champ = 'PSA_CODESTAT') or (Copy(Champ,1,13) = 'PSA_DATELIBRE')
    or (Copy(Champ,1,9) = 'PSA_LIBRE') or (Copy(Champ,1,13) = 'PSA_SALAIRANN') or (Copy(Champ,1,15) = 'PSA_SALAIREMOIS') then
    Libelle := RendLibelle(Champ);
    If Libelle = '' then Continue;
    T := Tob.Create('Ligne grille',TobGrille,-1);
    T.AddChampSupValeur('CHAMP',Libelle);
    If Histo = 'X' then T.AddChampSupValeur('HISTOSTD','#ICO#47')
    Else T.AddChampSupValeur('HISTOSTD','#ICO#46');
    If Histo = 'X' then T.AddChampSupValeur('HISTODOS','#ICO#47')  //Si histo standard, histo obligatoire pour dossier
    Else
    begin
       //Recherche des dossiers
      Q := OpenSQL('SELECT PPP_HISTORIQUE FROM PARAMSALARIE WHERE PPP_PREDEFINI="DOS" '+
      'AND PPP_NODOSSIER="'+V_PGI.NoDossier+'" AND PPP_PGINFOSMODIF="'+TobChamp.Detail[i].GetValue('PPP_PGINFOSMODIF')+'"',True);
      If Not Q.Eof then Histo := Q.FindField('PPP_HISTORIQUE').AsString
      else Histo := '-';
      If Histo = 'X' then T.AddChampSupValeur('HISTODOS','#ICO#47')
      Else T.AddChampSupValeur('HISTODOS','#ICO#46');
      Ferme(Q);
    end;
    T.AddChampSupValeur('CODE',Champ);
    T.AddChampSupValeur('CODETHEME',Theme);
  end;
  //Gestion profils
  Theme := 'PRO';
  T := Tob.Create('ligne grille',TobGrille,-1);
  T.AddChampSupValeur('CHAMP',RechDom('PGTHEMESAL',Theme,False));
  T.AddChampSupValeur('HISTOSTD','#ICO#00');
  T.AddChampSupValeur('HISTODOS','#ICO#00');
  T.AddChampSupValeur('CODE','THEME');
  T.AddChampSupValeur('CODETHEME',Theme);
  Libelle := 'Gestion des profils';
  Champ := 'LESPROFILS';
  T := Tob.Create('Ligne grille',TobGrille,-1);
  T.AddChampSupValeur('CHAMP',Libelle);
  If ProfilParDefaut then
  begin
     T.AddChampSupValeur('HISTOSTD','#ICO#47');
     T.AddChampSupValeur('HISTODOS','#ICO#47');
  end
  else
  begin
    If ExisteSQL('SELECT PPP_HISTORIQUE FROM PARAMSALARIE WHERE PPP_PREDEFINI="STD" AND PPP_HISTORIQUE="X" AND PPP_PGTHEMESALARIE="PRO"') then
    begin
      T.AddChampSupValeur('HISTOSTD','#ICO#47');
      T.AddChampSupValeur('HISTODOS','#ICO#47');
    end
    Else
    begin
       T.AddChampSupValeur('HISTOSTD','#ICO#46');
       If ExisteSQL('SELECT PPP_HISTORIQUE FROM PARAMSALARIE WHERE PPP_PREDEFINI="DOS" AND PPP_NODOSSIER="'+V_PGI.NoDossier+'" AND PPP_HISTORIQUE="X" AND PPP_PGTHEMESALARIE="PRO"') then T.AddChampSupValeur('HISTODOS','#ICO#47')
       Else T.AddChampSupValeur('HISTODOS','#ICO#46');
    end;
  end;
  T.AddChampSupValeur('CODE',Champ);
  T.AddChampSupValeur('CODETHEME',Theme);
  TobChamp.Free;
  TobGrille.PutGridDetail(grille,False,False,'CHAMP;HISTOSTD;HISTODOS;CODE;CODETHEME',False);
  HMTrad.ResizeGridColumns(Grille);
end;

procedure TOF_PARAMHISTOSTDDOS.ClickGrille(Sender : TObject);
var Ligne,i,Colonne : Integer;
    DefautDos,DefautStd : Boolean;
    Valeur : String;
begin
     Colonne := Grille.Col;
     If (Colonne = 1) or (Colonne = 2) then
     begin
          Ligne := Grille.Row;
          If Grille.CellValues[3,Ligne] = 'THEME' then Exit;
          Valeur := Grille.CellValues[1,Ligne];
          If (Valeur = '#ICO#89') OR (Valeur = '#ICO#47') then DefautStd := True
          else DefautStd := False;
          Valeur := Grille.CellValues[2,Ligne];
          If (Valeur = '#ICO#89') OR (Valeur = '#ICO#47') then DefautDos := True
          else DefautDos := False;
          If (Colonne = 1) and (STD) then    //PT1
          begin
            If defautStd then Grille.CellValues[1,Ligne] := '#ICO#88'
            else
            begin
              Grille.CellValues[1,Ligne] := '#ICO#89';
              Grille.CellValues[2,Ligne] := '#ICO#89';
            end;
          end
          else
          if (Colonne = 2) and (DOS or STD) then  //PT1
          begin
            If defautStd then Grille.CellValues[2,Ligne] := '#ICO#89'
            else
            begin
               If not defautDos then Grille.CellValues[2,Ligne] := '#ICO#89'
               else Grille.CellValues[2,Ligne] := '#ICO#88';
            end;
          end;
     end;
end;

procedure TOF_PARAMHISTOSTDDOS.GrilleGetCellCanvas( ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
begin
 If (Grille.CellValues[3,ARow] = 'THEME') and Not(Grille.IsSelected(ARow)) then
  begin
    Canvas.Font.Style := [fsBold];
    Canvas.Brush.Color := ClInactiveCaption;
    end;
 end;

Function TOF_PARAMHISTOSTDDOS.ValiderSaisie : Boolean;
var TobParam,T : Tob;
    i : Integer;
    Q : TQuery;
    Champ,HistoStd,HistoDos,Theme,Libelle : String;
    Suffix,TypeD,Tablette : String;
begin
  Result := False;
  TobGrille.GetGridDetail (Grille,Grille.RowCount-1,'','CHAMP;HISTOSTD;HISTODOS;CODE;CODETHEME');
  Q := OpenSQL('SELECT * FROM PARAMSALARIE WHERE PPP_PREDEFINI="STD" OR (PPP_PREDEFINI="DOS" AND PPP_NODOSSIER="'+V_PGI.NoDossier+'")',true);
  TobParam := Tob.Create('PARAMSALARIE',Nil,-1);
  TobParam.LoadDetailDB('PARAMSALARIE','','',Q,False);
  Ferme(Q);
  If Not VerifierLesDonnes(TobGrille,TobParam) then exit;
  For i := 0 to TobGrille.Detail.Count - 1 do
  begin
    Champ := TobGrille.Detail[i].GetValue('CODE');
    HistoStd := TobGrille.Detail[i].GetValue('HISTOSTD');
    HistoDos := TobGrille.Detail[i].GetValue('HISTODOS');
    If (HistoStd='#ICO#89') or (HistoDOS='#ICO#89') then Result := True;
    If Champ = 'LESPROFILS' then
    begin
      MajProfil(TobParam,(HistoStd = '#ICO#89') OR (HistoStd = '#ICO#47'),(HistoDos = '#ICO#89') OR (HistoDos = '#ICO#47')); //Traitement spécifique des profils
      Continue;
    end;
    Libelle := TobGrille.Detail[i].GetValue('CHAMP');
    Theme := TobGrille.Detail[i].GetValue('CODETHEME');
    T := TobParam.FindFirst(['PPP_PGINFOSMODIF','PPP_PREDEFINI'],[champ,'STD'],False);
    If T <> Nil then
    begin
      If (HistoStd = '#ICO#89') OR (HistoStd = '#ICO#47') then T.Putvalue('PPP_HISTORIQUE','X')
      else T.Putvalue('PPP_HISTORIQUE','-');
      T.UpdateDB(False);
    end
    else
    begin
      If (HistoStd = '#ICO#89') OR (HistoStd = '#ICO#47') then
      begin
        T := Tob.Create('PARAMSALARIE',TobParam,-1);
        Suffix := Champ;
//        Suffix := Copy(Suffix,5,Length(Suffix));
        Q := OpenSQL('SELECT * FROM PARAMSALARIE WHERE PPP_PGINFOSMODIF="'+Champ+'" AND PPP_PREDEFINI="CEG"',True);
        if not Q.eof then
        begin
          Tablette := Q.FindField('PPP_LIENASSOC').AsString;
          TypeD := Q.FindField('PPP_PGTYPEDONNE').AsString;
        end;
        Ferme(Q);
        T.Putvalue('PPP_PGINFOSMODIF',Champ);
        T.Putvalue('PPP_PGTYPEINFOLS','SAL');
        T.Putvalue('PPP_LIBELLE',Libelle);
        T.Putvalue('PPP_CHAMPSAL',Champ);
        T.Putvalue('PPP_LIENASSOC',Tablette);
        T.Putvalue('PPP_HISTORIQUE','X');
        T.Putvalue('PPP_PREDEFINI','STD');
        T.Putvalue('PPP_NODOSSIER','000000');
        T.Putvalue('PPP_PGTHEMESALARIE',Theme);
        T.Putvalue('PPP_PGTYPEDONNE',TypeD);
        T.Putvalue('PPP_TYPENIVEAU','');
        T.Putvalue('PPP_CODTABL','');
        T.InsertOrUpdateDB(False);
      end;
    end;
    If (HistoStd = '#ICO#46') OR (HistoStd = '#ICO#88') then
    begin
      If (HistoDos = '#ICO#89') OR (HistoDos = '#ICO#47') then
      begin
        T := TobParam.FindFirst(['PPP_PGINFOSMODIF','PPP_PREDEFINI'],[champ,'DOS'],False);
        If T <> Nil then
        begin
//          If (HistoStd = '#ICO#89') OR (HistoStd = '#ICO#47') then
//          begin
            T.Putvalue('PPP_HISTORIQUE','X');
            T.UpdateDB(False);
//          end
//          else T.DeleteDB(False);
        end
        else
        begin
          T := Tob.Create('PARAMSALARIE',TobParam,-1);
          Suffix := Champ;
//          Suffix := Copy(Suffix,5,Length(Suffix));
          Q := OpenSQL('SELECT * FROM PARAMSALARIE WHERE PPP_PGINFOSMODIF="'+Champ+'" AND PPP_PREDEFINI="CEG"',True);
          if not Q.eof then
          begin
            Tablette := Q.FindField('PPP_LIENASSOC').AsString;
            TypeD := Q.FindField('PPP_PGTYPEDONNE').AsString;
          end;
          Ferme(Q);
          T.Putvalue('PPP_PGINFOSMODIF',Champ);
          T.Putvalue('PPP_PGTYPEINFOLS','SAL');
          T.Putvalue('PPP_LIBELLE',Libelle);
          T.Putvalue('PPP_CHAMPSAL',Champ);
          T.Putvalue('PPP_LIENASSOC',Tablette);
          T.Putvalue('PPP_HISTORIQUE','X');
          T.Putvalue('PPP_PREDEFINI','DOS');
          T.Putvalue('PPP_NODOSSIER','000000');
          T.Putvalue('PPP_PGTHEMESALARIE',Theme);
          T.Putvalue('PPP_PGTYPEDONNE',TypeD);
          T.Putvalue('PPP_TYPENIVEAU','');
          T.Putvalue('PPP_CODTABL','');
          T.InsertOrUpdateDB(False);
        end;
      end
      else
      begin
        T := TobParam.FindFirst(['PPP_PGINFOSMODIF','PPP_PREDEFINI'],[champ,'DOS'],False);
        If T <> Nil then
        begin
            T.DeleteDB;
//            T.Putvalue('PPP_HISTORIQUE','-');
//            T.UpdateDB(False);
        end;
      end;
    end;
  end;
  TobParam.free;
  AfficheGrille;
end;

Function TOF_PARAMHISTOSTDDOS.RendLibelle(LeChamp : String) : String;
begin
    Result := '';
    If (Copy(LeChamp,1,15) = 'PSA_SALAIREMOIS') then //Salaire mensuel
    begin
      If (Copy(LeChamp,16,1) = '1') and (VH_Paie.PgNbSalLib > 0) then Result := 'Salaire mensuel : '+VH_Paie.PgSalLib1
      else if (Copy(LeChamp,16,1) = '2') and (VH_Paie.PgNbSalLib > 1) then Result := 'Salaire mensuel : '+VH_Paie.PgSalLib2
      else If (Copy(LeChamp,16,1) = '3') and (VH_Paie.PgNbSalLib > 2) then Result := 'Salaire mensuel : '+ VH_Paie.PgSalLib3
      else If (Copy(LeChamp,16,1) = '4') and (VH_Paie.PgNbSalLib > 3) then Result := 'Salaire mensuel : '+VH_Paie.PgSalLib4
      else If (Copy(LeChamp,16,1) = '5') and (VH_Paie.PgNbSalLib > 4) then Result := 'Salaire mensuel : '+VH_Paie.PgSalLib5
      else Result := '';
    end
    else If (Copy(LeChamp,1,13) = 'PSA_SALAIRANN') then //Salaire ANNUEL
    begin
      If (Copy(LeChamp,14,1) = '1') and (VH_Paie.PgNbSalLib > 0) then Result := 'Salaire annuel : '+VH_Paie.PgSalLib1
      else if (Copy(LeChamp,14,1) = '2') and (VH_Paie.PgNbSalLib > 1) then Result := 'Salaire annuel : '+VH_Paie.PgSalLib2
      else If (Copy(LeChamp,14,1) = '3') and (VH_Paie.PgNbSalLib > 2) then Result := 'Salaire annuel : '+VH_Paie.PgSalLib3
      else If (Copy(LeChamp,14,1) = '4') and (VH_Paie.PgNbSalLib > 3) then Result := 'Salaire annuel : '+VH_Paie.PgSalLib4
      else If (Copy(LeChamp,14,1) = '5') and (VH_Paie.PgNbSalLib > 4) then Result := 'Salaire annuel : '+VH_Paie.PgSalLib5
      else Result := '';
    end
    else If (Copy(LeChamp,1,12) = 'PSA_TRAVAILN') then // Champs organisation
    begin
      If (Copy(LeChamp,13,1) = '1') and (VH_Paie.PGNbreStatOrg > 0) then Result := VH_Paie.PGLibelleOrgStat1
      else if (Copy(LeChamp,13,1) = '2') and (VH_Paie.PGNbreStatOrg > 1) then Result := VH_Paie.PGLibelleOrgStat2
      else If (Copy(LeChamp,13,1) = '3') and (VH_Paie.PGNbreStatOrg > 2) then Result := VH_Paie.PGLibelleOrgStat3
      else If (Copy(LeChamp,13,1) = '4') and (VH_Paie.PGNbreStatOrg > 3) then Result := VH_Paie.PGLibelleOrgStat4
      else Result := '';
    end
    else If LeChamp = 'PSA_CODESTAT' then //Code statistique
    begin
      If VH_Paie.PGLibCodeStat <> '' then Result := VH_Paie.PGLibCodeStat
      else Result := '';
    end
    else If (Copy(LeChamp,1,13) = 'PSA_LIBREPCMB') then // Combos libre
    begin
      If (Copy(LeChamp,14,1) = '1') and (VH_Paie.PgNbCombo > 0) then Result := VH_Paie.PgLibCombo1
      else if (Copy(LeChamp,14,1) = '2') and (VH_Paie.PgNbCombo > 1) then Result := VH_Paie.PgLibCombo2
      else If (Copy(LeChamp,14,1) = '3') and (VH_Paie.PgNbCombo > 2) then Result := VH_Paie.PgLibCombo3
      else If (Copy(LeChamp,14,1) = '4') and (VH_Paie.PgNbCombo > 3) then Result := VH_Paie.PgLibCombo4
      else Result := '';
    end
    else if (Copy(LeChamp,1,13) = 'PSA_DATELIBRE') then
    begin
      If (Copy(LeChamp,14,1) = '1') and (VH_Paie.PgNbDate > 0) then Result := VH_Paie.PgLibDate1
      else if (Copy(LeChamp,14,1) = '2') and (VH_Paie.PgNbDate > 1) then Result := VH_Paie.PgLibDate2
      else If (Copy(LeChamp,14,1) = '3') and (VH_Paie.PgNbDate > 2) then Result := VH_Paie.PgLibDate3
      else If (Copy(LeChamp,14,1) = '4') and (VH_Paie.PgNbDate > 3) then Result := VH_Paie.PgLibDate4
      else Result := '';
    end;
end;

procedure TOF_PARAMHISTOSTDDOS.BAjoutClick(Sender: TObject);
var myRect : TGridRect;
begin
  myRect := Grille.Selection;
  ChangeValueSelection(myRect,'#ICO#89');
end;

procedure TOF_PARAMHISTOSTDDOS.BMoinsClick(Sender: TObject);
var myRect : TGridRect;
begin
  myRect := Grille.Selection;
  ChangeValueSelection(myRect,'#ICO#88');
end;

procedure TOF_PARAMHISTOSTDDOS.ChangeValueSelection(MyRect : TGridRect ; Value : string);
var iCol,iRow : integer;
    DefautStd,DefautDos : Boolean;
    Valeur : String;
begin
  myRect := Grille.Selection;
  for iRow := myRect.Top to myRect.Bottom  do
  begin
    for iCol := myRect.Left to myRect.Right do
    begin
      If Grille.CellValues[3,iRow] = 'THEME' then Continue;
      Valeur := Grille.CellValues[1,iRow];
      If (Valeur = '#ICO#89') OR (Valeur = '#ICO#47') then DefautStd := True
          else DefautStd := False;
          Valeur := Grille.CellValues[2,iRow];
          If (Valeur = '#ICO#89') OR (Valeur = '#ICO#47') then DefautDos := True
          else DefautDos := False;
          If (iCol = 1) and (STD) then   //PT1
          begin
            If (defautStd) and (Value = '#ICO#88') then Grille.CellValues[1,iRow] := '#ICO#88'
            else
            begin
              If Value = '#ICO#89' then
              begin
                Grille.CellValues[1,iRow] := '#ICO#89';
                Grille.CellValues[2,iRow] := '#ICO#89';
              end;
            end;
          end
          else
          if (iCol = 2) and (DOS or STD) then  //PT1
          begin
            If defautStd then Grille.CellValues[2,iRow] := '#ICO#89'
            else
            begin
               If (not defautDos) and (Value = '#ICO#89')  then Grille.CellValues[2,iRow] := '#ICO#89'
               else if (Value = '#ICO#88') then Grille.CellValues[2,iRow] := '#ICO#88';
            end;
          end;
    {  If ICol = 2 then
      begin
          If (Grille.CellValues[1,iRow] = '#ICO#89') OR (Grille.CellValues[1,iRow] = '#ICO#47') then DefautStd := True
          else DefautStd := False;
          If Not DefautStd then Grille.CellValues[iCol,iRow] := Value;
      end
      else
      begin
        If (Grille.CellValues[1,iRow] = '#ICO#89') OR (Grille.CellValues[1,iRow] = '#ICO#47') then DefautStd := True
          else DefautStd := False;
        Grille.CellValues[iCol,iRow] := Value;
        If defautStd then
        begin
          If (Grille.CellValues[2,iRow] = '#ICO#88') OR (Grille.CellValues[2,iRow] = '#ICO#46') then Grille.CellValues[iCol,iRow] := '#ICO#89';
        end;
      end;  }
    end;
  end;
end;

procedure TOF_PARAMHISTOSTDDOS.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var FocusGrid : Boolean;
begin
FocusGrid := (Screen.ActiveControl = Grille);
Case Key of
    {Ctrl+A}  65 : if Shift=[ssCtrl] then if FocusGrid then begin Key:=0 ; ChangeValueSelection(Grille.Selection,'#ICO#89'); end ;
    {Ctrl+R}  82 : if Shift=[ssCtrl] then if FocusGrid then begin Key:=0 ; ChangeValueSelection(Grille.Selection,'#ICO#88'); end ;
    end;
end;

procedure TOF_PARAMHISTOSTDDOS.MajProfil(TobParam : Tob;ProfilStd,ProfilDos : Boolean);
var T,TobCegProf : Tob;
    i : Integer;
    Q : TQuery;
    Champ : String;
begin
  If (ProfilSTD = False) and (ProfilDOS = False) then Exit;
  Q := OpenSQL('SELECT * FROM PARAMSALARIE WHERE PPP_PREDEFINI="CEG" AND PPP_PGTHEMESALARIE="PRO"',true);
  TobCegProf := Tob.Create('les profils',Nil,-1);
  TobCegProf.LoadDetailDB('les profils','','',Q,False);
  Ferme(Q);
  For i := 0 to TobCegProf.Detail.Count - 1 do
    begin
      Champ := TobCegProf.Detail[i].GetValue('PPP_PGINFOSMODIF');
      T := TobParam.FindFirst(['PPP_PGINFOSMODIF','PPP_PREDEFINI'],[champ,'STD'],False);
      If T = Nil then T := Tob.Create('PARAMSALARIE',TobParam,-1);
      T.Putvalue('PPP_PGINFOSMODIF',TobCegProf.Detail[i].GetValue('PPP_PGINFOSMODIF'));
      T.Putvalue('PPP_PGTYPEINFOLS','SAL');
      T.Putvalue('PPP_LIBELLE',TobCegProf.Detail[i].GetValue('PPP_LIBELLE'));
      T.Putvalue('PPP_CHAMPSAL',Champ);
      T.Putvalue('PPP_LIENASSOC',TobCegProf.Detail[i].GetValue('PPP_LIENASSOC'));
      If ProfilSTD then T.Putvalue('PPP_HISTORIQUE','X')
      else T.Putvalue('PPP_HISTORIQUE','-');
      T.Putvalue('PPP_PREDEFINI','STD');
      T.Putvalue('PPP_NODOSSIER','000000');
      T.Putvalue('PPP_PGTHEMESALARIE',TobCegProf.Detail[i].GetValue('PPP_PGTHEMESALARIE'));
      T.Putvalue('PPP_PGTYPEDONNE',TobCegProf.Detail[i].GetValue('PPP_PGTYPEDONNE'));
      T.Putvalue('PPP_TYPENIVEAU','');
      T.Putvalue('PPP_CODTABL','');
      T.InsertOrUpdateDB(False);
      T := TobParam.FindFirst(['PPP_PGINFOSMODIF','PPP_PREDEFINI','PPP_NODOSSIER'],[champ,'DOS',V_PGI.NoDossier],False);
      If (T <> Nil) and (Not ProfilDOS) then T.DeleteDB(False)
      else If T <> Nil then
      begin
        T.Putvalue('PPP_HISTORIQUE','X');
        T.UpdateDB;
      end
      else If ProfilDOS then
      begin
        If T = Nil then T := Tob.Create('PARAMSALARIE',TobParam,-1);
        T.Putvalue('PPP_PGINFOSMODIF',TobCegProf.Detail[i].GetValue('PPP_PGINFOSMODIF'));
        T.Putvalue('PPP_PGTYPEINFOLS','SAL');
        T.Putvalue('PPP_LIBELLE',TobCegProf.Detail[i].GetValue('PPP_LIBELLE'));
        T.Putvalue('PPP_CHAMPSAL',Champ);
        T.Putvalue('PPP_LIENASSOC',TobCegProf.Detail[i].GetValue('PPP_LIENASSOC'));
        T.Putvalue('PPP_HISTORIQUE','X');
        T.Putvalue('PPP_NODOSSIER',V_PGI.NoDossier);
        T.Putvalue('PPP_PREDEFINI','DOS');
        T.Putvalue('PPP_PGTHEMESALARIE',TobCegProf.Detail[i].GetValue('PPP_PGTHEMESALARIE'));
        T.Putvalue('PPP_PGTYPEDONNE',TobCegProf.Detail[i].GetValue('PPP_PGTYPEDONNE'));
        T.Putvalue('PPP_TYPENIVEAU','');
        T.Putvalue('PPP_CODTABL','');
        T.InsertOrUpdateDB(False);
      end
  end;
end;

Function TOF_PARAMHISTOSTDDOS.VerifierLesDonnes(TobModif,TobInit : Tob) : Boolean;
var i : Integer;
    HistoSTD,HistoDOS : boolean;
    Champ,Predefini,ListeChamp : String;
    T,TobDossier,Tdos : Tob;
    Rep : Word;
    Q : TQuery;
    Retour : String;
begin
  Q := OpenSQL('SELECT * FROM PARAMSALARIE WHERE PPP_PREDEFINI="DOS" AND PPP_NODOSSIER="'+V_PGI.NoDossier+'"',true);
  TobDossier := Tob.Create('les dossier',Nil,-1);
  TobDossier.LoadDetailDB('le dossier','','',Q,False);
  Ferme(Q);
  Result := True;
  ListeChamp := '';
  for i := 0 to TobInit.Detail.Count - 1 do
  begin
    Predefini := TobInit.Detail[i].GetValue('PPP_PREDEFINI');
    If Predefini <> 'STD' then Exit;
    HistoSTD := (TobInit.Detail[i].GetValue('PPP_HISTORIQUE') = 'X');
    Champ := TobInit.Detail[i].GetValue('PPP_PREDEFINI');
    TDos := TobDossier.FindFirst(['CODE'],[Champ],False);
    If TDos <> Nil then
    begin
      If TDos.GetValue('PPP_HISTORIQUE') = 'X' then HistoDos   := True
      else HistoDos := False;
    end
    else HistoDOS := False;
    T := TobModif.FindFirst(['CODE'],[Champ],False);
    If T <> Nil then
    begin
      If Predefini = 'STD' then
      begin
        If (HistoSTD = True) and ((T.GetValue('HISTOSTD') = '#ICO#88') OR (T.GetValue('HISTOSTD') = '#ICO#46')) then
        begin
          ListeChamp := ListeChamp + '#13#10 - '+RechDom('PGCHAMPSAL',Champ,False);
        end;
      end
      else If Predefini = 'DOS' then
      begin
        If (HistoSTD = False) and (HistoDOS = True) and ((T.GetValue('HISTODOS') = '#ICO#88') OR (T.GetValue('HISTODOS') = '#ICO#46')) then
        begin
          ListeChamp := ListeChamp + '#13#10 - '+RechDom('PGCHAMPSAL',Champ,False);
        end;
      end;
    end;
  end;
  If ListeChamp <> '' then
  begin
    Rep := PGIAsk('Attention l''historisation ne sera plus active pour les champs suivants : '+ListeChamp+
     '#13#10 Voulez vous continuer ?',
          ecran.Caption);
    if Rep = MrYes then
      begin
        Retour := AglLanceFiche('PAY', 'CONFIRMPAIE', '', '', 'Confirmez-vous la désactivation de ces champs');
        if retour <> 'OUI' then
          begin
            Exit;
          end;
        end
      else exit;
    end;
end;

Initialization
  registerclasses ( [ TOF_PARAMHISTOSTDDOS ] ) ;
end.


