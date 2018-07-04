{***********UNITE*************************************************
Auteur  ...... : B. LOCATELLI
Créé le ...... : 19/04/2012
Modifié le ... : 19/04/2012
Description .. : Fonctions @ pour les MUL
Mots clefs ... : BTP
*****************************************************************}
unit CalcMulBTP;

interface

uses
  Classes,
  Hctrls,
  db,
  controls,
  SaisUtil,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fe_Main,
  Forms,
  sysutils,
  Windows,
  HEnt1,
  HMsgBox,
  UTOB,
  Entgc,
  uEntCommun,
  Ent1,
  AGLInit,
  AGLInitBTP,
  DateUtils,
  ED_TOOLS,
  Variants,
  HRichOle,
  utilxlsBTP;

function BTProcCalcMul(Func, Params, WhereSQL: hstring; TT: TDataset; Total: boolean): hString;

implementation

uses NomenUtil,UtilArticle;

Function RecupAvancement(Aff : String) : String;
Var Req : String;
		Q : Tquery;
    DateDebutPrev, DateFin, Dt : TDateTime;
    i : integer;
begin
  Result := '';

  // récupération date de début
  Req := 'SELECT ZPR_DATEDEBUT FROM ZPREVAFF WHERE ZPR_AFFAIRE ="'+Aff+'"';
  Q := OpenSQL(Req, True);
  if Not Q.EOF then
  begin
    DateDebutPrev := Q.Fields[0].AsDateTime;
  end else
  begin
    DateDebutPrev := Date();
  end;
  ferme(Q);

  // Avancement à prendre pour le mois précédent la date du jour donc la date de fin est la date du jour moins 1 mois
  DateFin := PLUSMOIS(V_PGI.DateEntree,-1);

  // On prend l'avancement si la date de début de prévisionnel est inférieure ou égale à la date de fin
  if DateDebutPrev <= DateFin then
  begin
    //détermination de l'avancement total
    // Calcul du numéro du champ correspondant au mois de la date de fin
    for i := 0 to 23 do
    begin
      Dt := PLUSMOIS(DateDebutPrev,i);
      if (YearOf(Dt) = YearOf(DateFin)) and (MonthOf(Dt) = MonthOf(DateFin)) then Break;
    end;
    // Préparation de la requête
    Req := 'SELECT ';
    if i = 0 then Req := Req + 'ZAV_POURCENTAGE1 '
    else if i = 1 then Req := Req + 'ZAV_POURCENTAGE2 '
    else if i = 2 then Req := Req + 'ZAV_POURCENTAGE3 '
    else if i = 3 then Req := Req + 'ZAV_POURCENTAGE4 '
    else if i = 4 then Req := Req + 'ZAV_POURCENTAGE5 '
    else if i = 5 then Req := Req + 'ZAV_POURCENTAGE6 '
    else if i = 6 then Req := Req + 'ZAV_POURCENTAGE7 '
    else if i = 7 then Req := Req + 'ZAV_POURCENTAGE8 '
    else if i = 8 then Req := Req + 'ZAV_POURCENTAGE9 '
    else if i = 9 then Req := Req + 'ZAV_POURCENTAGE10 '
    else if i = 10 then Req := Req + 'ZAV_POURCENTAGE11 '
    else if i = 11 then Req := Req + 'ZAV_POURCENTAGE12 '
    else if i = 12 then Req := Req + 'ZAV_POURCENTAGE13 '
    else if i = 13 then Req := Req + 'ZAV_POURCENTAGE14 '
    else if i = 14 then Req := Req + 'ZAV_POURCENTAGE15 '
    else if i = 15 then Req := Req + 'ZAV_POURCENTAGE16 '
    else if i = 16 then Req := Req + 'ZAV_POURCENTAGE17 '
    else if i = 17 then Req := Req + 'ZAV_POURCENTAGE18 '
    else if i = 18 then Req := Req + 'ZAV_POURCENTAGE19 '
    else if i = 19 then Req := Req + 'ZAV_POURCENTAGE20 '
    else if i = 20 then Req := Req + 'ZAV_POURCENTAGE21 '
    else if i = 21 then Req := Req + 'ZAV_POURCENTAGE22 '
    else if i = 22 then Req := Req + 'ZAV_POURCENTAGE23 '
    else Req := Req + 'ZAV_POURCENTAGE24 ';
    Req := Req + 'FROM ZAVANCEMENTAFF WHERE ZAV_AFFAIRE ="'+Aff+'"';
    // Lecture avancement
    Q := OpenSQL(Req, True);
    if Not Q.EOF then
    begin
      Result := FloatToStr(Q.Fields[0].AsFloat);
    end;
    ferme(Q);
  end;
end;

Function GetPuOuvHT(CodeArt,TypeArticle : String; PVHt : double) : String;
Var Valeurs : T_Valeurs;
begin

  if (pos(typearticle, 'ARP;OUV')>0) then
  begin
    ValoriseOuvrage (CodeArt,nil,nil,valeurs,true,true);
    Result := Strs(Valeurs[2],V_PGI.OkDecP);
  end
  else if (pos(typearticle ,'MAR;PRE;POU;FRS')>0) then
  begin
     REsult := Strs(PVHT,V_PGI.OkDecP);
  end else
  begin
    Result :=  '***';
  end;

end;

Function DECRYPTEBLOB(TT : TDataset; Params : String) : HString;
Var i					: Integer;
    BlocNote  : THRichEditOLE;
Begin

  Result := '';
  if TT.FindField(Params) = nil then exit;
  BlocNote := THRichEditOLE.Create(application.MainForm);
  BlocNote.Parent := Application.MainForm;
  BlocNote.Visible := false;

  StringToRich(BlocNote,TT.FindField(Params).AsString);

  TRY
    for i := 0 to BlocNote.lines.Count - 1 do
    Begin
      if i > 0 then result := Result + ' ' + CHR(10);
      Result := Result + BlocNote.lines[i];
		end;
  FINALLY
    FreeAndNil(Blocnote);
  END;

end;

function BTProcCalcMul(Func, Params, WhereSQL: hstring; TT: TDataset; Total: boolean): hString;
begin

  Result := '';

  if Func = 'CEPP_Avancement' then // Fonction spécifique à POUCHAIN
  begin
    Result := RecupAvancement(TT.findfield('AFF_AFFAIRE').AsString);
  end
  Else if Func = 'BTCalcPUOuv' then  //calcul du PU HT sur article en prix posé et Ouvrage
  begin
    if TT.FindField('GA_TYPEARTICLE')=nil then Exit;
    if TT.FindField ('GA_PVHT')=nil then Exit;
    Result := GetPuOuvHT(TT.findfield('GA_CODEARTICLE').AsString,TT.findfield('GA_TYPEARTICLE').AsString,TT.findfield('GA_PVHT').Asfloat);
  end
  else if func = 'DECRYPTEBLOB' then //Remise en forme du Blob dans un mul
  begin
    if Params='' then  exit;
    Result := DECRYPTEBLOB(TT, Params);
  end
  else if func = 'GETPAREMISE' then //Remise en forme du Blob dans un mul
  begin
    if TT.findfield('GA_ARTICLE') <> nil then
    begin
      Result := StrF00(GetPrixRemise (TT.findfield('GA_ARTICLE').AsString,'ACHAT'),V_PGI.okdecP);
    end else
    begin
      result := '0.0';
    end;
  end
  Else if func='RTISFAX' then
  begin
    Result := 'N';
    //if (pos('((T_FAX <> "") or (C_FAX <> ""))',WhereSQL) <> 0) then  Result := 'O';
    if (pos('(T_FAX <> "")',WhereSQL) <> 0) or (pos('(C_FAX <> "")',WhereSQL) <> 0) then  Result := 'O';
  end;


end;

end.
