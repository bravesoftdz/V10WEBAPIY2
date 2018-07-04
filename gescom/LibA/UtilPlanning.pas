{***********UNITE*************************************************
Auteur  ...... : CB
Créé le ...... : 26/03/2001
Modifié le ... :
Description .. : Fonction générales au planning
Suite ........ : Notamment gestion des paramètres
Mots clefs ... : PLANNING
*****************************************************************}
unit UtilPlanning;

interface

Uses HEnt1, HCtrls, UTOB, Ent1, Hplanning, Controls,
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     DB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,
{$ENDIF}
     SysUtils, Dialogs, UtilPGI, AGLInit, EntGC, HMsgBox, paramsoc, AFPlanningCst,
     HStatus, graphics, dicobtp;        

  function GetNumLignePlanning (pStAffaire : String) : integer;

  function DecodeFontStyle(pFontStyle : TFontStyles) : String;
  function EncodeFontStyle(pFontStyle : String) : TFontStyles;

  Function FormatJour(pRdQte : Double) : String;
  Function LoadMois(var pTobLigne : Tob; pTobTache : Tob; pBlFonction : Boolean;
                    pStAffaire, pStPrefixe, pStChamp : String) : Double;

  procedure LoadMoisCra(var pTobLigne : Tob; pStAffaire : String);

  Function TobStrToFloat(pVaQte : Variant) : Double;
  Procedure ReplaceSubStr(Var pSt : String; pStSup, pStRep : string);
  function EtatModifAutorisee(pStEtat : String) : boolean;
  function EtatInterdit(pInMode : Integer; pStAffaire : String; pTache : RecordTache) : Boolean;

implementation

//*********************** Numérotation des tâches ******************************
function GetNumLignePlanning (pStAffaire : String) : integer;
Var
  Q : Tquery;
Begin
  Result := 1;
  if pStAffaire <> '' then
    Begin
      Q := OpenSQL ('SELECT MAX(APL_NUMEROLIGNE) FROM AFPLANNING WHERE APL_AFFAIRE = "'+ pStAffaire +'"',True);
      if Not Q.EOF then
        Result:=Q.Fields[0].AsInteger + 1;
      Ferme(Q) ;
    End;
End;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 16/03/2002
Modifié le ... :
Description .. : decode le style de font
Suite ........ :
Mots clefs ... :
*****************************************************************}
function DecodeFontStyle(pFontStyle : TFontStyles) : String;
begin

  if fsBold in pFontStyle then
    result := 'fsBold';

  if fsItalic in pFontStyle then
    if result = '' then
      result := 'fsItalic'
    else
      result := result + ';' + 'fsItalic';

  if fsUnderline in pFontStyle then
    if result = '' then
      result := 'fsUnderline'
    else
      result := result + ';' + 'fsUnderline';

  if fsStrikeOut in pFontStyle then
    if result = '' then
      result := 'fsStrikeOut'
    else
      result := result + ';' + 'fsStrikeOut';

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 16/03/2002
Modifié le ... :
Description .. : encode le style de font
Suite ........ :
Mots clefs ... :
*****************************************************************}
function EncodeFontStyle(pFontStyle : String) : TFontStyles;
var
  vFontStyle : String;
  vStyle : String;

begin

  result := [];
  vFontSTyle := pFontStyle;
  while (vFontSTyle <> '') do
    begin
      vStyle := ReadTokenSt(vFontSTyle);
      if vStyle = 'fsUnderline' then result := result + [fsUnderline]
      else if vStyle = 'fsBold' then result := result + [fsBold]
      else if vStyle = 'fsItalic' then result := result + [fsItalic]
      else if vStyle = 'fsStrikeOut' then result := result + [fsStrikeOut];
    end;
                                        
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 14/05/2002
Modifié le ... :   /  /
Description .. : Formatage des jours dans le plan de charge
Mots clefs ... :
*****************************************************************}
Function FormatJour(pRdQte : Double) : String;
begin
  if getparamsoc('SO_AFPDCDEC') then
    begin
      if FormatFloat('##0.00', pRdQte) = '' then
        Result := FormatFloat('##0.00', 0)
      else
        Result := FormatFloat('##0.00', pRdQte);
    end
  else
    begin
      if FormatFloat('##', pRdQte) = '' then
        Result := '0'
      else
        Result := FormatFloat('##', pRdQte);
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 15/05/2002
Modifié le ... :   /  /
Description .. : Calcul des quantites des mois pour une ligne
               : et calcul de la quantité affectée
               : la quantité affectée est calculée a partir du 1er
               : du mois en cours
Mots clefs ... :
*****************************************************************}
Function LoadMois(var pTobLigne : Tob; pTobTache : Tob; pBlFonction : Boolean; pStAffaire, pStPrefixe, pStChamp : String) : Double;
var
  i             : Integer;
  S             : String;
  vTobMois      : Tob;
  vQrMois       : TQuery;
  vWdYear       : Word;
  vWdMonth      : Word;
  vWdDay        : Word;
  vStmois       : String;
  vDtDepart     : TDateTime;
  vDtNow        : TDateTime;
  vSt           : String;

begin

  Result := 0;
  vTobMois := TOB.create('Les mois', nil, -1);
  vDtDepart := DebutDeMois(now);

  // chargement des mois pour la ligne fonction (pas de ressource affectée) de cette tache
  S := 'SELECT APL_DATEDEBPLA, ' + pStChamp + ' FROM AFPLANNING ';
  S := S + ' where APL_AFFAIRE = "' + pStAffaire + '"';
  S := S + ' and APL_NUMEROTACHE = ' + intToStr(pTobTache.Getvalue(pStPrefixe + '_NUMEROTACHE'));
  S := S + ' and APL_DATEDEBPLA >= "' + UsDateTime(vDtDepart) +'"';

  if pBlFonction then
    S := S + 'and APL_RESSOURCE = ""'
  else
    S := S + ' and APL_RESSOURCE = "' + pTobTache.GetValue(pStPrefixe + '_RESSOURCE') + '"';


  vQrMois := nil;
  Try
    vQrMois := OpenSql(S,True);
    if Not vQrMois.Eof then
      begin
        vTobMois.LoadDetailDB('LesMois','','',vQrMois,False,True);

        // affectation des mois pour la tache
        for i := 0 to vTobMois.Detail.Count - 1 do
          begin

            // calcul du mois
            // si 1er du mois, on se positionne le 2,
            // pour eviter d'avoir 31 jours au lien d'un mois lors de la difference entre les 2 dates
            // idem si 31, on se positionne sur le 30 
            DecodeDate(now, vWdYear, vWdMonth, vWdDay);
            if vWdDay = 1 then
              vDtNow := now + 1
            else if vWdDay = 31 then
              vDtNow := now - 1
            else vDtNow := now;
                            
            DecodeDate(strToDate(vTobMois.detail[i].GetValue('APL_DATEDEBPLA')) - vDtNow, vWdYear, vWdMonth, vWdDay);
            if vWdMonth <> 12 then
              vStMois := 'PC_MOIS' + IntToStr(vWdMonth)
            else
              vStMois := 'PC_MOIS';

            //if (FormatJour(vTobMois.detail[i].getvalue(pStChamp )) <> '0') then
            if (vTobMois.detail[i].getvalue(pStChamp) <> 0) then
              // somme des quantités
              begin
                vSt := pTobLigne.GetValue(vStMois);
//                pTobLigne.putValue(vStMois, FormatJour(TobStrToFloat(vSt) + StrToFLoat(vTobMois.detail[i].getvalue(pStChamp))));
                pTobLigne.putValue(vStMois, TobStrToFloat(vSt) + StrToFLoat(vTobMois.detail[i].getvalue(pStChamp)));

              end
            else
//              pTobLigne.putValue(vStMois, FormatJour(0));
              pTobLigne.putValue(vStMois, 0);
              
            // calcul de la quantité affectée
            Result := Result + vTobMois.detail[i].getvalue(pStChamp);
          end;
      end;

  Finally
    if vQrMois <> nil then Ferme(vQrMois);
    vTobMois.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 02/07/2002
Modifié le ... :   /  /
Description .. : Calcul des quantites des mois provenant du cra
               : pour une ligne
Mots clefs ... :
*****************************************************************}
procedure LoadMoisCra(var pTobLigne : Tob; pStAffaire : String);
var
  vSt           : String;
  vQR           : TQuery;
  vTobAct       : Tob;
  i             : Integer;
  vWdYear       : Word;
  vWdMonth      : Word;
  vWdDay        : Word;
  vDtNow        : TDateTime;
  vStMois       : String;
  vDtDepart     : TDateTime;

Begin

  vDtDepart := DebutDeMois(now);

  vSt := vSt + ' SELECT ACT_DATEACTIVITE, SUM(ACT_QTE) As PlanifieCra ';
  vSt := vSt + ' FROM AFFAIRE, ACTIVITE ';
  vSt := vSt + ' WHERE ACT_ACTIVITEREPRIS = "A"';

  if pStAffaire <> '' then vSt := vSt + ' AND ACT_AFFAIRE = "' + pStAffaire + '"';

    // C.B 03/11/2003
  if pTobLigne.GetNumChamp('ATA_RESSOURCE') <> -1 then
    vSt := vSt + ' and ACT_RESSOURCE = "' + pTobLigne.GetValue('ATA_RESSOURCE') + '"';

  vSt := vSt + ' AND ACT_AFFAIRE = AFF_AFFAIRE ';
  vSt := vSt + ' AND ACT_TYPEARTICLE = "PRE" ';
  vSt := vSt + ' and ACT_DATEACTIVITE >= "' + UsDateTime(vDtDepart) +'"';
  vSt := vSt + ' AND ACT_ETATVISA="VIS" ';
                                      
  vSt := vSt + ' GROUP BY ACT_AFFAIRE, ACT_AFFAIRE0, ACT_AFFAIRE1, ';
  vSt := vSt + ' ACT_AFFAIRE2, ACT_AFFAIRE3, ACT_AVENANT,  ACT_ARTICLE,';
  vSt := vSt + ' ACT_RESSOURCE, ACT_DATEACTIVITE';


  vTobAct := TOB.create('Tob Viewer Mother', nil, -1);

  vQr := nil;
  Try
    vQR := OpenSql(vSt,True);
    if Not vQR.Eof then
      begin
        vTobAct.LoadDetailDB('MONTOBVIEWER','','',vQR,False,True);

        for i := vTobAct.Detail.count -1 downto 0 do
          begin
                              
            // calcul du mois
            // si 1er du mois, on se positionne le 2,
            // pour eviter d'avoir 31 jours au lien d'un mois lors de la difference entre les 2 dates
            DecodeDate(now, vWdYear, vWdMonth, vWdDay);
            if vWdDay = 1 then vDtNow := now + 1 else vDtNow := now;
            DecodeDate(strToDate(vTobAct.detail[i].GetValue('ACT_DATEACTIVITE')) - vDtNow, vWdYear, vWdMonth, vWdDay);
            if vWdMonth <> 12 then
              vStMois := 'PC_MOIS' + IntToStr(vWdMonth)
            else
              vStMois := 'PC_MOIS';

//            if FormatJour(vTobAct.detail[i].getvalue('PlanifieCra')) <> '0' then
//              pTobLigne.putValue(vStMois + 'CRA', 1);//FormatJour(StrToFLoat(vTobAct.detail[i].getvalue('PlanifieCra'))));
            if vTobAct.detail[i].getvalue('PlanifieCra') <> 0 then
              pTobLigne.putValue(vStMois + 'CRA', StrToFLoat(vTobAct.detail[i].getvalue('PlanifieCra')));
          end;
      end;
  Finally
    if vQR <> Nil then Ferme(vQR);
    vTobAct.Free;
  End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 15/05/2002
Modifié le ... :   /  /
Description .. : Teste la validité de la donnée en grille
Mots clefs ... :
*****************************************************************}
function TobStrToFloat(pVaQte : Variant) : Double;
begin
  // C.B 22/10/02
  try
    Result := valeur(pVaQte);
  except
    result := 0;                                 
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 04/07/2002
Modifié le ... :   /  /
Description .. : remplace une sous chaine par une autre dans une chaine
Mots clefs ... :
*****************************************************************}
Procedure ReplaceSubStr(Var pSt : String; pStSup, pStRep : string);
Var
  lInPos, lInLen : Integer;
begin
  lInLen := Length(pStSup);
  While true do begin
    lInPos := Pos(pStSup, pSt);
    if (lInPos > 0) then
      begin
        System.Delete(pSt, lInPos, lInLen);
        Insert(pStRep, pSt, lInPos);
      end
    else
      Exit;
  end;
end;

function EtatModifAutorisee(pStEtat : String) : boolean;
begin
  if pos(pStEtat, getParamSoc('SO_AFETATINTERDIT')) <> 0 then
    result := false
  else
    result := true;
end;

function EtatInterdit(pInMode : Integer; pStAffaire : String; pTache : RecordTache) : Boolean;
var
  vStEtats : String;
  vStEtat  : String;
  vSt      : String;

begin
  result := false;
  if pInMode <> cInAjout then
  begin
    vStEtats := getParamSoc('SO_AFETATINTERDIT');

    vStEtat := (Trim(ReadTokenSt(vStEtats)));
    While (vStEtat <> '') do
    Begin
      vSt := 'SELECT APL_ETATLIGNE FROM AFPLANNING WHERE APL_ETATLIGNE = "';
      vSt := vSt + vStEtat + '" AND APL_AFFAIRE = "' + pStAffaire + '" ';
      if pTache.StNumeroTache <> '' then vSt := vSt + 'AND APL_NUMEROTACHE = ' + pTache.StNumeroTache;
      if ExisteSql(vSt) then
      begin
        If (PGIAskAF(format('Ce planning contient des éléments dans l''état %s. Voulez vous vraiment régénerer ce planning ?',
                            [rechdom('AFTETAT', vStEtat, false)]),'')= mrYes) then
        begin
          result := false;
          break;
        end;
        begin
          result := true;
          break;
        end;
      end;
      vStEtat := (Trim(ReadTokenSt(vStEtats)));
    end;
  end;
end;


end.
