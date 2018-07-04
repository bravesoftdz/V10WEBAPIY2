{***********UNITE*************************************************
Auteur  ...... : F.Vautrain 
Créé le ...... : 01/12/2016
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTETATSTKPAYE ()
Mots clefs ... : TOF;BTETATSTKPAYE
*****************************************************************}
Unit BTEtatStkPaye ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul,
{$ENDIF}
      uTob,
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     UtilsEtat,
     UTOF ;

Type
  TOF_BTETATSTKPAYE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  Private
    TobStkPaye    : TOB;
    //
    Article       : THEdit;
    Article_      : THEdit;
    Famille1      : THValComboBox;
    Famille2      : THValComboBox;
    Famille3      : THValComboBox;
    Depot         : THValComboBox;
    DateFin       : THEdit;
    //
    Idef          : Integer;
    //
    //Variable névcessaire pour la gestion de l'état
    BParamEtat    : TToolBarButton97;
    //
    ChkApercu     : TCheckBox;
    ChkReduire    : TCheckBox;
    //
    FETAT         : THValComboBox;
    TEtat         : ThLabel;
    //
    Pages         : TPageControl;
    //
    OptionEdition : TOptionEdition;
    TheType       : String;
    TheNature     : String;
    TheTitre      : String;
    TheModele     : String;
    //
    function  CalculEntrees(CodeArticle : String; DateDeb : TdateTime) : Double;
    procedure ChargeEtat;
    procedure Controlechamp(Champ, Valeur: String);
    Procedure GetObjects;
    procedure SetScreenEvents;
    procedure EditeEtat;
    procedure OnChangeEtat(Sender: Tobject);
    procedure OnClickApercu(Sender: Tobject);
    procedure OnClickReduire(Sender: Tobject);
    procedure ParamEtat(Sender: TOBJect);

  end ;

Implementation

uses DateUtils;

procedure TOF_BTETATSTKPAYE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTETATSTKPAYE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTETATSTKPAYE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTETATSTKPAYE.OnLoad ;
Var StSQL       : string;
    QQ          : TQuery;
    Ind         : Integer;
    DateDeb     : TDateTime;
    TobLStkPaye : Tob;
    StkInitial  : Double;
    Stkentree   : Double;
    StkDiff     : Double;
    Pmap        : Double;
    Dpa         : Double;
    DD          : Word;
    YY          : Word;
    MM          : Word;
begin
  Inherited ;

  if (Article.Text = '') AND (Article_.Text = '') then
  begin
    if PGIAsk('Etes-vous sûr de vouloir lancer l''impression sur l''ensemble des articles tenue en Stock de la société ?', 'Etat des WStocks Payés')=Mrno then Exit;
  end;

  TobStkPaye := Tob.Create('STKPAYE', nil, -1);

  //calcul de la date de début d'entrée...
  //si le 15/12/2016 on cherche le début du mois, on enlève un mois et on se positionne au 15 du mois précédent à savoir le 15/10/2016
  DecodeDate(StrToDate(DateFin.Text), YY,MM,DD);
  DateDeb := IncDay(StrToDate(Datefin.text), DD * -1);
  DateDeb := IncMonth(DateDeb, -1);
  DateDeb := IncDay(DateDeb, -16);
  Decodedate(DateDeb, YY, MM, DD);
  DateDeb := EncodeDate(YY, MM, 15);

  StSql :=         'SELECT DISTINCT GA_ARTICLE, GA_CODEARTICLE, GA_LIBELLE, GDE_DEPOT, GDE_LIBELLE, ';
  StSql := StSQL + 'GQ_PHYSIQUE AS STKINITIAL, 0 AS STKENTREE, 0 AS STKDIFF, GQ_DPA, GQ_PMAP, 0 AS STKPPMAP, 0 AS STKPDPA, 0 AS STKNPPMAP, 0 AS STKNPDPA';
  StSql := StSQL + '  FROM DISPO LEFT JOIN ARTICLE ON GQ_ARTICLE=GA_ARTICLE ';
  StSql := StSQL + '  LEFT JOIN DEPOTS ON GDE_DEPOT=GQ_DEPOT ';
  StSql := StSQL + ' WHERE GA_TENUESTOCK="X" AND GA_STATUTART="UNI"';

  if Depot.Value <> ''      then StSQL := StSQL + ' AND GQ_DEPOT="' + Depot.Value + '"';

  if (Article.Text <> '')   then StSql := StSql + ' AND GA_CODEARTICLE >= "'+ Article.Text+ '"';

  if (Article_.Text <> '')  then StSql := StSql + ' AND GA_CODEARTICLE <= "' + Article_.Text + '"';

  if (Famille1.Value) <> '' then StSql := StSql + ' AND GA_FAMILLENIV1 = "' + Famille1.Value + '"';

  if (Famille2.Value) <> '' then StSql := StSql + ' AND GA_FAMILLENIV2 = "' + Famille2.Value + '"';

  if (Famille3.Value) <> '' then StSql := StSql + ' AND GA_FAMILLENIV3 = "' + Famille3.Value + '"';

  If Depot.Value <> ''      then
    StSql := StSQL + ' ORDER BY GDE_DEPOT, GA_ARTICLE'
  else
    StSql := StSQL + ' ORDER BY GA_ARTICLE';

  QQ := OpenSQL(StSQL, False, -1, '', True);

  if not QQ.Eof then
  begin
    TobStkPaye.LoadDetailFromSQL(StSQL, True, True);
    for Ind := 0 to TobStkPaye.detail.count -1 do
    begin
      TOBLStkPaye := TobStkPaye.detail[Ind];
      //
      StkEntree   := CalculEntrees(TobLStkPaye.GetValue('GA_CODEARTICLE'),DateDeb);
      StkInitial  := TobLStkPaye.GetValue('STKINITIAL');
      Stkdiff     := (StkInitial - Stkentree);
      If Stkdiff < 0 then Stkdiff := 0;
      Pmap        := TobLStkPaye.GetValue('GQ_PMAP');
      Dpa         := TobLStkPaye.GetValue('GQ_DPA');
      //
      TOBLStkPaye.PutValue('STKENTREE', Stkentree);
      TOBLStkPaye.PutValue('STKDIFF',   StkDiff);
      TobLStkPaye.PutValue('STKPPMAP', (StkDiff * Pmap));
      TobLStkPaye.PutValue('STKPDPA',  (StkDiff * Dpa));
      TobLStkPaye.PutValue('STKNPPMAP',(Stkentree * Pmap));
      TobLStkPaye.PutValue('STKNPDPA', (Stkentree * Dpa));
    end;
  end;

  Ferme(QQ);

  EditeEtat;

  FreeAndNil(TobStkPaye);

end ;

procedure TOF_BTETATSTKPAYE.OnArgument (S : String ) ;
var i       : Integer;
    Critere : string;
    Champ   : string;
    Valeur  : string;
begin
  Inherited ;

  GetObjects;

  Critere := S;
  //
  While (Critere <> '') do
  BEGIN
    i:=pos(':',Critere);
    if i = 0 then i:=pos('=',Critere);
    if i <> 0 then
       begin
       Champ:=copy(Critere,1,i-1);
       Valeur:=Copy (Critere,i+1,length(Critere)-i);
       end
    else
       Champ := Critere;
    Controlechamp(Champ, Valeur);
    Critere:=(Trim(ReadTokenSt(S)));
  END;

  //Gestion des évènement des zones écran
  SetScreenEvents;

  ChargeEtat;

  //Article.plus  := 'GA_STATUTART="UNI" AND GA_TENUESTOCK="X"';
  //Article_.plus := 'GA_STATUTART="UNI" AND GA_TENUESTOCK="X"';
   
end ;

procedure TOF_BTETATSTKPAYE.OnClose ;
begin
  Inherited ;

  FreeAndNil(OptionEdition);

end ;

procedure TOF_BTETATSTKPAYE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTETATSTKPAYE.OnCancel () ;
begin
  Inherited ;
end ;

Function TOF_BTETATSTKPAYE.CalculEntrees(CodeArticle : String; DateDeb: TdateTime) : Double;
Var StSql     : String;
    QQ        : TQuery;
begin

  Result := 0;

  {*
  StSql :=         'SELECT GL_QTESTOCK ';
  StSql := StSQL + '  FROM LIGNE LEFT JOIN PARPIECE ON GL_NATUREPIECEG=GPP_NATUREPIECEG ';
  StSql := StSQL + ' WHERE GL_NATUREPIECEG IN ("BLF","FF","LFR") AND NOT (GL_IDENTIFIANTWOL=-1 AND GL_AFFAIRE <>"") ';
  StSql := StSQL + '   AND GL_CODEARTICLE = "' + CodeArticle + '"';
  if Depot.value <> '' then StSql := StSQL + ' AND GL_DEPOT = "' + Depot.Value + '" ';
  StSql := StSql + '   AND GL_DATEPIECE > "' + USDATETIME(DateDeb) + '" ';
  StSql := StSQL + ' UNION ALL ';

  StSql := StSQL + 'SELECT GL_QTESTOCK ';
  StSql := StSql + '  FROM LIGNE LEFT JOIN PARPIECE ON GL_NATUREPIECEG=GPP_NATUREPIECEG ';
  StSql := StSql + ' WHERE GL_NATUREPIECEG IN ("AFS","BFA") ';
  StSQL := StSQL + '   AND GL_CODEARTICLE = "' + CodeArticle + '"';
  if Depot.value <> '' then StSql := StSQL + ' AND GL_DEPOT = "' + Depot.value + '" ';
  StSql := StSql + '   AND GL_DATEPIECE > "' + USDATETIME(DateDeb) + '" ';
  StSql := StSQL + ' UNION ALL ';

  StSql := StSQL + 'SELECT GL_QTESTOCK ';
  StSql := StSQL + '  FROM LIGNE LEFT JOIN PARPIECE ON GL_NATUREPIECEG=GPP_NATUREPIECEG ';
  StSql := StSQL + ' WHERE GL_NATUREPIECEG IN ("EEX","TRE") ';
  StSQL := StSQL + '   AND GL_CODEARTICLE = "' + CodeArticle + '"';
  if Depot.value <> '' then StSql := StSQL + ' AND GL_DEPOT = "' + Depot.value + '" ';
  StSql := StSql + '   AND GL_DATEPIECE > "' + USDATETIME(DateDeb) + '" ';
  *}

  //Un select au lieu de trois....
  StSQL := 'SELECT ';
  StSql := StSQL + '(SELECT SUM(GL_QTESTOCK) ';
  StSql := StSQL + '  FROM LIGNE LEFT JOIN PARPIECE ON GL_NATUREPIECEG=GPP_NATUREPIECEG ';
  StSql := StSQL + ' WHERE GL_NATUREPIECEG IN ("BLF","FF","LFR") AND NOT (GL_IDENTIFIANTWOL=-1 AND GL_AFFAIRE <>"") AND GL_VIVANTE="X" ';
  StSql := StSQL + '   AND GL_CODEARTICLE = "' + CodeArticle + '"';
  if Depot.value <> '' then StSql := StSQL + ' AND GL_DEPOT = "' + Depot.Value + '" ';
  StSql := StSql + '   AND GL_DATEPIECE > "' + USDATETIME(DateDeb) + '") AS STKENTREE, ';

  StSql := StSQL + '(SELECT SUM(GL_QTESTOCK) ';
  StSql := StSql + '  FROM LIGNE LEFT JOIN PARPIECE ON GL_NATUREPIECEG=GPP_NATUREPIECEG ';
  StSql := StSql + ' WHERE GL_NATUREPIECEG IN ("AFS","BFA") ';
  StSQL := StSQL + '   AND GL_CODEARTICLE = "' + CodeArticle + '"';
  if Depot.value <> '' then StSql := StSQL + ' AND GL_DEPOT = "' + Depot.value + '" ';
  StSql := StSql + '   AND GL_DATEPIECE > "' + USDATETIME(DateDeb) + '") AS STKENTREE1, ';

  StSql := StSQL + 'SUM(GL_QTESTOCK) AS STKENTREE2 ';
  StSql := StSQL + '  FROM LIGNE LEFT JOIN PARPIECE ON GL_NATUREPIECEG=GPP_NATUREPIECEG ';
  StSql := StSQL + ' WHERE GL_NATUREPIECEG IN ("EEX","TRE") ';
  StSQL := StSQL + '   AND GL_CODEARTICLE = "' + CodeArticle + '"';
  if Depot.value <> '' then StSql := StSQL + ' AND GL_DEPOT = "' + Depot.value + '" ';
  StSql := StSql + '   AND GL_DATEPIECE > "' + USDATETIME(DateDeb) + '" ';

  QQ := OpenSQL(StSql, False, -1, '', True);

  If not QQ.eof then
  begin
    //QQ.first;
    //repeat
    Result := QQ.Findfield('STKENTREE').AsFloat + QQ.Findfield('STKENTREE1').AsFloat + QQ.Findfield('STKENTREE2').AsFloat;
    //  QQ.next;
    //until QQ.eof;
  end;

  ferme(QQ);

end;

procedure TOF_BTETATSTKPAYE.GetObjects;
begin
  //
  Pages         := TPageControl(GetControl('PAGES'));
  //
  Article       := THEdit(GetControl('GA_CODEARTICLE'));
  Article_      := THEdit(GetControl('GA_CODEARTICLE_'));
  Depot         := THValComboBox(GetControl('DEPOT'));
  DateFin       := THEdit(GetControl('DATEFIN'));
  //
  Famille1      := ThValComboBox(Getcontrol('GA_FAMILLENIV1'));
  Famille2      := ThValComboBox(Getcontrol('GA_FAMILLENIV2'));
  Famille3      := ThValComboBox(Getcontrol('GA_FAMILLENIV3'));
  //
  BParamEtat    := TToolbarButton97(GetCONTROL('BPARAMETAT'));
  TEtat         := ThLabel(ecran.FindComponent('TEtat'));
  FEtat         := ThValComboBox(ecran.FindComponent('FEtat'));
  //
  ChkApercu     := TCheckBox(Ecran.FindComponent('fApercu'));
  ChkReduire    := TCheckBox(Ecran.FindComponent('fReduire'));
  //
end;

Procedure TOF_BTETATSTKPAYE.Controlechamp(Champ, Valeur : String);
begin

end;

procedure TOF_BTETATSTKPAYE.SetScreenEvents;
begin
  //
  if assigned(BParamEtat)   then BParamEtat.OnClick    := ParamEtat;
  if assigned(ChkApercu)    then ChkApercu.OnClick    := OnClickApercu;
  if assigned(ChkReduire)   then ChkReduire.OnClick   := OnClickReduire;
  if assigned(FETAT)        then FETAT.OnChange       := OnChangeEtat;    
  //
end;

procedure TOF_BTETATSTKPAYE.ChargeEtat;
begin

  TheType   := 'E';
  TheNature := 'B00';
  TheModele := 'BSP';
  TheTitre  := 'Etats des Stocks Payes';

  OptionEdition := TOptionEdition.Create(TheType, TheNature, TheModele, TheTitre, '', ChkApercu.Checked, ChkReduire.Checked, True, False, False, Pages, fEtat);

  OptionEdition.first := True;
  OptionEdition.ChargeListeEtat(fEtat,Idef);

  if fetat.itemindex  >= 0 then TheModele := FETAT.values[fetat.itemindex];

end;

procedure TOF_BTETATSTKPAYE.ParamEtat(Sender : TOBJect);
begin

  OptionEdition.Appel_Generateur

end;

procedure TOF_BTETATSTKPAYE.OnClickApercu(Sender : Tobject);
begin
  OptionEdition.Apercu  := ChkApercu.checked;
end;

procedure TOF_BTETATSTKPAYE.OnClickReduire(Sender : Tobject);
begin
  OptionEdition.DeuxPages  := ChkReduire.checked;
end;

Procedure TOF_BTETATSTKPAYE.OnChangeEtat(Sender : Tobject);
Begin

   OptionEdition.Modele := FETAT.values[fetat.itemindex];

   TheModele := OptionEdition.Modele;

end;

procedure TOF_BTETATSTKPAYE.EditeEtat;
begin

  if (TobStkPaye=nil) Or (TobStkPaye.detail.count = 0) then
  Begin
    PGIInfo('Aucun enregistrements ne correspondent à votre sé&lection', 'Etat des Stocks Payés');  
    exit;
  end;

  OptionEdition.Apercu    := ChkApercu.Checked;
  OptionEdition.DeuxPages := ChkReduire.Checked;
  OptionEdition.Spages    := Pages;

  if OptionEdition.LanceImpression('', TobStkPaye) < 0 then V_PGI.IoError:=oeUnknown;

end;

Initialization
  registerclasses ( [ TOF_BTETATSTKPAYE ] ) ; 
end.

