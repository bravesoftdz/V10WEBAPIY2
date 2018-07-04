unit ULibIdentBancaire;

interface

uses StdCtrls,
  Controls,
  Classes,
  graphics,
  {$IFNDEF EAGLCLIENT}
  db,DBCtrls,
  {$IFNDEF DBXPRESS}dbtables, {$ELSE}uDbxDataSet, {$ENDIF}
{$IFNDEF EAGLSERVER}
  Fiche,
  FichList,
  FE_Main, // AGLLanceFiche
{$ENDIF}
  {$ELSE}
  eFiche,
  eFichList,
  MaineAGL, // AGLLanceFiche
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  HTB97,
{$IFNDEF EAGLSERVER}
  wCommuns,
  UTOF,
  uTobDebug, // TobDebug
{$ENDIF}
  UTob,
  HDB,uEntCommun;

function ParametreRIB(FF : Tform ;IsoPays: string; TYPEPAYS, prefixe : string;  var TRIB : TOB;affiche : boolean=true; TabOrder : integer=0) : String ;
function chercheLEQUEL(pos: integer; TobIdent: tob): string;
function cherchePOS(lequel : string ; TobIdent: tob): integer;
function chercheLONG(lequel,IsoPays,TypePays : string): integer;
function TraduitCode(st: string): string;
function existeLEQUEL(lequel : string ; PaysIso : string ; TypePays : string): boolean ;
function existeParamPays(PaysIso : string; TypePays : string='VIDE'): boolean ;
function CountParamPays(PaysIso : string): integer ;
function GetParamPays(PaysIso : string): string ;
function GetLibelleParamPays(PaysIso,TypePays : string): string ;
function quelLongChamps(IsoPays, TYPEPAYS, Champ : string) : Integer ;
function afficheRIB(FF: TForm; gauche: integer; prefixe: string; TobIdent: TOB;TabOrder : integer=0): string;
function afficheBBAN(FF : TForm; gauche : integer; prefixe: string; TRIB : Tob): string;
function quelType(IsoPays: string; TYPEPAYS : string) : String ;
procedure Affiche ( FF : TForm; PaysIso, TypePays, prefixe : string ;TabOrder: integer=0 ) ;
procedure AfficheTYPEPAYS(FF : TForm; DS:TDataSet; prefixe, PaysIso:string) ;
procedure AfficheDefault(FF : TForm; prefixe,PaysISO :string; Active : Boolean;TabOrder : integer=0) ;
procedure SaveAffichageDefaut (FF : TForm; prefixe : string) ;
procedure DestroySaveAffichage ;
function GetMask(Long : integer) : string  ;

var
{$IFNDEF EAGLCLIENT}
	THDB_BanqueSav,THDB_GuichetSav,THDB_CleSav,THDB_CompteSav,THDB_CodeIbanSav: THDBEdit;
{$ELSE}
	THDB_BanqueSav,THDB_GuichetSav,THDB_CleSav,THDB_CompteSav,THDB_CodeIbanSav: THEdit;
{$ENDIF}
THLabel_BanqueSav,THLabel_GuichetSav,THLabel_CleSav,THLabel_CompteSav: THLabel;

implementation

uses
  Ent1,
  {$IFDEF MODENT1}
  CPTypeCons,
  CPVersion,
  {$ENDIF MODENT1}
  paramsoc;


{***********A.G.L.***********************************************
Auteur  ...... : Regis Rohault
Créé le ...... : 20/11/2006
Modifié le ... :   /  /    
Description .. : Retourne le type de Compte selon le type de pays
Mots clefs ... :
*****************************************************************}
function quelType(IsoPays: string; TYPEPAYS : string) : String ;
var
  Q : TQuery ;
begin
  if length(trim(isopays))=0 then exit ;
try
  Q := OpenSQL('SELECT YIB_TYPECOMPTE FROM YIDENTBANCAIRE WHERE YIB_PAYSISO="' + isopays + '" AND YIB_TYPEPAYS="' + TYPEPAYS+'"', true) ;
  if Not Q.EOF then result := Q.Fields[0].AsString ;
  ferme(q);
except ;
  result := 'BBA' ;
end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 04/01/2007
Modifié le ... : 04/01/2007
Description .. : Retourne la longueur du champs demandé pour un 
Suite ........ : paramétrage donné
Mots clefs ... : 
*****************************************************************}
function quelLongChamps(IsoPays,TYPEPAYS,Champ : string) : Integer ;
var
  Q : TQuery ;
begin       
  result := 0;
  if length(trim(isopays))=0 then exit ;
  if (Champ = 'CODEIBAN') and ((QuelType(IsoPays,TYPEPAYS) = 'RIB')) then
  begin
      // La longueur n'est pas renseigné on la calcul :
      Q := OpenSQL('SELECT SUM(YIB_LGETABBQ+YIB_LGGUICHET+YIB_LGNUMEROCOMPTE+YIB_LGCLERIB) FROM YIDENTBANCAIRE WHERE YIB_PAYSISO="' + isopays + '" AND YIB_TYPEPAYS="' + TYPEPAYS+'"', true) ;
      if Not Q.EOF then result := Q.Fields[0].AsInteger + 4 ;
         ferme(Q);
  end
  else if (Champ = 'CODEIBAN') and (QuelType(IsoPays,TYPEPAYS) = '') then
  begin
  {$IFDEF NOVH}
       if (GetParamSocSecur('SO_PAYSLOCALISATION','') = CodeISOES) and (IsoPays = CodeISOES) then
  {$ELSE}
       if (VH^.PaysLocalisation = CodeISOES) and (IsoPays = CodeISOES) then
  {$ENDIF}
          // On est en Espagne
          result := 24
       else
          // On est en France ou par default
          result := 27;
  end
  else if (Champ = 'ETABBQ') or (Champ = 'GUICHET') or (Champ = 'NUMEROCOMPTE') or  (Champ = 'CLERIB') or  (Champ = 'CODEIBAN') then
  begin
    Q := OpenSQL('SELECT YIB_LG' + Champ + ' FROM YIDENTBANCAIRE WHERE YIB_PAYSISO="' + isopays + '" AND YIB_TYPEPAYS="' + TYPEPAYS+'"', true) ;
    if Not Q.EOF then result := Q.Fields[0].AsInteger ;
      ferme(Q);
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Regis Rohault
Créé le ...... : 20/11/2006
Modifié le ... :   /  /    
Description .. : Charge la TOB avec les données correspondant au Type et 
Suite ........ : au Pays.
Suite ........ : Affiche le RIB/IBAN/BBAN si besoin (parametre affiche).
Suite ........ : Retourne les libellés des composants dans l'ordre 
Suite ........ : d'affichage. ou le mot cle 'VIDE' pour reinitialiser
Suite ........ : les positions des params RIB
Mots clefs ... :
*****************************************************************}
function ParametreRIB(FF : Tform ;IsoPays: string; TYPEPAYS, prefixe : string;  var TRIB : TOB ;affiche : boolean=true;TabOrder :integer=0) : String ;
var
  lRIB : string ;
  OK : boolean ;
  Q : TQuery ;
begin
  OK := false ;
  if length(trim(isopays))=0 then exit ;
try
  Q := OpenSQL('select * from yidentbancaire where yib_paysiso="' + isopays + '" and yib_TYPEPAYS="' + TYPEPAYS+'"', true) ;
  if Not Q.EOF then  OK := TRIB.SelectDB('', Q, True) ;
  ferme(q);
  if not OK then
  begin
    //MESSAGEALERTE('ATTENTION : Paramétrage d''indentification bancaire inconnue');
    result := 'VIDE'
  end
  else
  begin
    if affiche then begin
      if TRIB.GetValue('YIB_TYPECOMPTE')<>'BBA' then
        lRIB := afficheRIB(FF , THLabel(FF.FindComponent('T' + prefixe + 'DOMICILIATION')).left, prefixe, TRIB, TabOrder)
      else
        lRIB := afficheBBAN(FF , THLabel(FF.FindComponent('T'+ prefixe + 'DOMICILIATION')).left, prefixe, TRIB);
      result := lRIB ;
    end ;
  end;
except ;
  result := '' ;
end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Regis Rohault
Créé le ...... : 20/11/2006
Modifié le ... :   /  /
Description .. : Retourne le nom du champs du numero d'ordre
Suite ........ : correspondant au type du composant RIB
Mots clefs ... :
*****************************************************************}
function cherchePOS(lequel : string ; TobIdent: tob): integer;
begin
try
  result := TobIdent.GetValue('YIB_N' + lequel) ;
except
  result := 0;
end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Regis Rohault
Créé le ...... : 20/11/2006
Modifié le ... :   /  /
Description .. : Retourne le nom du champs du numero d'ordre
Suite ........ : correspondant au type du composant RIB
Mots clefs ... :
*****************************************************************}
function chercheLONG(lequel,IsoPays,TypePays : string): integer;
var TobIdent : TOB;
    Q : TQuery;
begin
  result := 0;
  TobIdent := Tob.Create('YIDENTBANCAIRE',nil,-1);
  Q := OpenSQL('select * from yidentbancaire where yib_paysiso="' + isopays + '" and yib_TYPEPAYS="' + TYPEPAYS+'"', true) ;
  try
  if Not Q.EOF then
  begin
     TobIdent.SelectDB('', Q, True);
     result := TobIdent.GetValue('YIB_LG' + lequel) ;
  end;
finally
  ferme(q);
  freeandnil(tobident);
end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Regis Rohault
Créé le ...... : 20/11/2006
Modifié le ... :   /  /
Description .. : Retourne le type du composant RIB correspondant au 
Suite ........ : champs du numero d'ordre dans la TOB
Mots clefs ... : 
*****************************************************************}
function chercheLEQUEL(pos: integer; TobIdent: tob): string;
begin
try
  if TobIdent.GetValue('YIB_NETABBQ') = pos then
    result := 'ETABBQ';
  if TobIdent.GetValue('YIB_NGUICHET') = pos then
    result := 'GUICHET';
  if TobIdent.GetValue('YIB_NNUMEROCOMPTE') = pos then
    result := 'NUMEROCOMPTE';
  if TobIdent.GetValue('YIB_NCLERIB') = pos then
    result := 'CLERIB';
except
  result := '';
end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Regis Rohault
Créé le ...... : 20/11/2006
Modifié le ... :   /  /
Description .. : Indique si le composant du RIB est renseigné pour un Type
Suite ........ : de pays donné.
Mots clefs ... :
*****************************************************************}
function existeLEQUEL(lequel : string ; PaysIso : string ; TypePays : string): boolean ;
var SQL : string ;
		Q   : TQuery ;
begin  
	result := false;
  SQL := 'SELECT YIB_LG'+ lequel + ' from YIDENTBANCAIRE where YIB_PAYSISO="' + PaysIso + '" and YIB_TYPEPAYS="' + TypePays + '" and YIB_LG'+ lequel + '>0' ;
  // FQ 19472 BV 11/01/2007
  Q := OpenSQL(SQL, true) ;
  try
 	 if Not Q.EOF then
 	 begin
   	result := (Q.Fields[0].AsInteger > 0);
 	 end;
	finally
  	ferme(q);
	end;
  // END FQ 19472
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/11/2006
Modifié le ... :   /  /
Description .. : Indique s'il existe un paramétrage pour le pays donné
Suite ........ :
Mots clefs ... :
*****************************************************************}
function existeParamPays(PaysIso : string; TypePays : string='VIDE'): boolean ;
var SQL : string ;
begin
  if TypePays = 'VIDE' then
    SQL := 'SELECT * from YIDENTBANCAIRE where YIB_PAYSISO="' + PaysIso + '"'
  else
    SQL := 'SELECT * from YIDENTBANCAIRE where YIB_PAYSISO="' + PaysIso + '" AND YIB_TYPEPAYS="' + TypePays + '"';
  result := ExisteSQL(SQL);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 04/01/2007
Modifié le ... :   /  /
Description .. : Indique le nombre d'identification pour un pays
Suite ........ :
Mots clefs ... :
*****************************************************************}
function CountParamPays(PaysIso : string): integer ;
var SQL : string ;
    Q   : TQuery ;
begin
  result := -1;
  SQL := 'SELECT COUNT(*) from YIDENTBANCAIRE where YIB_PAYSISO="' + PaysIso + '"';
  Q := OpenSQL(SQL, true) ;
  try
  if Not Q.EOF then
  begin
     result := Q.Fields[0].AsInteger;
  end;
finally
  ferme(q);
end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 04/01/2007
Modifié le ... : 04/01/2007
Description .. : Retourne le caractere du premier parametrage trouvé pour
Suite ........ : un pays donné.
Suite ........ : 
Mots clefs ... :
*****************************************************************}
function GetParamPays(PaysIso : string): string ;
var SQL : string ;
    Q   : TQuery ;
begin
  SQL := 'SELECT  ##TOP 1## YIB_TYPEPAYS from YIDENTBANCAIRE where YIB_PAYSISO="' + PaysIso + '"';
  Q := OpenSQL(SQL, true) ;
  try
  if Not Q.EOF then
  begin
     result := Q.Fields[0].AsString;
  end;
finally
  ferme(q);
end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 04/01/2007
Modifié le ... :   /  /    
Description .. : Retourne le libelle du parametre pays
Mots clefs ... : 
*****************************************************************}
function GetLibelleParamPays(PaysIso,TypePays : string): string ;
var SQL : string ;
    Q   : TQuery ;
begin
  SQL := 'SELECT YIB_LIBELLE from YIDENTBANCAIRE where YIB_PAYSISO="' + PaysIso + '" AND YIB_TYPEPAYS="' + TypePays + '"';
  Q := OpenSQL(SQL, true) ;
  try
  if Not Q.EOF then
  begin
     result := Q.Fields[0].AsString;
  end;
finally
  ferme(q);
end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Regis Rohault
Créé le ...... : 20/11/2006
Modifié le ... : 20/11/2006
Description .. : Affiche les composants BBAN sur la fiche
Suite ........ : Retourne vide
Mots clefs ... : 
*****************************************************************}
function AfficheBBAN(FF : TForm; gauche : integer; prefixe: string; TRIB : Tob): string;
var
  lgMAX,lgMaxIBAN : integer ;
begin
    if TRIB.GetValue('YIB_TYPECOMPTE') = 'BBA' then
    begin
{$IFNDEF EAGLCLIENT}
      THDBEDIT(FF.FindComponent(prefixe + 'ETABBQ')).Enabled:= false ;
      THDBEDIT(FF.FindComponent(prefixe + 'GUICHET')).Enabled:= false ;
      THDBEDIT(FF.FindComponent(prefixe + 'NUMEROCOMPTE')).Enabled:= false ;
      THDBEDIT(FF.FindComponent(prefixe + 'CLERIB')).Enabled:= false ;
{$ELSE}
      THEDIT(FF.FindComponent(prefixe + 'ETABBQ')).Enabled:= false ;
      THEDIT(FF.FindComponent(prefixe + 'GUICHET')).Enabled:= false ;
      THEDIT(FF.FindComponent(prefixe + 'NUMEROCOMPTE')).Enabled:= false ;
      THEDIT(FF.FindComponent(prefixe + 'CLERIB')).Enabled:= false ;
{$ENDIF}
      THLabel(FF.FindComponent('T' + prefixe + 'CODEIBAN')).Caption := TraduireMemoire('Code BBAN') ;
{$IFNDEF EAGLCLIENT}
      THDBEdit(FF.FindComponent(prefixe + 'CODEIBAN')).Enabled := True ;
      THDBEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).Hint := TraduireMemoire('Code BBAN') ;
{$ELSE}
      THEdit(FF.FindComponent(prefixe + 'CODEIBAN')).Enabled := True ;
      THEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).Hint := TraduireMemoire('Code BBAN') ;
{$ENDIF}
      lgMaxIBAN := StrToInt(TRIB.GetValue('YIB_LGCODEIBAN'));
{$IFNDEF EAGLCLIENT}
      THDBEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).MaxLength := lgMaxIBAN ;
{$ELSE}
      THEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).MaxLength := lgMaxIBAN ;
{$ENDIF}
      lgMax := FF.Canvas.TextWidth(StringOfChar('W', lgMaxIBAN))+9 ;

{$IFNDEF EAGLCLIENT}
      THDBEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).Width := lgMax ;
{$ELSE}
      THEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).Width := lgMax ;
{$ENDIF}

      result := '&' + TraduireMemoire('Banque');
    end ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Regis Rohault
Créé le ...... : 20/11/2006
Modifié le ... :   /  /
Description .. : Affiche les composants RIB sur la fiche
Suite ........ : Retourne le libellé des composants existant dans l'ordre.
Mots clefs ... : 
*****************************************************************}
function afficheRIB(FF: TForm; gauche: integer; prefixe: string; TobIdent: TOB;TabOrder : integer=0): string;
var
  lequel: string;
  i, max: integer;
  libel : string;
  lgMax, lgMaxIBAN : integer ;
  stMax : string ;
  affiche : boolean;
begin
  if TobIdent = nil then
    exit;
  // Initialisation :
  max := 0 ; 
  lgMaxIBAN := 0 ;
  // Cas specifique Treso pour les comptes généraux :
  if prefixe = 'BQ_' then
     affiche := not EstComptaTreso
  else
     affiche := true;
{$IFNDEF EAGLCLIENT}
  THDBEDIT(FF.FindComponent(prefixe + 'ETABBQ')).Enabled:= affiche;
  THDBEDIT(FF.FindComponent(prefixe + 'GUICHET')).Enabled:= affiche;
  THDBEDIT(FF.FindComponent(prefixe + 'NUMEROCOMPTE')).Enabled:= true ;
  THDBEDIT(FF.FindComponent(prefixe + 'CLERIB')).Enabled:= true ;
  THDBEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).Enabled:= true ;
{$ELSE} 
  THEDIT(FF.FindComponent(prefixe + 'ETABBQ')).Enabled:= affiche;
  THEDIT(FF.FindComponent(prefixe + 'GUICHET')).Enabled:= affiche;
  THEDIT(FF.FindComponent(prefixe + 'NUMEROCOMPTE')).Enabled:= true ;
  THEDIT(FF.FindComponent(prefixe + 'CLERIB')).Enabled:= true ;
  THEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).Enabled:= true ;
{$ENDIF}


  // Affichage dynamique
  libel := '';
  THLabel(FF.FindComponent('T' + prefixe + 'GUICHET')).Visible := false ;
  THLabel(FF.FindComponent('T' + prefixe + 'NUMEROCOMPTE')).Visible := false ;
  THLabel(FF.FindComponent('T' + prefixe + 'CLERIB')).Visible := false ;
  //on récupère les infos du paramétrage

{$IFNDEF EAGLCLIENT}
  for i := 1 to 4 do
  begin
    lequel := chercheLEQUEL(i, tobIdent);
    if (length(trim(TobIdent.GetValue('YIB_LG' + lequel ))) <> 0) and (tobIdent.GetValue('YIB_LG'+lequel)<>0) then
    begin
      inc(max);
      FF.Canvas.Font := THDBEDIT(FF.FindComponent(prefixe + lequel)).Font;
      THDBEDIT(FF.FindComponent(prefixe + lequel)).MaxLength := StrToInt(TobIdent.GetValue('YIB_LG' + lequel));

      FF.Canvas.Font := THEDIT(FF.FindComponent(prefixe + lequel)).Font;
      THEDIT(FF.FindComponent(prefixe + lequel)).MaxLength := StrToInt(TobIdent.GetValue('YIB_LG' + lequel));
      lgMaxIBAN := lgMaxIBAN + StrToInt(TobIdent.GetValue('YIB_LG' + lequel)) ;
      stMax := StringOfChar('W', StrToInt(TobIdent.GetValue('YIB_LG' + lequel))+1) ; //+ 2) ;
      lgMax := FF.Canvas.TextWidth(stMax);
      THDBEDIT(FF.FindComponent(prefixe + lequel)).Width := lgMax;
      THDBEDIT(FF.FindComponent(prefixe + lequel)).Left := gauche;
      gauche := 5 +THDBEDIT(FF.FindComponent(prefixe + lequel )).Width + gauche;
      THDBEDIT(FF.FindComponent(prefixe + lequel )).TabOrder := i + TabOrder;
      THDBEDIT(FF.FindComponent(prefixe + lequel )).EditMask := stringofchar('9',StrToInt(TobIdent.GetValue('YIB_LG' + lequel))) ;
      if max = 1 then
        libel := TraduitCode(lequel)
      else
        libel := libel + ',' + TraduitCode(lequel);
    end
    else
    begin
      THDBEDIT(FF.FindComponent(prefixe + lequel)).visible := false;
    end;
    stMAx := StringOfChar('W', lgMaxIBAN ); //+ 6) ;
    lgMax := FF.Canvas.TextWidth(stMax);
    THDBEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).MaxLength := lgMaxIBAN ;
    THDBEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).Width := lgMax;
    THLabel(FF.FindComponent('T' + prefixe + 'CODEIBAN')).Caption := 'Code IBAN' ;
    THDBEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).Hint := TraduireMemoire('Code IBAN') ;
  end;
{$ELSE}
  for i := 1 to 4 do
  begin
    lequel := chercheLEQUEL(i, tobIdent);
    if (length(trim(TobIdent.GetValue('YIB_LG' + lequel ))) <> 0) and (tobIdent.GetValue('YIB_LG'+lequel)<>0) then
    begin
      inc(max);
      FF.Canvas.Font := THEDIT(FF.FindComponent(prefixe + lequel)).Font;
      THEDIT(FF.FindComponent(prefixe + lequel)).MaxLength := StrToInt(TobIdent.GetValue('YIB_LG' + lequel));

      FF.Canvas.Font := THEDIT(FF.FindComponent(prefixe + lequel)).Font;
      THEDIT(FF.FindComponent(prefixe + lequel)).MaxLength := StrToInt(TobIdent.GetValue('YIB_LG' + lequel));
      lgMaxIBAN := lgMaxIBAN + StrToInt(TobIdent.GetValue('YIB_LG' + lequel)) ;
      stMax := StringOfChar('W', StrToInt(TobIdent.GetValue('YIB_LG' + lequel))+1) ; //+ 2) ;
      lgMax := FF.Canvas.TextWidth(stMax);
      THEDIT(FF.FindComponent(prefixe + lequel)).Width := lgMax;
      THEDIT(FF.FindComponent(prefixe + lequel)).Left := gauche;
      gauche := 5 +THEDIT(FF.FindComponent(prefixe + lequel )).Width + gauche;
      THEDIT(FF.FindComponent(prefixe + lequel )).TabOrder := i + TabOrder;
      THEDIT(FF.FindComponent(prefixe + lequel )).EditMask := stringofchar('9',StrToInt(TobIdent.GetValue('YIB_LG' + lequel))) ;
      if max = 1 then
        libel := TraduitCode(lequel)
      else
        libel := libel + ',' + TraduitCode(lequel);
    end
    else
    begin
      THEDIT(FF.FindComponent(prefixe + lequel)).visible := false;
    end;
    stMAx := StringOfChar('W', lgMaxIBAN ); //+ 6) ;
    lgMax := FF.Canvas.TextWidth(stMax);
    THEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).MaxLength := lgMaxIBAN ;
    THEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).Width := lgMax;
    THLabel(FF.FindComponent('T' + prefixe + 'CODEIBAN')).Caption := 'Code IBAN' ;
    THEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).Hint := TraduireMemoire('Code IBAN') ;
  end;
{$ENDIF}
  result := libel
end;

function GetMask(Long : integer) : string ;
var i,j : integer;
begin
     result := '';
     j := 0;
     for i := Long downto 1 do
     begin
          Inc(j);
          result := result + 'a';
          if (j = 4) and (i > 1) then
          begin
               result := result + ' ';
               j := 0;
          end;
     end;
     result :=  result + ';0;_' ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Regis Rohault
Créé le ...... : 20/11/2006
Modifié le ... :   /  /
Description .. : 
Mots clefs ... :
*****************************************************************}
function CorrespondChamps(st: string): string;
begin
  /// préfixe BQ_, BT_, ILN_
  if UpperCase(st) = 'BANQUE' then
    result := 'ETABBQ';
  //préfixe ANL_
  if UpperCase(st) = 'BANQGUICHET' then
    result := 'GUICHET';
  if UpperCase(st) = 'BANQCODE' then
    result := 'ETABBQ';
  if UpperCase(st) = 'BANQGUICHET' then
    result := 'GUICHET';
  if UpperCase(st) = 'BANQCLERIB' then
    result := 'CLERIB';

  if UpperCase(st) = 'NUMEROCOMPTE' then
    result := TraduireMemoire('Numéro Compte');
  if UpperCase(st) = 'CLERIB' then
    result := TraduireMemoire('Clé');

end;


{***********A.G.L.***********************************************
Auteur  ...... : Regis Rohault
Créé le ...... : 20/11/2006
Modifié le ... :   /  /    
Description .. : Retourne le libellé dans la bonne langue en fonction du type
Suite ........ : de composant.
Mots clefs ... : 
*****************************************************************}
function TraduitCode(st: string): string;
begin
  if UpperCase(st) = 'ETABBQ' then
    result := '&' + TraduireMemoire('Banque');
  if UpperCase(st) = 'GUICHET' then
    result := TraduireMemoire('Guichet');
  if UpperCase(st) = 'NUMEROCOMPTE' then
    result := TraduireMemoire('Numéro Compte');
  if UpperCase(st) = 'CLERIB' then
    result := TraduireMemoire('Clé');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 23/11/2006
Modifié le ... :   /  /    
Description .. : Procedure permettant l'affichage correct des differentes 
Suite ........ : zone, en fonction du pays et du type pays.
Mots clefs ... : 
*****************************************************************}
procedure Affiche ( FF : TForm; PaysIso, TypePays, prefixe : string ;TabOrder : integer=0) ;
var
  TypeCompte, stRetour : string;
  TRIB : TOB;
begin
  // On verifie s'il existe un parametrage spécifique pour ce pays.
  if existeParamPays(PaysIso, TypePays) then
  begin
    // Il existe un paramétrage.
    TypeCompte := quelType(PaysIso, TypePays) ;
    if TypeCompte = 'RIB' then
    begin
      // On est sur un RIB
      
      // On positionne les zones par defaut
      AfficheDefault(FF,prefixe,PaysIso,true,TabOrder) ;
      
      // On positionne les champs
      TRIB := Tob.Create('YIDENTBANCAIRE',nil,-1) ;
      stRetour := ParametreRIB( FF, PaysIso, TypePays , prefixe,  TRIB, true,TabOrder) ;
      THLabel(FF.FindComponent('T' + prefixe + 'ETABBQ')).Caption := stRetour ;

      // On desactive l'IBAN
      THLabel(FF.FindComponent('T' + prefixe + 'CODEIBAN')).Caption := TraduireMemoire('Code IBAN') ;
{$IFNDEF EAGLCLIENT}
      THDBEdit(FF.FindComponent(prefixe + 'CODEIBAN')).Enabled := False ;
      THDBEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).Hint := TraduireMemoire('Code IBAN') ;
{$ELSE}
      THEdit(FF.FindComponent(prefixe + 'CODEIBAN')).Enabled := False ;
      THEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).Hint := TraduireMemoire('Code IBAN') ;
{$ENDIF}
    end
    else if TypeCompte = 'IBA' then
    begin
      // On est sur un IBAN
      // On rend inactif les zones RIB apres les avoir positionné par defaut
      AfficheDefault(FF,prefixe,PaysIso,false,TabOrder) ;

      // On active l'IBAN
      THLabel(FF.FindComponent('T' + prefixe + 'CODEIBAN')).Caption := TraduireMemoire('Code IBAN') ;
{$IFNDEF EAGLCLIENT}
      THDBEdit(FF.FindComponent(prefixe + 'CODEIBAN')).Enabled := True ;
      THDBEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).Hint := TraduireMemoire('Code IBAN') ;
      // FQ 19433 BV 11/01/2007
      THDBEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).MaxLength := quelLongChamps(PaysIso, TypePays, 'CODEIBAN') ;
      // END FQ 19433
{$ELSE}
      THEdit(FF.FindComponent(prefixe + 'CODEIBAN')).Enabled := True ;
      THEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).Hint := TraduireMemoire('Code IBAN') ;
      // FQ 19433 BV 11/01/2007
      THEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).MaxLength := quelLongChamps(PaysIso, TypePays, 'CODEIBAN') ;
      // END FQ 19433
{$ENDIF}
    end
    else if TypeCompte = 'BBA' then
    begin
      // On est sur un BBAN
      // On rend inactif les zones RIB apres les avoir positionné par defaut
      AfficheDefault(FF,prefixe,PaysIso,false,TabOrder) ;

      // On active l'IBAN
      TRIB := Tob.Create('YIDENTBANCAIRE',nil,-1) ;
      ParametreRIB( FF, PaysIso, TypePays , prefixe,  TRIB, true,TabOrder) ;
    end
    else
    begin
      // Type de compte inconnu affichage RIB par defaut
      AfficheDefault(FF,prefixe,PaysIso,true,TabOrder);
      THLabel(FF.FindComponent('T' + prefixe + 'CODEIBAN')).Caption := TraduireMemoire('Code IBAN') ;
{$IFNDEF EAGLCLIENT}
      THDBEdit(FF.FindComponent(prefixe + 'CODEIBAN')).Enabled := False ;
      THDBEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).Hint := TraduireMemoire('Code IBAN') ;
{$ELSE}
      THEdit(FF.FindComponent(prefixe + 'CODEIBAN')).Enabled := False ;
      THEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).Hint := TraduireMemoire('Code IBAN') ;
{$ENDIF}
    end;
  end
  else
  begin
    // Il n'y a pas de paramétrage.
    // Selon le pays on ne va pas afficher les mêmes choses.
    {$IFDEF NOVH}
    if (PaysIso = 'FR') and (GetParamSocSecur('SO_PAYSLOCALISATION','') = 'FR') then  // FQ 19469 BV 11/01/2007
    {$ELSE}
    if (PaysIso = 'FR') and (VH^.PaysLocalisation = 'FR') then  // FQ 19469 BV 11/01/2007
    {$ENDIF}
    begin
      // On est en France sans paramétrage on affiche le RIB traditionnel et on desactive l'IBAN.
      AfficheDefault(FF,prefixe,PaysIso,true,TabOrder);
      THLabel(FF.FindComponent('T' + prefixe + 'CODEIBAN')).Caption := TraduireMemoire('Code IBAN') ;
{$IFNDEF EAGLCLIENT}
      THDBEdit(FF.FindComponent(prefixe + 'CODEIBAN')).Enabled := False ;
      THDBEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).Hint := TraduireMemoire('Code IBAN') ;
{$ELSE}
      THEdit(FF.FindComponent(prefixe + 'CODEIBAN')).Enabled := False ;
      THEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).Hint := TraduireMemoire('Code IBAN') ;
{$ENDIF}
    end
    else
    begin
      // Affichage standard
      AfficheDefault(FF,prefixe,PaysIso,false,TabOrder);
      THLabel(FF.FindComponent('T' + prefixe + 'CODEIBAN')).Caption := TraduireMemoire('Code IBAN') ;
{$IFNDEF EAGLCLIENT}
      THDBEdit(FF.FindComponent(prefixe + 'CODEIBAN')).Enabled := True ;
      THDBEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).Hint := TraduireMemoire('Code IBAN') ;
{$ELSE}
      THEdit(FF.FindComponent(prefixe + 'CODEIBAN')).Enabled := True ;
      THEDIT(FF.FindComponent(prefixe + 'CODEIBAN')).Hint := TraduireMemoire('Code IBAN') ;
{$ENDIF}
     end;
  end; 
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/11/2006
Modifié le ... :   /  /
Description .. : Fonction permettant d'afficher ou non les parametres
Suite ........ : de type pays : les deux libellés et la combobox
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure AfficheTYPEPAYS(FF : TForm; DS : TDataSet; prefixe, PaysIso:string) ;
var TypePays : string;
begin
  if existeParamPays(PaysIso)then
  begin
    THLabel(FF.FindComponent('T' + prefixe + 'TYPEPAYS')).Visible := True ;
{$IFNDEF EAGLCLIENT}
    //SDA le 14/03/2007 THDBEdit(FF.FindComponent(prefixe + 'TYPEPAYS')).Visible := True ;
    THDBValComboBox(FF.FindComponent(prefixe + 'TYPEPAYS')).Visible := True ;
{$ELSE}
    //SDA le 14/03/2007 THEdit(FF.FindComponent(prefixe + 'TYPEPAYS')).Visible := True ;
    THValComboBox(FF.FindComponent(prefixe + 'TYPEPAYS')).Visible := True ;
{$ENDIF}
    //SDA le 14/03/2007 THLabel(FF.FindComponent('L' + prefixe + 'TYPEPAYS')).Visible := True ;
    // FQ 19434 BV 11/01/2007
    // S'il n'y a qu'un seul parametrage on le selectionne par defaut.
    if DS.State in [dsInsert,dsEdit] then
    begin
        if (CountParamPays(PaysIso) = 1 ) then
            begin
                 TypePays := GetParamPays(PaysIso);
{$IFNDEF EAGLCLIENT}
                 //SDA le 14/03/2007 THDBEdit(FF.FindComponent(prefixe + 'TYPEPAYS')).Field.Value := TypePays ;
                  THDBValComboBox(FF.FindComponent(prefixe + 'TYPEPAYS')).ItemIndex := 0;
{$ELSE}
                 //SDA le 14/03/2007 THEdit(FF.FindComponent(prefixe + 'TYPEPAYS')).Text := TypePays ;
                 THValComboBox(FF.FindComponent(prefixe + 'TYPEPAYS')).ItemIndex := 0;
{$ENDIF}
                 //THLabel(FF.FindComponent('L' + prefixe + 'TYPEPAYS')).Caption := GetLibelleParamPays(PaysIso,TypePays);
            end;
    end;
    // END FQ 19434
  end
  else
  begin
    THLabel(FF.FindComponent('T' + prefixe + 'TYPEPAYS')).Visible := False ;
{$IFNDEF EAGLCLIENT}
    //SDA le 14/03/2007 THDBEdit(FF.FindComponent(prefixe + 'TYPEPAYS')).Visible := False ;
    THDBValComboBox(FF.FindComponent(prefixe + 'TYPEPAYS')).Visible := False ;
{$ELSE}
    //SDA le 14/03/2007 THEdit(FF.FindComponent(prefixe + 'TYPEPAYS')).Visible := False ;
    THValComboBox(FF.FindComponent(prefixe + 'TYPEPAYS')).Visible := False ;
{$ENDIF}
    //SDA le 14/03/2007 THLabel(FF.FindComponent('L' + prefixe + 'TYPEPAYS')).Visible := False ;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 21/11/2006
Modifié le ... :   /  /
Description .. : Fonction permettant de replacer les champs Banque,
Suite ........ : Guichet,Cle, Numéro de compte à leur taille et leurs
Suite ........ : positions initiales dans la fiche
Mots clefs ... :
*****************************************************************}
procedure AfficheDefault (FF : TForm; prefixe, PaysISO : string; active : boolean; TabOrder : integer=0) ;
var affiche : boolean;
begin
  // Cas specifique Treso pour les comptes généraux :
  if prefixe = 'BQ_' then
     affiche := active and not EstComptaTreso
  else
     affiche := active;
{$IFNDEF EAGLCLIENT}
  THDBEdit(FF.FindComponent(prefixe + 'ETABBQ')).Left := THDB_BanqueSav.Left;
  THDBEdit(FF.FindComponent(prefixe + 'ETABBQ')).Width := THDB_BanqueSav.Width;
  THDBEdit(FF.FindComponent(prefixe + 'ETABBQ')).Enabled := affiche;
  THDBEdit(FF.FindComponent(prefixe + 'ETABBQ')).TabOrder := 1 + TabOrder;
  THDBEdit(FF.FindComponent(prefixe + 'ETABBQ')).Visible := true;
  {$IFDEF NOVH}
  if (GetParamSocSecur('SO_PAYSLOCALISATION','') = CodeISOES) and (PaysISO = CodeISOES) then
  {$ELSE}
  if (VH^.PaysLocalisation = CodeISOES) and (PaysISO = CodeISOES) then
  {$ENDIF}
    THDBEdit(FF.FindComponent(prefixe + 'ETABBQ')).MaxLength := 4
  else
    THDBEdit(FF.FindComponent(prefixe + 'ETABBQ')).MaxLength := 5;
{$ELSE} 
  THEdit(FF.FindComponent(prefixe + 'ETABBQ')).Left := THDB_BanqueSav.Left;
  THEdit(FF.FindComponent(prefixe + 'ETABBQ')).Width := THDB_BanqueSav.Width;
  THEdit(FF.FindComponent(prefixe + 'ETABBQ')).Enabled := affiche;
  THEdit(FF.FindComponent(prefixe + 'ETABBQ')).TabOrder := 1 + TabOrder;
  THEdit(FF.FindComponent(prefixe + 'ETABBQ')).Visible := true;
  {$IFDEF NOVH}
  if (GetParamSocSecur('SO_PAYSLOCALISATION','') = CodeISOES) and (PaysISO = CodeISOES) then
  {$ELSE}
  if (VH^.PaysLocalisation = CodeISOES) and (PaysISO = CodeISOES) then
  {$ENDIF}
    THEdit(FF.FindComponent(prefixe + 'ETABBQ')).MaxLength := 4
  else
    THEdit(FF.FindComponent(prefixe + 'ETABBQ')).MaxLength := 5;
{$ENDIF}

{$IFNDEF EAGLCLIENT}
  THDBEdit(FF.FindComponent(prefixe + 'GUICHET')).Left := THDB_GuichetSav.Left;
  THDBEdit(FF.FindComponent(prefixe + 'GUICHET')).Width := THDB_GuichetSav.Width;
  THDBEdit(FF.FindComponent(prefixe + 'GUICHET')).Enabled := affiche;
  THDBEdit(FF.FindComponent(prefixe + 'GUICHET')).TabOrder := 2 + TabOrder;
  THDBEdit(FF.FindComponent(prefixe + 'GUICHET')).Visible := true;
  {$IFDEF NOVH}
  if (GetParamSocSecur('SO_PAYSLOCALISATION','') = CodeISOES) and (PaysISO = CodeISOES) then
  {$ELSE}
  if (VH^.PaysLocalisation = CodeISOES) and (PaysISO = CodeISOES) then
  {$ENDIF}
    THDBEdit(FF.FindComponent(prefixe + 'GUICHET')).MaxLength := 4
  else
    THDBEdit(FF.FindComponent(prefixe + 'GUICHET')).MaxLength := 5;
{$ELSE}
  THEdit(FF.FindComponent(prefixe + 'GUICHET')).Left := THDB_GuichetSav.Left;
  THEdit(FF.FindComponent(prefixe + 'GUICHET')).Width := THDB_GuichetSav.Width;
  THEdit(FF.FindComponent(prefixe + 'GUICHET')).Enabled := affiche;
  THEdit(FF.FindComponent(prefixe + 'GUICHET')).TabOrder := 2 + TabOrder;
  THEdit(FF.FindComponent(prefixe + 'GUICHET')).Visible := true;
  {$IFDEF NOVH}
  if (GetParamSocSecur('SO_PAYSLOCALISATION','') = CodeISOES) and (PaysISO = CodeISOES) then
  {$ELSE}
  if (VH^.PaysLocalisation = CodeISOES) and (PaysISO = CodeISOES) then
  {$ENDIF}
    THEdit(FF.FindComponent(prefixe + 'GUICHET')).MaxLength := 4
  else
    THEdit(FF.FindComponent(prefixe + 'GUICHET')).MaxLength := 5;
{$ENDIF}

{$IFNDEF EAGLCLIENT}
  THDBEdit(FF.FindComponent(prefixe + 'NUMEROCOMPTE')).Left := THDB_CompteSav.Left;
  THDBEdit(FF.FindComponent(prefixe + 'NUMEROCOMPTE')).Width := THDB_CompteSav.Width;
  THDBEdit(FF.FindComponent(prefixe + 'NUMEROCOMPTE')).Enabled := active;
  THDBEdit(FF.FindComponent(prefixe + 'NUMEROCOMPTE')).TabOrder := 3 + TabOrder;
  THDBEdit(FF.FindComponent(prefixe + 'NUMEROCOMPTE')).Visible := true;
  {$IFDEF NOVH}
  if (GetParamSocSecur('SO_PAYSLOCALISATION','') = CodeISOES) and (PaysISO = CodeISOES) then
  {$ELSE}
  if (VH^.PaysLocalisation = CodeISOES) and (PaysISO = CodeISOES) then
  {$ENDIF}
    THDBEdit(FF.FindComponent(prefixe + 'NUMEROCOMPTE')).MaxLength := 10
  else
    THDBEdit(FF.FindComponent(prefixe + 'NUMEROCOMPTE')).MaxLength := 11;
{$ELSE}
  THEdit(FF.FindComponent(prefixe + 'NUMEROCOMPTE')).Left := THDB_CompteSav.Left;
  THEdit(FF.FindComponent(prefixe + 'NUMEROCOMPTE')).Width := THDB_CompteSav.Width;
  THEdit(FF.FindComponent(prefixe + 'NUMEROCOMPTE')).Enabled := active;
  THEdit(FF.FindComponent(prefixe + 'NUMEROCOMPTE')).TabOrder := 3 + TabOrder;
  THEdit(FF.FindComponent(prefixe + 'NUMEROCOMPTE')).Visible := true;
  {$IFDEF NOVH}
  if (GetParamSocSecur('SO_PAYSLOCALISATION','') = CodeISOES) and (PaysISO = CodeISOES) then
  {$ELSE}
  if (VH^.PaysLocalisation = CodeISOES) and (PaysISO = CodeISOES) then
  {$ENDIF}
    THEdit(FF.FindComponent(prefixe + 'NUMEROCOMPTE')).MaxLength := 10
  else
    THEdit(FF.FindComponent(prefixe + 'NUMEROCOMPTE')).MaxLength := 11;
{$ENDIF}

{$IFNDEF EAGLCLIENT}
  THDBEdit(FF.FindComponent(prefixe + 'CLERIB')).Left := THDB_CleSav.Left;
  THDBEdit(FF.FindComponent(prefixe + 'CLERIB')).Width := THDB_CleSav.Width;
  THDBEdit(FF.FindComponent(prefixe + 'CLERIB')).Enabled := active;
  THDBEdit(FF.FindComponent(prefixe + 'CLERIB')).TabOrder := 4 + TabOrder;
  THDBEdit(FF.FindComponent(prefixe + 'CLERIB')).Visible := true;
  THDBEdit(FF.FindComponent(prefixe + 'CLERIB')).MaxLength := 2;
{$ELSE}
  THEdit(FF.FindComponent(prefixe + 'CLERIB')).Left := THDB_CleSav.Left;
  THEdit(FF.FindComponent(prefixe + 'CLERIB')).Width := THDB_CleSav.Width;
  THEdit(FF.FindComponent(prefixe + 'CLERIB')).Enabled := active;
  THEdit(FF.FindComponent(prefixe + 'CLERIB')).TabOrder := 4 + TabOrder;
  THEdit(FF.FindComponent(prefixe + 'CLERIB')).Visible := true;
  THEdit(FF.FindComponent(prefixe + 'CLERIB')).MaxLength := 2;
{$ENDIF}

{$IFNDEF EAGLCLIENT}
  THDBEdit(FF.FindComponent(prefixe + 'CODEIBAN')).Width := THDB_CodeIbanSav.Width;
  THDBEdit(FF.FindComponent(prefixe + 'CODEIBAN')).MaxLength := THDB_CodeIbanSav.MaxLength;
  THDBEdit(FF.FindComponent(prefixe + 'CODEIBAN')).Enabled := THDB_CodeIbanSav.Enabled;
{$ELSE}
  THEdit(FF.FindComponent(prefixe + 'CODEIBAN')).Width := THDB_CodeIbanSav.Width;
  THEdit(FF.FindComponent(prefixe + 'CODEIBAN')).MaxLength := THDB_CodeIbanSav.MaxLength;
  THEdit(FF.FindComponent(prefixe + 'CODEIBAN')).Enabled := THDB_CodeIbanSav.Enabled;
{$ENDIF}

  THLabel(FF.FindComponent('T' + prefixe + 'ETABBQ')).Visible := true;
  THLabel(FF.FindComponent('T' + prefixe + 'ETABBQ')).Caption := THLabel_BanqueSav.Caption; 
  THLabel(FF.FindComponent('T' + prefixe + 'ETABBQ')).Left := THLabel_BanqueSav.Left;
  THLabel(FF.FindComponent('T' + prefixe + 'GUICHET')).Visible := true;
  THLabel(FF.FindComponent('T' + prefixe + 'GUICHET')).Caption := THLabel_GuichetSav.Caption; 
  THLabel(FF.FindComponent('T' + prefixe + 'GUICHET')).Left := THLabel_GuichetSav.Left;
  THLabel(FF.FindComponent('T' + prefixe + 'NUMEROCOMPTE')).Visible := true;
  THLabel(FF.FindComponent('T' + prefixe + 'NUMEROCOMPTE')).Caption := THLabel_CompteSav.Caption  ;
  THLabel(FF.FindComponent('T' + prefixe + 'NUMEROCOMPTE')).Left := THLabel_CompteSav.Left ;
  THLabel(FF.FindComponent('T' + prefixe + 'CLERIB')).Visible := true;
  THLabel(FF.FindComponent('T' + prefixe + 'CLERIB')).Caption := THLabel_CleSav.Caption ;
  THLabel(FF.FindComponent('T' + prefixe + 'CLERIB')).Left := THLabel_CleSav.Left;

  THLabel(FF.FindComponent('T' + prefixe + 'ETABBQ')).Caption := '&' + TraduireMemoire('Banque');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 21/11/2006
Modifié le ... :   /  /
Description .. : Fonction permettant de sauvegarder les postions des
Suite ........ : composants RIB de la fiche
Mots clefs ... :
*****************************************************************}
procedure SaveAffichageDefaut (FF : TForm; prefixe : string) ;
begin
  // Zones
{$IFNDEF EAGLCLIENT}
  THDB_BanqueSav := THDBEdit.Create(nil);
  THDB_BanqueSav.Left := THDBEdit(FF.FindComponent(prefixe + 'ETABBQ')).Left;
  THDB_BanqueSav.Width := THDBEdit(FF.FindComponent(prefixe + 'ETABBQ')).Width;

  THDB_GuichetSav := THDBEdit.Create(nil);
  THDB_GuichetSav.Left := THDBEdit(FF.FindComponent(prefixe + 'GUICHET')).Left;
  THDB_GuichetSav.Width := THDBEdit(FF.FindComponent(prefixe + 'GUICHET')).Width;

  THDB_CompteSav := THDBEdit.Create(nil);
  THDB_CompteSav.Left := THDBEdit(FF.FindComponent(prefixe + 'NUMEROCOMPTE')).Left;
  THDB_CompteSav.Width := THDBEdit(FF.FindComponent(prefixe + 'NUMEROCOMPTE')).Width;

  THDB_CleSav := THDBEdit.Create(nil);
  THDB_CleSav.Left := THDBEdit(FF.FindComponent(prefixe + 'CLERIB')).Left;
  THDB_CleSav.Width := THDBEdit(FF.FindComponent(prefixe + 'CLERIB')).Width;

  THDB_CodeIbanSav := THDBEdit.Create(nil);
  THDB_CodeIbanSav.MaxLength := THDBEdit(FF.FindComponent(prefixe + 'CODEIBAN')).MaxLength;
  THDB_CodeIbanSav.Width := THDBEdit(FF.FindComponent(prefixe + 'CODEIBAN')).Width;
  THDB_CodeIbanSav.Enabled := THDBEdit(FF.FindComponent(prefixe + 'CODEIBAN')).Enabled;
{$ELSE}
  THDB_BanqueSav := THEdit.Create(nil);
  THDB_BanqueSav.Left := THEdit(FF.FindComponent(prefixe + 'ETABBQ')).Left;
  THDB_BanqueSav.Width := THEdit(FF.FindComponent(prefixe + 'ETABBQ')).Width;

  THDB_GuichetSav := THEdit.Create(nil);
  THDB_GuichetSav.Left := THEdit(FF.FindComponent(prefixe + 'GUICHET')).Left;
  THDB_GuichetSav.Width := THEdit(FF.FindComponent(prefixe + 'GUICHET')).Width;

  THDB_CompteSav := THEdit.Create(nil);
  THDB_CompteSav.Left := THEdit(FF.FindComponent(prefixe + 'NUMEROCOMPTE')).Left;
  THDB_CompteSav.Width := THEdit(FF.FindComponent(prefixe + 'NUMEROCOMPTE')).Width;

  THDB_CleSav := THEdit.Create(nil);
  THDB_CleSav.Left := THEdit(FF.FindComponent(prefixe + 'CLERIB')).Left;
  THDB_CleSav.Width := THEdit(FF.FindComponent(prefixe + 'CLERIB')).Width;

  THDB_CodeIbanSav := THEdit.Create(nil);
  THDB_CodeIbanSav.MaxLength := THEdit(FF.FindComponent(prefixe + 'CODEIBAN')).MaxLength;
  THDB_CodeIbanSav.Width := THEdit(FF.FindComponent(prefixe + 'CODEIBAN')).Width;
  THDB_CodeIbanSav.Enabled := THEdit(FF.FindComponent(prefixe + 'CODEIBAN')).Enabled;
{$ENDIF}

  // Labels
  THLabel_BanqueSav := THLabel.Create(nil);
  THLabel_BanqueSav.Left := THLabel(FF.FindComponent('T' + prefixe + 'ETABBQ')).Left;
  THLabel_BanqueSav.Caption := THLabel(FF.FindComponent('T' + prefixe + 'ETABBQ')).Caption;
  THLabel_GuichetSav := THLabel.Create(nil);
  THLabel_GuichetSav.Left := THLabel(FF.FindComponent('T' + prefixe + 'GUICHET')).Left;
  THLabel_GuichetSav.Caption := THLabel(FF.FindComponent('T' + prefixe + 'GUICHET')).Caption;
  THLabel_CompteSav := THLabel.Create(nil);
  THLabel_CompteSav.Left := THLabel(FF.FindComponent('T' + prefixe + 'NUMEROCOMPTE')).Left;
  THLabel_CompteSav.Caption := THLabel(FF.FindComponent('T' + prefixe + 'NUMEROCOMPTE')).Caption;
  THLabel_CleSav := THLabel.Create(nil);
  THLabel_CleSav.Left := THLabel(FF.FindComponent('T' + prefixe + 'CLERIB')).Left;
  THLabel_CleSav.Caption := THLabel(FF.FindComponent('T' + prefixe + 'CLERIB')).Caption;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 21/11/2006
Modifié le ... :   /  /
Description .. : Fonction permettant de supprimer les objet THDBEdit
Suite ........ : créer pour sauvegarder l'affichage
Mots clefs ... :
*****************************************************************}
procedure DestroySaveAffichage ;
begin
	if assigned(THDB_BanqueSav) then
  	THDB_BanqueSav.Free ;
  if assigned(THDB_GuichetSav) then
  	THDB_GuichetSav.Free ;
  if assigned(THDB_CompteSav) then
	  THDB_CompteSav.Free ;
  if assigned(THDB_CleSav) then
	  THDB_CleSav.Free ;
  if assigned(THDB_CodeIbanSav) then
    THDB_CodeIbanSav.Free;
  if assigned(THLabel_BanqueSav) then
	  THLabel_BanqueSav.Free;
  if assigned(THLabel_GuichetSav) then
	  THLabel_GuichetSav.Free;
  if assigned(THLabel_CompteSav) then
	  THLabel_CompteSav.Free;
  if assigned(THLabel_CleSav) then
  	THLabel_CleSav.Free;
end;


end.




