unit DecisionAchatUtil;

interface
uses Classes,UTOB,FactComm,Hctrls,HEnt1,SaisUtil,UtilPGi,Forms,Vierge,
{$IFDEF EAGLCLIENT}
  maineagl
{$ELSE}
  Doc_Parser,DBCtrls, Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} fe_main
{$ENDIF}
	,HmsgBox
;

procedure AddlesChampsSupDet (TOBD : TOB);
procedure AddlesChampsSup(TOBdecisionnel : TOB;var MaxRemplace : integer);
procedure AfficheDetailDecision (GS : THGrid;TOBSimulpres : TOB;NiveauMax : integer;ListeSaisie:String;var LigneDepart : integer);
procedure AfficheLaGrille (XX : TForm;ListeSaisie: string;TOBDecision : TOB;NiveauMax,NbElt : integer;RefNumLigne:integer=0);
procedure ConstitueLaTOBGestion (TOBdecisionnel,TOBInterm : TOB);
procedure FermeBranche (GS:THGrid; TobSimulPres: TOB;var Arow:integer;var Niveau : integer; var Found : boolean);
function  FindFirstParent (TOBL : TOB) : TOB;
function  FindParentInit (TOBL : TOB) : TOB;
function  GetTOBDecisAch(TOBDecis: TOB; ARow : integer) : TOB;
procedure InitBranche (GS : THGrid; TOBSimulPres : TOB; Arow : integer);
procedure IndiceGrilleInit (TOBDecisionnel : TOB; NiveauMax : integer; var Numligne : integer);
procedure OuvreBranche (GS : THGrid; TobSimulPres: TOB;var Arow:integer;var niveau : integer; var Found : boolean; Force : boolean);
procedure OuvreTouteLaBranche (GS : THGrid; TOBL,TobSimulPres: TOB;var Arow:integer;var niveau : integer; var Found : boolean; Force : boolean);
function RecupTypeGraph (TOBL : TOB) : integer;
function RecupTypeGraphOpenClose (TOBL : TOB) : integer;
procedure ReIndiceGrille (TOBDecisionnel : TOB; var Numligne : integer);
procedure ReinitStkPhysique (TOBBranche : TOB);
procedure SetStkPhysique (TOBInterm : TOB);
function RecupTypeConsFou (TOBL : TOB) : integer;
function TrouveLigneinDecisionnel (TOBDecisionnel : TOB; Niveau : integer; Article,LivraisonChantier : string; Depot : string) : TOB;

implementation

procedure AddlesChampsSupDet (TOBD : TOB);
begin
  TOBD.AddChampSupValeur ('OPEN','-');
  TOBD.AddChampSupValeur ('OPEN1','-');
  TOBD.AddChampSupValeur ('OPEN2','-');
  if TOBD.GetValue('BAD_TYPEL')='NI1' then TOBD.AddChampSupValeur ('NIVEAU',1) else
  if TOBD.GetValue('BAD_TYPEL')='NI2' then TOBD.AddChampSupValeur ('NIVEAU',2) else
  if TOBD.GetValue('BAD_TYPEL')='NI3' then TOBD.AddChampSupValeur ('NIVEAU',3) else
  if TOBD.GetValue('BAD_TYPEL')='ART' then
  begin
    TOBD.AddChampSupValeur ('NIVEAU',4);
    TOBD.AddChampSupValeur ('OPEN',' ');
  end;
  TOBD.AddChampSupValeur ('NUMAFF',0);
  TOBD.AddChampSupValeur ('QTEPHY',0.0);
  TOBD.AddChampSupValeur ('__MODIFIE','-');
  TOBD.AddChampSupValeur ('ARTICLEREMPL','');
	TOBD.AddChampSupValeur ('ARTICLEREMPLPAR','');
  TOBD.AddChampSupValeur ('ARTICLEREMPLACE','');
  if not TOBD.FieldExists ('T_LIBELLE') then TOBD.AddChampSupValeur  ('T_LIBELLE','');
  if not TOBD.FieldExists ('GA_TENUESTOCK') then TOBD.AddChampSupValeur  ('GA_TENUESTOCK','-');
  if not TOBD.FieldExists ('GA_LIBELLE') then TOBD.AddChampSupValeur  ('GA_LIBELLE','');
  if not TOBD.FieldExists ('AFF_LIBELLE') then TOBD.AddChampSupValeur  ('AFF_LIBELLE','');
  TOBD.AddChampSupValeur ('INDICECONSULT',0);
  TOBD.AddChampSupValeur ('POSITIONCONSULT','AUC');
  TOBD.AddChampSupValeur ('SELECT','-');
  TOBD.AddChampSupValeur ('_A_DELETE','-');
  TOBD.AddChampSupValeur ('_A_TRAITER','X');
  TOBD.AddChampSupValeur ('MULTIPABASE','-');
end;

procedure AddlesChampsSup(TOBdecisionnel : TOB; var MaxRemplace : integer);
var Indice : integer;
		TOBD : TOB;
begin
	for Indice := 0 to TOBdecisionnel.detail.count -1 do
  begin
    TOBD := TOBdecisionnel.detail[Indice];
    if MAxRemplace < TOBD.GetValue('BAD_NUMREMPLACE') then  MAxRemplace := TOBD.GetValue('BAD_NUMREMPLACE');
    AddlesChampsSupDet (TOBD);
  end;
end;

procedure AfficheDetailDecision (GS : THGrid;TOBSimulpres : TOB;NiveauMax : integer;ListeSaisie:String;var LigneDepart : integer);
var Indice : integer;
    TOBS : TOB;
    Ligne : integer;
begin
  if LigneDepart = 0 then LigneDepart := 1;
	TOBS := GetTOBDecisAch (TOBSimulpres, LigneDepart);
  repeat
  	if TOBS <> nil then
    begin
			Ligne := TOBS.GetValue('NUMAFF');
      TOBS.PutLigneGrid (GS,Ligne,false,false,ListeSaisie);
      GS.InvalidateRow (Ligne);
      if (TOBS.detail.count > 0) and (NiveauMax > TOBS.GetValue('NIVEAU')) then
      begin
        AfficheDetailDecision (GS,TOBS,NiveauMax,ListeSaisie,Ligne);
      end;
    	inc(Ligne);
			TOBS := GetTOBDecisAch (TOBSimulpres, Ligne);
    end;
  until TOBS = nil;
end;

procedure AfficheLaGrille (XX : TForm;ListeSaisie: string;TOBDecision : TOB;NiveauMax,NbElt : integer; RefNumLigne:integer=0);
var LaLigneDep : integer;
    GS : THgrid;
begin
  GS := THgrid(TForm(XX).FindComponent ('GS'));
  gS.BeginUpdate;
	Gs.SynEnabled := false;
  if GS.rowCount <> NbElt + 1 then GS.rowCount := NbElt+1;
  LaLigneDep := RefNumLigne;
  AfficheDetailDecision (GS,TOBDecision,NiveauMax,ListeSaisie,LaLigneDep);
  gs.SynEnabled := true;
  gs.EndUpdate;
  TFVierge(XX).HMTrad.ResizeGridColumns (GS);
end;

procedure ConstitueLaTOBGestion (TOBdecisionnel,TOBInterm : TOB);
var Indice : Integer;
		TOBI,TOBN1,TOBN2,TOBN3 : TOB;
begin
	Indice := 0;
	repeat
  	TOBI := TOBInterm.detail[Indice];
    if TOBI.GetValue('BAD_TYPEL') = 'NI1' then
    begin
      TOBI.ChangeParent (TOBdecisionnel,-1);
    end else if TOBI.GetValue('BAD_TYPEL') = 'NI2' then
    begin
    	TOBN1 := TOBdecisionnel.findFirst (['BAD_NUMN1'],[TOBI.GetValue('BAD_NUMN1')],true);
      if TOBN1 <> nil then TOBI.ChangeParent (TOBN1,-1) else inc(Indice);
    end else if TOBI.GetValue('BAD_TYPEL') = 'NI3' then
    begin
    	TOBN2 := TOBdecisionnel.findFirst (['BAD_NUMN1','BAD_NUMN2'],[TOBI.GetValue('BAD_NUMN1'),TOBI.GetValue('BAD_NUMN2')],true);
      if TOBN2 <> nil then TOBI.ChangeParent (TOBN2,-1) else inc(Indice);
    end else if TOBI.GetValue('BAD_TYPEL') = 'ART' then
    begin
    	TOBN3 := TOBdecisionnel.findFirst (['BAD_NUMN1','BAD_NUMN2','BAD_NUMN3'],
      																	[TOBI.GetValue('BAD_NUMN1'),TOBI.GetValue('BAD_NUMN2'),TOBI.GetValue('BAD_NUMN3')],true);
      if TOBN3 <> nil then TOBI.ChangeParent (TOBN3,-1) else inc(indice);
    end;
  until (TobInterm.Detail.count = 0) or (Indice >= TobInterm.Detail.count) ;
end;

function  FindFirstParent (TOBL : TOB) : TOB;
var Loc : TOB;
begin
	result := nil;
  if not TOBL.fieldExists('NIVEAU') then exit;
  if TOBL.GetValue ('NIVEAU') = 1 then
  begin
  	result := TOBL;
  end else
  begin
  	Loc := FindFirstParent (TOBL.parent);
    if Loc = nil then result := TOBL else result := loc
  end;
end;

function  FindParentInit (TOBL : TOB) : TOB;
var Loc : TOB;
begin
	result := nil;
  // ca c'est pour verifier que l'on est bien sur un niveau a prendre en compte
  if not TOBL.fieldExists('NIVEAU') then exit;
  // --
	if TOBL.Parent <> nil then
  begin
  	LOc := findparentInit(TOBL.Parent);
    if Loc <> nil then result := Loc else result := TOBL.parent;
  end;
end;

function GetTOBDecisAch(TOBDecis: TOB; ARow : integer) : TOB;
begin
	Result := TOBDecis.findFirst(['NUMAFF'],[Arow],true);
end;

procedure IndiceGrilleInit (TOBDecisionnel : TOB; NiveauMax : integer; var Numligne : integer);
var Indice : integer;
    TOBS : TOB;
begin
	for Indice := 0 to TOBDecisionnel.detail.count -1 do
  begin
    TOBS := TOBDecisionnel.detail[Indice];
    if TOBS.GetValue('NIVEAU')> 0 then
    BEGIN
      TOBS.putValue('NUMAFF',NumLigne);
      inc(NumLigne);
    END;
(*
		if NiveauMax > TOBS.GetValue('NIVEAU') then
    begin
      TOBS.putValue('OPEN','X');
      IndicegrilleInit (TOBS,NiveauMax,Numligne);
    end;
*)
  end;
end;


procedure ReIndiceGrille (TOBDecisionnel : TOB; var Numligne : integer);
var Indice : integer;
    TOBS : TOB;
begin
	for Indice := 0 to TOBDecisionnel.detail.count -1 do
  begin
    TOBS := TOBDecisionnel.detail[Indice];
    if TOBS.GetValue('NIVEAU')> 0 then
    BEGIN
      inc(NumLigne);
      TOBS.putValue('NUMAFF',NumLigne);
    END;
    if (TOBS.GetValue('OPEN')='X') and (TOBS.GetValue('NIVEAU')=1) then
    begin
      ReIndicegrille (TOBS,Numligne);
    end else if (TOBS.GetValue('OPEN1')='X') and (TOBS.GetValue('NIVEAU')=2) then
    begin
      ReIndicegrille (TOBS,Numligne);
    end else if (TOBS.GetValue('OPEN2')='X') and (TOBS.GetValue('NIVEAU')=3) then
    begin
      ReIndicegrille (TOBS,Numligne);
    end;
  end;
end;

procedure PositionneBrancheOpen (TOBL : TOB);
var TOBI : TOB;
		Niveau : integer;
begin
  if not TOBL.fieldExists('NIVEAU') then exit;
  NIveau := TOBL.GetValue ('NIVEAU');
  if Niveau = 1 then
  begin
  	TOBL.PutValue('OPEN','X');
  end else if Niveau = 2 then
  begin
    TOBL.PutValue('OPEN','X');
    TOBL.PutValue('OPEN1','X');
  end else if NIveau = 3 then
  begin
    TOBL.PutValue('OPEN','X');
    TOBL.PutValue('OPEN1','X');
    TOBL.PutValue('OPEN2','X');
  end;
  if TOBL.Parent <> nil then  PositionneBrancheOpen (TOBL.Parent);
end;

procedure OuvreTouteLaBranche (GS : THGrid; TOBL,TobSimulPres: TOB;var Arow:integer;var niveau : integer; var Found : boolean; Force : boolean);
var TOBI : TOB;
    Indice,NiveauLoc : integer;
    LigneCourante: integer;
begin
	TOBI := FindFirstParent (TOBL);
  LigneCourante := TOBI.GetValue('NUMAFF');
  PositionneBrancheOpen (TOBL);
  Arow := 0;
	ReIndiceGrille (TObSimulPres,Arow);
end;


procedure OuvreBranche (GS : THGrid; TobSimulPres: TOB;var Arow:integer;var niveau : integer; var Found : boolean; Force : boolean);
var TOBS : TOB;
    Indice : integer;
    LigneCourante: integer;
begin
  for Indice := 0 to TOBSimulPres.detail.count -1 do
  begin
    TOBS := TOBSimulPres.detail[Indice];
    LigneCourante := TOBS.GetValue('NUMAFF');
    if (LigneCourante <> Arow) and (lignecourante > 0) then if TOBS.GetValue('NIVEAU') > Niveau then NIveau := TOBS.GetValue('NIVEAU');
    if (found) and ((LigneCourante<> 0) or (force)) then
    BEGIN
      inc(Arow);
      TOBS.putValue('NUMAFF',Arow);
      if TOBS.GetValue('NIVEAU') > niveau then niveau := TOBS.GetValue('NIVEAU');
      if TOBS.detail.count > 0 then OuvreBranche (GS,TOBS,Arow,niveau,Found,false);
      continue;
    END;
    if (LigneCourante = 0 ) then break;
    if LigneCourante = Arow then
    BEGIN
      found := true;
      // on force la mAJ du numero de ligne dans le niveai immediatement superieur
      OuvreBranche (GS,TOBS,Arow,niveau,Found,true);
    END else
    if TOBS.detail.count > 0 then OuvreBranche (GS,TOBS,Arow,niveau,Found,Force);
  end;
end;

procedure FermeBranche (GS:THGrid; TobSimulPres: TOB;var Arow:integer;var Niveau : integer; var Found : boolean);
var TOBS: TOB;
    Indice : integer;
    LigneCourante: integer;
begin
  for Indice := 0 to TOBSimulPres.detail.count -1 do
  begin
    TOBS := TOBSimulPres.detail[Indice];
    LigneCourante := TOBS.GetValue('NUMAFF');
    if (LigneCourante <> Arow) and (lignecourante > 0) then if TOBS.GetValue('NIVEAU') > Niveau then NIveau := TOBS.GetValue('NIVEAU');
    if (found) and (LigneCourante<> 0)  then
    BEGIN
      if TOBS.GetValue('NIVEAU') > Niveau then NIveau := TOBS.GetValue('NIVEAU');
      inc(Arow);
      TOBS.putValue('NUMAFF',Arow);
      FermeBranche (GS,TOBS,Arow,niveau,Found);
      continue;
    END;
    if (LigneCourante = 0 ) then break;
    if LigneCourante = Arow then
    BEGIN
      found := true;
      InitBranche (GS,TOBS,Arow);
    END else if TOBS.detail.count > 0 then FermeBranche (GS,TOBS,Arow,niveau,Found);
  end;
end;

procedure InitBranche (GS : THGrid; TOBSimulPres : TOB; Arow : integer);
var TOBS : TOB;
    Indice : integer;
    Niveau : integer;
begin
  for Indice := 0 to TOBSimulPres.detail.count -1 do
  begin
    TOBS := TOBSimulPres.detail[Indice];
    Niveau := TOBS.GetValue('NIVEAU');
    if TOBS.GetValue('NUMAFF') = 0 then break;
    TOBS.PutValue('NUMAFF',0);
    if niveau = 1 then
    begin
    	TOBS.PuTValue('OPEN','-');
    	TOBS.PuTValue('OPEN1','-');
    	TOBS.PuTValue('OPEN2','-');
    end else if Niveau = 2 then
    begin
    	TOBS.putValue('OPEN1','-');
    	TOBS.putValue('OPEN2','-');
    end else if Niveau = 3 then
    begin
    	TOBS.putValue('OPEN2','-');
    end;
    if TOBS.detail.count > 0 then
    begin
    	IniTBranche (GS,TOBS,Arow);
    end;
  end;
end;


function RecupTypeGraph (TOBL : TOB) : integer;
BEGIN
	if TOBL = nil then exit;
	if TOBL.GetValue('BAD_TYPEARTICLE')='MAR' then result := 1 else
	if TOBL.GetValue('BAD_TYPEARTICLE')='ARP' then result := 2 else
	if TOBL.GetValue('BAD_TYPEARTICLE')='PRE' then result := 0;
END;

function RecupTypeGraphOpenClose (TOBL : TOB) : integer;
var Niveau : integer;
begin
	if TOBL = nil then exit;
	niveau :=TOBL.getValue('NIVEAU');
	if (niveau = 1) then
  begin
  	if TOBL.GetValue('OPEN')='-' then result := 0 else result := 1;
  end else if (Niveau = 2) then
  begin
  	if TOBL.GetValue('OPEN1')='-' then result := 0 else result := 1;
  end else if (niveau = 3) then
  begin
  	if TOBL.GetValue('OPEN2')='-' then result := 0 else result := 1;
  end;

end;

function RecupTypeConsFou (TOBL : TOB) : integer;
var Pos : string;
begin
	if TOBL = nil then exit;
	Pos :=TOBL.getValue('POSITIONCONSULT');
  if Pos='AUC' then result := 0 else
  if Pos='APR' then result := 1 else
  if Pos='ICI' then result := 2 else
  if Pos='AVA' then result := 3;
end;

procedure ReinitStkPhysique (TOBBranche : TOB);
begin
	if not TOBBranche.FieldExists ('NIVEAU') then exit;
	if TOBBranche.getValue('NIVEAU') < 4 then TOBBranche.PutValue('QTEPHY',0.0);
  if TOBBranche.parent <> nil then
  begin
  	ReinitStkPhysique (TOBBranche.parent);
  end;
end;

procedure SetStkPhysique (TOBInterm : TOB);
var Indice : integer;
		TOBI,TOBL : TOB;
    Requete : string;
    QQ: TQuery;
    OkFind : boolean;
begin
	for Indice := 0 to tobinterm.detail.count-1 do
  begin
    TOBI := TOBInterM.detail[Indice];
    if (TOBI.GetValue('GA_TENUESTOCK') = 'X') and (TOBI.GetValue('BAD_NUMN4')>0) then
    begin
    	requete := 'SELECT GQ_PHYSIQUE FROM DISPO WHERE GQ_ARTICLE="'+TOBI.getValue('BAD_ARTICLE')+'" AND '+
      					 'GQ_DEPOT="'+TOBI.GetValue('BAD_DEPOT')+'" AND GQ_CLOTURE="-"';
      QQ := OpenSql (Requete,true,-1,'',true);
      if not QQ.eof then TOBI.PutValue('QTEPHY',QQ.findField('GQ_PHYSIQUE').asfloat);
      ferme (QQ);
    end;
    if TOBI.GetValue('BAD_NUMREMPLACE') <> 0 then
    begin
    	TOBL := TOBInterm.FindFirst (['BAD_NUMREMPLACE'],[TOBI.GetValue('BAD_NUMREMPLACE')],true);
      OkFind := false;
      repeat
      	if (TOBL <> TOBI) and (TOBL <> nil) then begin OkFind := true; break end;
        if TOBL <> nil then TOBL := TOBInterm.Findnext (['BAD_NUMREMPLACE'],[TOBI.GetValue('BAD_NUMREMPLACE')],true);
      until (TOBL = nil) or (OKfind);

      if TOBL <> Nil then
      begin
      	if TOBI.GetValue('BAD_REMPLACE')='X' then
        begin
        	TOBI.putValue ('ARTICLEREMPLPAR',TOBL.GetValue('BAD_CODEARTICLE'));
      		TOBL.putValue ('ARTICLEREMPLACE',TOBI.GetValue('BAD_CODEARTICLE'));
        end else
        begin
      		TOBI.putValue ('ARTICLEREMPLACE',TOBL.GetValue('BAD_CODEARTICLE'));
      		TOBL.putValue ('ARTICLEREMPLPAR',TOBI.GetValue('BAD_CODEARTICLE'));
        end;
      end;
    end;

  end;
end;

function TrouveLigneinDecisionnel (TOBDecisionnel : TOB; Niveau : integer; Article,LivraisonChantier: string; Depot : string) : TOB;
begin
  case niveau of
  1 : result := TOBdecisionnel.findFirst(['BAD_TYPEL','BAD_ARTICLE'],['NI1',Article],true);
  2 : result := TOBdecisionnel.findFirst(['BAD_TYPEL','BAD_ARTICLE','BAD_LIVCHANTIER'],['NI2',Article,LivraisonChantier],True);
  3 : result:= TOBdecisionnel.findFirst(['BAD_TYPEL','BAD_ARTICLE','BAD_LIVCHANTIER','BAD_DEPOT'],['NI3',Article,LivraisonChantier,Depot],True);
  else result := nil;
  end;

end;


end.
