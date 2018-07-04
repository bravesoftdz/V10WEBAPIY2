unit GCMZSUtil;

interface

uses  HQry,HDB,HEnt1,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ENDIF}
{$IFDEF GRC}
UtilRT,
{$ENDIF}
      GCModifZS,Utob,HCtrls,SysUtils,TiersUtil;

Type
     TO_ModifParLot = class
       Q : THQuery;
{$IFDEF EAGLCLIENT}
       F : THGrid;
{$ELSE}
       F : THDBGrid;
{$ENDIF}
       NatureTiers : string;
       TableName : string;
       Titre : String;
       FCode : String;
       Nature : String;
       FicheAOuvrir : String;
       StSQL : String;
       // Modif Btp
       StContexte : String;
       stAbrege : String;
     public
       constructor create;
       // -------- 
     end;

procedure ModifieEnSerie(ObjModif : TO_ModifParLot; Parametres : String);
procedure ModifieRafale(Obj : TO_ModifParLot; Table :string) ;
procedure ModifieFicheParFiche(Obj : TO_ModifParLot) ;
procedure InsertOrUpdateTable(StSQL,Table,cle : string);
function RechListeChamps(StSQL,stPrefixe : string) : string;
function VerifPrefixe(Table, Code : string) : string;

implementation
uses CbpMCD
  ,CbpEnumerator;

constructor  TO_ModifParLot.create ;
begin
inherited create;
stcontexte := '';
end;

procedure FinalizeMajFromChamp(TOBdetail : TOB);

  function ValoBaseModified(TOBD : TOB) : boolean;
  begin
    result := false;
    if TOBD.IsFieldModified ('GA_PAHT') or TOBD.IsFieldModified ('GA_COEFFG') then result := true;
  end;

  function ValoRevientModified(TOBD : TOB) : boolean;
  begin
    result := false;
    if TOBD.IsFieldModified ('GA_DPR') or TOBD.IsFieldModified ('GA_COEFFG') or
       TOBD.IsFieldModified ('GA_DPRAUTO') or TOBD.IsFieldModified ('GA_CALCPRIXPR') then result := true;
  end;

  function ValoVenteHtModified(TOBD : TOB) : boolean;
  begin
    result := false;
    if TOBD.IsFieldModified ('GA_PVHT') or TOBD.IsFieldModified ('GA_COEFCALCHT') or
       TOBD.IsFieldModified ('GA_CALCAUTOHT') or TOBD.IsFieldModified ('GA_CALCPRIXHT') then result := true;
  end;

var Indice : integer;
    TOBD : TOB;
begin
  for Indice := 0 to TOBDetail.detail.count -1 do
  begin
    TOBD := TOBdetail.detail[Indice];
    if TOBD.NomTable <> 'ARTICLE' then break; // pour l'instant pas de maj sup pour les autres tables

    // Partie achat modifiée
    if ValobaseModified(TOBD) then
    begin
      // DPR depaendant du PA
      if (TOBD.GetValue('GA_CALCPRIXPR')='PAA') and (TOBD.GetValue('GA_DPRAUTO')='X') then
      begin
        TOBD.PutValue('GA_DPR',Arrondi(TOBD.getValue('GA_PAHT')*TOBD.getValue('GA_COEFFG'),V_PGI.OkdecP));
      end;
      // PVHT depaendant du PA
      if (TOBD.GetValue('GA_CALCPRIXHT')='PAA') and (TOBD.GetValue('GA_CALCAUTOHT')='X') then
      begin
        TOBD.PutValue('GA_PVHT',Arrondi(TOBD.getValue('GA_PAHT')*TOBD.getValue('GA_COEFCALCHT'),V_PGI.OkdecP));
      end;
    end;
    // partie revient modifie
    if ValoRevientModified(TOBD) then
    begin
      //
      if (TOBD.GetValue('GA_CALCPRIXPR')='PAA') and (TOBD.GetValue('GA_DPRAUTO')='X') then
      begin
        if TOBD.IsFieldModified ('GA_COEFFG') then
        begin
          if TOBD.getValue('GA_COEFFG') = 0 then TOBD.PutValue('GA_COEFFG',1);
          TOBD.PutValue('GA_DPR',Arrondi(TOBD.getValue('GA_PAHT')*TOBD.getValue('GA_COEFFG'),V_PGI.okdecP));
        end;
        if TOBD.getValue('GA_PAHT') <> 0 then
          TOBD.PutValue('GA_COEFFG',Arrondi(TOBD.getValue('GA_DPR')/TOBD.getValue('GA_PAHT'),4));
      end;
      //
      if (TOBD.GetValue('GA_CALCPRIXHT')='DPR') and (TOBD.GetValue('GA_CALCAUTOHT')='X') then
      begin
        TOBD.PutValue('GA_PVHT',Arrondi(TOBD.getValue('GA_DPR')*TOBD.getValue('GA_COEFCALCHT'),V_PGI.OkdecP));
      end;
    end;
    // partie vente modifie
    if ValoVenteHtModified (TOBD) then
    begin
      if (TOBD.GetValue('GA_CALCPRIXHT')='DPR') and (TOBD.GetValue('GA_CALCAUTOHT')='X') then
      begin
        if TOBD.IsFieldModified ('GA_COEFCALCHT') then
        begin
          if TOBD.getValue('GA_COEFCALCHT') = 0 then TOBD.PutValue('GA_COEFCALCHT',1);
          TOBD.PutValue('GA_PVHT',Arrondi(TOBD.getValue('GA_DPR')*TOBD.getValue('GA_COEFCALCHT'),V_PGI.okdecP));
        end;
        if TOBD.getValue('GA_DPR') <> 0 then
          TOBD.PutValue('GA_COEFCALCHT',Arrondi(TOBD.getValue('GA_PVHT')/TOBD.getValue('GA_DPR'),4));
      end else
      if (TOBD.GetValue('GA_CALCPRIXHT')='PAA') and (TOBD.GetValue('GA_CALCAUTOHT')='X') then
      begin
        if TOBD.IsFieldModified ('GA_COEFCALCHT') then
        begin
          if TOBD.getValue('GA_COEFCALCHT') = 0 then TOBD.PutValue('GA_COEFCALCHT',1);
          TOBD.PutValue('GA_PVHT',Arrondi(TOBD.getValue('GA_PAHT')*TOBD.getValue('GA_COEFCALCHT'),V_PGI.okdecP));
        end;
        if TOBD.getValue('GA_PAHT') <> 0 then
          TOBD.PutValue('GA_COEFCALCHT',Arrondi(TOBD.getValue('GA_PVHT')/TOBD.getValue('GA_PAHT'),4));
      end;
    end;
  end;
end;

procedure InsertOrUpdateTable(StSQL,Table,Cle : string);
var TobTemp,TobT : TOB;
    St,Champ,stValeur,stDate,stPrefixe : String;
    ListeChp,CleTob,critere,Clepart,StType : String;
    i,i_ind1,i_ind2 : integer;
    NbTob : integer;
    ChpDate : boolean;
    MajTobFilles : boolean;
    okInsert : boolean;
    QSelect : TQuery;
		Mcd : IMCDServiceCOM;

begin
MCD := TMCD.GetMcd;
if not mcd.loaded then mcd.WaitLoaded();

Clepart:=Cle;
MajTobFilles:=false;
okInsert:=false;
stPrefixe:=TableToPrefixe(table);
TobTemp:=TOB.Create(Table,nil,-1);
if (Table='SUSPECTSCOMPL') and
   (not ExisteSQL('SELECT RSC_SUSPECT FROM SUSPECTSCOMPL WHERE RSC_SUSPECT="'+Cle+'"')) then
      begin
      TobTemp.PutValue('RSC_SUSPECT',Cle);
      end else
if (Table='PROSPECTS') and
   (not ExisteSQL('SELECT RPR_AUXILIAIRE FROM PROSPECTS WHERE RPR_AUXILIAIRE="'+Cle+'"')) then
      begin
      TobTemp.PutValue('RPR_AUXILIAIRE',Cle);
      QSelect:=OpenSQL('SELECT T_TIERS FROM TIERS WHERE T_AUXILIAIRE="'+Cle+'"',false);
//      if not QSelect.Eof then TobTemp.Putvalue('RPR_TIERS',QSelect.FindField('T_TIERS').AsString);
      Ferme(QSelect);
      end else
if (Table='TIERSCOMPL') and
   (not ExisteSQL('SELECT YTC_AUXILIAIRE FROM TIERSCOMPL WHERE YTC_AUXILIAIRE="'+Cle+'"')) then
      begin
      TobTemp.PutValue('YTC_AUXILIAIRE',Cle);
      QSelect:=OpenSQL('SELECT T_TIERS FROM TIERS WHERE T_AUXILIAIRE="'+Cle+'"',false);
      if not QSelect.Eof then TobTemp.Putvalue('YTC_TIERS',QSelect.FindField('T_TIERS').AsString);
      Ferme(QSelect);
      end else

if (Table='ARTICLE') then
      begin
      // Si Article générique, mise à jour également des fiches articles dimensionnés
      ListeChp:=RechListeChamps(StSQL,stPrefixe);
      QSelect:=OpenSQL('SELECT * FROM ARTICLE WHERE GA_ARTICLE LIKE "'+Copy(Cle,1,18)+'%"',false);
      if not QSelect.Eof then
         begin
         TobTemp.LoadDetailDB('ARTICLE','','',QSelect,False,False);
         MajTobFilles:=true;
         end;
      Ferme(QSelect);
      end else
          begin
          // modif GA
          if (Table<>'CONTACT') and (Table<>'AFPLANNING') and (Table <> 'AFREVISION') then
              TobTemp.SelectDB('"'+Cle+'"',nil)
          else
              begin
              Repeat
                 critere:=ReadTokenSt (Clepart);
                 if CleTob <>'' then CleTob:=CleTob+';';
                 if critere<>'' then CleTob:=CleTob+'"'+critere+'"';
              until critere='';
              TobTemp.SelectDB(CleTob,nil);
              end;
          end;

While StSQL<>'' do
  begin
  St:=ReadTokenPipe(StSQL,'|');
  i_ind1:=Pos('=',St); if i_ind1=0 then continue;
  Champ:=Trim(Copy(St,1,i_ind1-1)); stValeur:=Trim(Copy(St,i_ind1+1,Length(St)));
  i_ind2:=Pos('_',Champ);
  ChpDate:=false;
  if copy(Champ,1,i_ind2-1)=stPrefixe then
     begin
		 sttype := mcd.getField(Champ).tipe;
//     StType:=V_PGI.DEChamps[TableToNum(Table),TobTemp.GetNumChamp(Champ)].tipe;
     if StType = 'DATE' then
        begin
        stDate:='';
        for i_ind1:=1 to length(stValeur) do
            if stValeur[i_ind1]<>'"' then stDate:=stDate+stValeur[i_ind1];
        ChpDate:=true;
        end;
     if MajTobFilles=true then NbTob := TobTemp.Detail.Count
                          else NbTob := 1;
     for i:=0 to NbTob-1 do
      begin
        if MajTobFilles=true then TobT := TobTemp.Detail[i]
                             else TobT := TobTemp;
        if ChpDate=true then TobT.PutValue(Champ,StrToDate(stDate))
        else
          begin
          if ((StType = 'DOUBLE') or (StType = 'INTEGER') or (StType = 'EXTENDED') or (StType = 'RATE')) and
                ( stValeur='' ) then
             TobT.PutValue(Champ,0)
          // C.B correction conversion variant incorrect si ,
          else if ((StType = 'DOUBLE') or (StType = 'EXTENDED') or (StType = 'RATE')) and ( stValeur<>'' ) then
             TobT.PutValue(Champ, valeur(stValeur))
          else
             TobT.PutValue(Champ, stValeur);
          end;
        // Mise à jour des champs systèmes _DATEMODIF et _UTILISATEUR
        // DCA - FQ MODE 10923 - Test existence sur TobT au lieu de TobTemp
        if TobT.FieldExists ( stPrefixe + '_DATEMODIF')
          then TobT.PutValue(stPrefixe + '_DATEMODIF', NowH ) ;
        if ( stPrefixe <> 'US' ) and ( TobT.FieldExists ( stPrefixe + '_UTILISATEUR') )
          then TobT.PutValue(stPrefixe + '_UTILISATEUR', V_PGI.User ) ;
      end;
     OkInsert:=true;
     end;
  end;
if OkInsert then
begin
  FinalizeMajFromChamp(TOBTemp);
  TobTemp.UpdateDB(false);
end;
TobTemp.Free;
end;

function RechListeChamps(StSQL,stPrefixe : string) : string;
var St,Champ : String;
    i_ind1,i_ind2 : integer;
begin
Result := '';
While StSQL<>'' do
  begin
  St:=ReadTokenPipe(StSQL,'|');
  i_ind1:=Pos('=',St); if i_ind1=0 then continue;
  Champ:=Trim(Copy(St,1,i_ind1-1));
  i_ind2:=Pos('_',Champ);
  if copy(Champ,1,i_ind2-1)=stPrefixe then
     begin
     if Result='' then Result:=Champ
     else Result:=Result+','+Champ;
     end;
  end;
  if Result='' then Result:='*';
end;

function VerifPrefixe(Table, Code : string) : string;
var stPrefixe : string;
    i_ind : integer;
begin
Result := Code;
stPrefixe:=TableToPrefixe(Table);
i_ind:=Pos('_',Code);
if i_ind<>0 then
   if stPrefixe<>Trim(Copy(Code,1,i_ind-1)) then
      Result:=stPrefixe+Copy(Code,i_ind,Length(Code));
end;

(*======================================================================*)

procedure ModifieRafale(Obj : TO_ModifParLot; Table :string) ;
var st1Table,st2Table, stCode, Cle, Champs, critere  : String;
    i : integer;
begin
With Obj Do
  begin
  if F.AllSelected then
    begin
    Q.DisableControls ;
    Q.First ;
    While not Q.Eof do
      begin
      st1Table:=Table;
      // GRC
      //Cle := Q.FindField(FCode).AsString;
      Cle:='';
      Champs:=FCode;
      repeat
        critere:=ReadTokenSt(Champs);
        if critere = '' then break;
        if Cle <> '' then Cle:=Cle+';';
        Cle := Cle+Q.FindField(critere).AsString;
      until critere='';

      st2Table:=ReadTokenSt(st1Table);
      While st2Table<>'' do
          begin
          stCode:=VerifPrefixe(st2Table,FCode);
          InsertOrUpdateTable(StSQL,st2Table,cle) ;
          st2Table:=ReadTokenSt(st1Table);
          end;
      Q.Next ;
      end;
    end
  else
    begin
    for i:=0 to F.NbSelected-1 do
      begin
      F.GotoLeBookMark(i) ;
{$IFDEF EAGLCLIENT}
      Q.TQ.Seek(F.Row-1) ;
{$ENDIF}
      st1Table:=Table;
      // GRC
      //Cle := Q.FindField(FCode).AsString;
      Cle:='';
      Champs:=FCode;
      repeat
        critere:=ReadTokenSt(Champs);
        if critere = '' then break;
        if Cle <> '' then Cle:=Cle+';';
        Cle := Cle+Q.FindField(critere).AsString;
      until critere='';

      st2Table:=ReadTokenSt(st1Table);
      While st2Table<>'' do
          begin
          stCode:=VerifPrefixe(st2Table,FCode);
          InsertOrUpdateTable(StSQL,st2Table,cle) ;
          st2Table:=ReadTokenSt(st1Table);
          end;
      end;
    end;
    Q.EnableControls ;
  end;
end;

(*======================================================================*)

procedure ModifieFicheParFiche(Obj : TO_ModifParLot) ;
var i : integer;
    Cle,critere,Champs,stArg,cle2,LocalStSQL : string;
begin
With Obj Do
  begin
  if F.AllSelected then
    begin
    Q.DisableControls ;
    Q.First ;
    While not Q.Eof do
      begin
      // GRC
      //Cle := Q.FindField(FCode).AsString;
      Cle:=''; cle2:='';
      Champs:=FCode;

      repeat
        critere:=ReadTokenSt(Champs);
        if critere = '' then break;
        if Cle <> '' then
           begin
           Cle:=Cle+';';
           if cle2='' then cle2:=Q.FindField(critere).AsString;
           end;
        Cle := Cle+Q.FindField(critere).AsString;
      until critere='';
      if cle2='' then cle2:=Cle;

      stArg:='';
      LocalStSQL:=StSQL;
{$IFDEF GRC}
      if (ctxGRC in V_PGI.PGIContexte) and (NatureTiers <> 'FOU') and (NatureTiers <> 'SUS') then
         if (RTDroitModiftiers(TiersAuxiliaire(cle2,True))=False) then
            begin
            stArg:= 'ACTION=CONSULTATION;';
            LocalStSQL:='';
            end;
{$ENDIF}            
      if NatureTiers = 'PRO' then
         begin
         if AGLlancefiche(Nature,FicheAOuvrir,'','',stArg+Cle+';MODIFLOT;'+LocalStSQL)='Stop' then Break;
         end
      else
         if NatureTiers = 'CON' then
            begin
            if AGLlancefiche(Nature,FicheAOuvrir,Cle,'',stArg+'MODIFLOT;'+LocalStSQL)='Stop' then Break;
            end
         else
            begin
            if AGLlancefiche(Nature,FicheAOuvrir,'',Cle,stArg+'MODIFLOT;'+LocalStSQL)='Stop' then Break;
            end;

      //ACTION=MODIFICATION;MONOFICHE;
      Q.Next ;
      end;
    end
  else
    begin
    for i:=0 to F.NbSelected-1 do
      begin
      F.GotoLeBookmark(i) ;
{$IFDEF EAGLCLIENT}
      Q.TQ.Seek(F.Row-1) ;
{$ENDIF}
      // GRC
      //Cle := Q.FindField(FCode).AsString;
      Cle:=''; cle2:='';
      Champs:=FCode;
      repeat
        critere:=ReadTokenSt(Champs);
        if critere = '' then break;
        if Cle <> '' then
           begin
           Cle:=Cle+';';
           if cle2='' then cle2:=Q.FindField(critere).AsString;
           end;
        Cle := Cle+Q.FindField(critere).AsString;
      until critere='';
      if cle2='' then cle2:=Cle;

      stArg:='';
      LocalStSQL:=StSQL;
{$IFDEF GRC}
      if (ctxGRC in V_PGI.PGIContexte) and (NatureTiers <> 'FOU') and (NatureTiers <> 'SUS') then
         if (RTDroitModiftiers(TiersAuxiliaire(cle2,True))=False) then
            begin
            stArg:= 'ACTION=CONSULTATION;';
            LocalStSQL:='';
            end;
{$ENDIF}            
      if NatureTiers = 'PRO' then
         begin
         if AGLlancefiche(Nature,FicheAOuvrir,'','',stArg+Cle+';MODIFLOT;'+LocalStSQL)='Stop' then Break;
         end
      else
         if NatureTiers = 'CON' then
            begin
            if AGLlancefiche(Nature,FicheAOuvrir,Cle,'',stArg+'MODIFLOT;'+LocalStSQL)='Stop' then Break;
            end
         else
            begin
            if AGLlancefiche(Nature,FicheAOuvrir,'',Cle,stArg+'MODIFLOT;'+LocalStSQL)='Stop' then Break;
            end;
      end;
    end;
    Q.EnableControls;
  end;
end;

(*======================================================================*)

procedure ModifieEnSerie(ObjModif : TO_ModifParLot; Parametres : String);
var OkRaf,OkRep : boolean;
    StSQL,Table,Titre,NatureTiers : string;
    // Modif BTP
    Contexte : string;
begin
//{$IFDEF EAGLCLIENT}
// AFAIREEAGL
//{$ELSE}
if ObjModif.FCode='' then Exit ;
NatureTiers:=ObjModif.NatureTiers;
if ObjModif.TableName<>'' then
   Table:=ObjModif.TableName
   else
{$IFNDEF EAGLCLIENT}
   Table:=GetTableNameFromDataSet(ObjModif.Q);
{$ELSE}
   begin
   ObjModif.Free;
   exit;                     
   end;
{$ENDIF}

Titre := ObjModif.Titre;
// Modif BTP
Contexte := OBJModif.StContexte;

//modif GA
if OBJModif.FicheAOuvrir = '' then
  OkRaf := True
else
  OkRaf := False;

StSQL:=ModifZoneSerie(Table,Titre,NatureTiers,OkRaf,OkRep,Contexte);
// ------
if StSQL='' then
   begin
   ObjModif.Free;
   exit;
   end;

ObjModif.StSQL := StSQL;
if OkRaf then
   begin
   if OkRep then ModifieRafale(ObjModif,Table);
   end
else
   begin
   ModifieFicheParFiche(ObjModif);
   end;
ObjModif.Free;
//{$ENDIF}
end;

(*======================================================================*)

end.
