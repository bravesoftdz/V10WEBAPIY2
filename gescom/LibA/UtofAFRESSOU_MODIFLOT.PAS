{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 28/09/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : AFRESSOU_MODIFLOT ()
Mots clefs ... : TOF;AFRESSOU_MODIFLOT
*****************************************************************}
Unit UTOFAFRESSOU_MODIFLOT ;

Interface

uses  StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
      eMul,  Maineagl,
{$ELSE}
      db,dbTables,mul,FE_Main ,
{$ENDIF}
      forms, sysutils, GCMZSUtil, Utob, UtilMultrt, Windows, Hstatus,
      ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF,
      DicoAF, SaisUtil, EntGC, M3FP, ParamSoc,
      utilpgi,AglInit,UtilGc,TraducAffaire,UtofAfTraducChampLibre;

Type
  TOF_AFRESSOU_MODIFLOT = Class (TOF_AFTRADUCCHAMPLIBRE) // mcd 05/06/02 pour traduciton ne doit pas descendre de TOF
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
    Type_Trait:string ;
    procedure ModifParLotDesRessources;
    procedure ModifAvecTrait;
  end ;
Procedure AFLanceFiche_ModifLog_Ress(TypeTrai:string);

Implementation


procedure TOF_AFRESSOU_MODIFLOT.OnArgument (S : String ) ;
Var stArg:TStringList;
begin
  Inherited ;
  stArg:=TStringList.Create;
  try
   stArg.Text:=S;
   Type_Trait:=stArg.values['TRAIT'];
   finally
   stArg.free;
   end;
  If Type_Trait='TRA' then
    begin
      Ecran.Caption := TraduitGa('Modification en série des ressources avec traitement');
      updatecaption(Ecran);
    end;
end ;

procedure TOF_AFRESSOU_MODIFLOT.OnClose ;
begin
  Inherited ;
V_PGI.ExtendedFieldSelection:='' ; //mcd 06/10/2003 10912
end ;

procedure TOF_AFRESSOU_MODIFLOT.ModifAvecTrait;
Var TobRes, TobDet: TOB;
    F : TFMul ;
    select,stmp,Sql,Ori,Des, tmp : string;
    ret:T_ErrTraiteEnreg;
    Recep : TStringList;
    wi,wj:integer;
    PV,PR : double;
    Date : TdateTime;
    TopPr,TopPV : boolean;
begin
  F:=TFMul(Ecran);
  TobRes:= Tob.Create('Les ress',NIL,-1);
  // on sélectionne les champs à mettre à jour
  Select := AGLLanceFiche('AFF','AFMODIFRES_TRAIT','','','');
  if (Select='') then
    begin
     TobRes.free;
     exit;
    end;
  Sql := 'SELECT ARS_RESSOURCE, ARS_CALCULPV, ARS_CALCULPR, ARS_DATEMODIF';
  Sql := Sql +', ARS_CALENSPECIF, ARS_STANDCALEN, ARS_TAUXUNIT, ARS_TAUXFRAISGEN1';
  Sql := Sql +', ARS_TAUXFRAISGEN2, ARS_TAUXCHARGEPAT, ARS_COEFMETIER, ARS_COEFPRPV';
  Sql := Sql +', ARS_TAUXREVIENTUN, ARS_PVHTCALCUL, ARS_DATEPRIX, ARS_PVHT, ARS_PVTTC';
  Sql := Sql +', ARS_TAUXREVIENTUN2, ARS_PVHTCALCUL2, ARS_DATEPRIX2, ARS_PVHT2, ARS_PVTTC2';
  Sql := Sql +', ARS_TAUXREVIENTUN3, ARS_PVHTCALCUL3, ARS_DATEPRIX3, ARS_PVHT3, ARS_PVTTC3';
  Sql := Sql +', ARS_TAUXREVIENTUN4, ARS_PVHTCALCUL4, ARS_DATEPRIX4, ARS_PVHT4, ARS_PVTTC4';
  Sql := Sql +' FROM RESSOURCE' ;
  ret:=TraiteEnregMulTable (F, Sql ,'ARS_RESSOURCE', 'RESSOURCE','ARS_RESSOURCE','RESSOURCE',TobRes, True);
  If ret <> tteOK then exit ;
  sTmp := StringReplace(Select, ';', chr(VK_RETURN), [rfReplaceAll]);
  Recep := TStringList.Create;
  Recep.Text := sTmp;
  InitMove(TobRes.Detail.count*2,' ');
  wi := 0;
  while wi < TobRes.Detail.count  do
      Begin
      TobDet := TobRes.detail[wi];
      PR :=   TObDEt.Getvalue('ARS_TAUXREVIENTUN'); // stocke ancien prix
      PV :=   TObDEt.Getvalue('ARS_PVHTCALCUL'); // stocke ancien prix
      Date:=   TObDEt.Getvalue('ARS_DatePrix'); // stocke ancien prix
      MoveCur(False);
      TopPR :=False; // si on doit calculer le PR
      TopPV :=False; // si on doit calculer le PV
      if  (recep.Values['COUT']<>'0')
        and (TobDEt.GetValue('ARS_TAUXUNIT')<>StrToFloat(recep.Values['COUT'])) then
       begin // modif Cout unitaire. on peut reclauler le pR
       TobDet.putValue('ARS_TAUXUNIT', StrToFloat(recep.Values['COUT']));
       TopPr:=True;
       end;
      if  (recep.Values['FRAIS1']<>'0')
        and (TobDEt.GetValue('ARS_TAUXFRAISGEN1')<>StrToFloat(recep.Values['FRAIS1'])) then
       begin
       TobDet.putValue('ARS_TAUXFRAISGEN1', StrToFloat(recep.Values['FRAIS1']));
       TopPr:=True;
       end;
      if  (recep.Values['FRAIS2']<>'0')
        and (TobDEt.GetValue('ARS_TAUXFRAISGEN2')<>StrToFloat(recep.Values['FRAIS2'])) then
       begin
       TobDet.putValue('ARS_TAUXFRAISGEN2', StrToFloat(recep.Values['FRAIS2']));
       TopPr:=True;
       end;
      if  (recep.Values['METIER']<>'0')
        and (TobDEt.GetValue('ARS_COEFMETIER')<>StrToFloat(recep.Values['METIER'])) then
       begin
       TobDet.putValue('ARS_COEFMETIER', StrToFloat(recep.Values['METIER']));
       TopPr:=True;
       end;
      if  (recep.Values['CHARGE']<>'0')
        and (TobDEt.GetValue('ARS_TAUXCHARGEPAT')<>StrToFloat(recep.Values['CHARGE'])) then
       begin
       TobDet.putValue('ARS_TAUXCHARGEPAT', StrToFloat(recep.Values['CHARGE']));
       TopPr:=True;
       end;
      if  (recep.Values['COEFPRPV']<>'0')
        and (TobDEt.GetValue('ARS_COEFPRPV')<>StrToFloat(recep.Values['COEFPRPV'])) then
       begin
       TobDet.putValue('ARS_COEFPRPV', StrToFloat(recep.Values['COEFPRPV']));
       TopPV:=True;
       end;
      if  (recep.Values['PVCAL']<>'0')
        and (TObDet.Getvalue('ARS_CALCULPV')='-')
        and (TobDEt.GetValue('ARS_PVHTCALCUL')<>StrToFloat(recep.Values['PVCAL'])) then
       begin // modif Cout unitaire. on peut reclauler le pv
       TobDet.putValue('ARS_PVHTCALCUL', StrToFloat(recep.Values['PVCAL']));
       end;
      if  (recep.Values['PRUNI']<>'0')
        and (TObDet.Getvalue('ARS_CALCULPR')='-')
        and (TobDEt.GetValue('ARS_TAUXREVIENTUN')<>StrToFloat(recep.Values['PRUNI'])) then
       begin // modif Cout unitaire. on peut reclauler le pv
       TobDet.putValue('ARS_TAUXREVIENTUN', StrToFloat(recep.Values['PRUNI']));
       TopPV:=True; // si le taux unitaire à changer, il faut recalcule le PV si auto
       end;
      if (recep.Values['PR']='1') then
         begin  // mise à jour cahmps CALCULPR SI Vrai, on recalcule le prix untitaire
           if  (recep.Values['CALPR']='1') then
             begin  // si modif voir fct Ressource_CalculPR UtomRessource
             TObDet.putvalue('ARS_CALCULPR','X');
             TopPr:=True;
             end
           else TObDet.putvalue('ARS_CALCULPR','-');
         end;
      if (recep.Values['PV']='1') then
         begin  // mise à jour cahmps CALCULPV SI Vrai, on recalcule le prix untitaire
           if  (recep.Values['CALPV']='1') then
             begin  // si modif voir fct Ressource_CalculPV UtomRessource
             TObDet.putvalue('ARS_CALCULPV','X');
             TopPV:=True;
             end
           else TObDet.putvalue('ARS_CALCULPV','-');
         end;
      if (recep.Values['STD']<>'') then
         begin  // code calendrier, si pour la ressource, le calendrier n'est pas spécif
         Tmp :=  recep.Values['STD'];
         if tmp ='@' then tmp :=''; //mcd 01/07/2003 pour remettre à blanc
         if TobDet.Getvalue('ARS_CALENSPECIF')='-' then TobDet.putValue('ARS_STANDCALEN',tmp);
         end;
      If (TopPr) and (TObDet.Getvalue('ARS_CALCULPR')='X') then
         begin  // recalul PR qui a changé
         TObDEt.putvalue('ARS_TAUXREVIENTUN',Arrondi ((TobDet.GetValue('ARS_TAUXUNIT')* TobDet.GetValue('ARS_TAUXFRAISGEN1')*
                    TobDet.GetValue('ARS_TAUXFRAISGEN2')* TobDet.GetValue('ARS_TAUXCHARGEPAT')* TobDet.GetValue('ARS_COEFMETIER'))
                    ,GetParamSoc('SO_DECPRIX')));
         end;
      If ((TopPV) or(TopPR)) and (TObDet.Getvalue('ARS_CALCULPV')='X') then  // il faut recalculer le PV, si le taux unitaire a aussi été changé ..
         begin  // recalul PV qui a changé
         TObDet.putvalue('ARS_PVHTCALCUL',Arrondi ((TobDet.GetValue('ARS_COEFPRPV')* TobDet.GetValue('ARS_TAUXREVIENTUN'))
                    ,GetParamSoc('SO_DECPRIX')));
         end;
      if (recep.Values['HISTO']='1') and ((TopPR) or (TopPV)) then
         begin    // mcd attention si modif voir fct historisation dans UtomRessource
       //mcd 04/04/03  TObDet.putvalue('ARS_DATEPRIX', V_PGI.DateEntree);
         TObDet.putvalue('ARS_DATEPRIX', StrToDate(Recep.values['DATEMAJ']));
         If TObDEt.Getvalue('ARS_DATEPRIX') > TObDEt.Getvalue('ARS_DATEPRIX2') then
           begin
           for wj:=0 to 2 do
             begin
             if wj= 0 then begin Ori:='3'; Des:='4' end
               else if wj= 1 then begin Ori:='2'; Des:='3' end
               else begin Ori:=' '; Des:='2' end;
             If wj <> 2 then begin     // on passe 3 ds 4, 2 ds 3 ...
               TObDet.putvalue('ARS_DATEPRIX'+Des, TObDet.Getvalue('ARS_DATEPRIX'+Ori));
               TObDet.putvalue('ARS_TAUXREVIENTUN'+Des, TObDet.Getvalue('ARS_TAUXREVIENTUN'+Ori));
               TObDet.putvalue('ARS_PVHTCALCUL'+Des, TObDet.Getvalue('ARS_PVHTCALCUL'+Ori));
               end
             else begin  // pour le prix le plus récent, il faut mettre l'ancienne valeur de la tob ...
               TObDet.putvalue('ARS_DATEPRIX'+Des, Date);
               TObDet.putvalue('ARS_TAUXREVIENTUN'+Des, PR);
               TObDet.putvalue('ARS_PVHTCALCUL'+Des, PV);
               end;
                   // ces prix ne bougent pas dans cette fct, on les agrde tel quel
             TObDet.putvalue('ARS_PVHT'+Des, TObDet.Getvalue('ARS_PVHT'+Ori));
             TObDet.putvalue('ARS_PVTTC'+Des, TObDet.Getvalue('ARS_PVTTC'+Ori));
             end;
           end;
         end;
      TObDet.putvalue('ARS_DATEMODIF', V_PGI.DateEntree);
      inc(wi);
      End;  // fin while  ,boucle sur la sélection
  Recep.Free;
  TobRes.UpdateDb(False,False);
  TobRes.free;
  FiniMove;
end ;


procedure TOF_AFRESSOU_MODIFLOT.ModifParLotDesRessources;
Var F : TFMul ;
    Parametrages : String;
    TheModifLot : TO_ModifParLot;
begin
F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
  begin
  MessageAlerte('Aucun élément sélectionné');
  V_PGI.ExtendedFieldSelection:='' ;   //mcd 06/10/2003 10912
  exit;
  end;
If Type_Trait='TRA' then
   begin
   ModifAvecTrait;  // cas modif de zones ressource qui demande un traitement
   exit;
   end;
{$IFDEF EAGLCLIENT}
if F.bSelectAll.Down then
   if not F.Fetchlestous then
     begin
     F.bSelectAllClick(Nil);
     F.bSelectAll.Down := False;
     exit;
     end;
{$ENDIF}    
TheModifLot := TO_ModifParLot.Create;
TheModifLot.F := F.FListe;
TheModifLot.Q := F.Q;
TheModifLot.NatureTiers := '';
TheModifLot.Nature := 'AFF';
TheModifLot.Titre := Ecran.Caption;
TheModifLot.TableName:='RESSOURCE';
TheModifLot.FCode := 'ARS_RESSOURCE';
TheModifLot.FicheAOuvrir := 'RESSOURCE';
V_PGI.ExtendedFieldSelection:='1';

ModifieEnSerie(TheModifLot, Parametrages) ;
V_PGI.ExtendedFieldSelection:='' ;   //mcd 06/10/2003 10912
end;

/////////////// Procedure appeleé par le bouton Validation //////////////
procedure AGLModifParLotDesRessources(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is TOF_AFRESSOU_MODIFLOT) then TOF_AFRESSOU_MODIFLOT(TOTOF).ModifParLotDesRessources else exit;
end;



Procedure AFLanceFiche_ModifLog_Ress(TypeTrai:string);
begin
AGLLanceFiche ('AFF','AFRESSOU_MODIFLOT','','','TRAIT='+TypeTrai);
end;

Initialization
  registerclasses ( [ TOF_AFRESSOU_MODIFLOT ] ) ;
  RegisterAglProc('ModifParLotDesRessources',TRUE,0,AGLModifParLotDesRessources);
end.
