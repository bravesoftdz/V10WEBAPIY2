unit ZGuide;

//=======================================================
//=================== Clés primaires ====================
//=======================================================
// GUIDE  : GU_TYPE, GU_GUIDE
// ECRGUI : EG_TYPE, EG_GUIDE, EG_NUMLIGNE
// ANAGUI : AG_TYPE, AG_GUIDE, AG_NUMLIGNE, AG_NUMVENTIL, AG_AXE

interface

uses SysUtils,
    {$IFNDEF EAGLCLIENT}
      {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
    {$ENDIF}

    {$IFDEF VER150}
      Variants,
    {$ENDIF}

     UTob,
     HEnt1 ,
     HCtrls ,
     Ent1 ,
     hmsgbox ,
     formule,   // pour GFormule
     UTOFLOOKUPTOB,
     ULibEcriture ;


type

 TTypeRecherche  = ( TColGeneral,TColDebit,TColCredit,TColLibelle,TColTVA, TColDate, TColRefInterne, TColRefExterne, TColCodeAFB , TColNaturePiece , TColJournal , TColAuxi);
 TTypeRecherches = set of TTypeRecherche;

{ TProcAction     = function : string of object;

 TRecRech       = record
  Value         : string;
  ProcAction    : TProcAction;
 end;


 PRecLib       = ^TRecLib ;
 TRecLib       = record
  Value         : string;
  Prio          : integer;
  PTOB          : TOB ;
 end;
}

 TZGuide = Class(TObjetCompta)
  private
   FStPrefixe          : string ;
   FInNumLigne         : integer;
   FStEG_GUIDE         : string;          // code du guide courant
   FListe              : THGrid ;
   FTOBResult          : TOB ;             // TOB genere par le moteur des guide
   FTOB                : TOB ;
   FTOBLigne           : TOB ;
   FInIndex            : integer ;
   FInOffset           : integer ;
   FInNumDep           : integer ;
   FInNbLigneGuide     : integer ;
   function    MoteurCalculGuide (vTOBLigneEcr,vTOBLigneGuide : TOB ) : boolean;
   function    CalculeGuide : boolean ;
   function    GetFormule( StFormule : hstring ) : variant ;
   procedure   CalculDebitCredit ( vTOBLigneEcr,vTOBLigneGuide : TOB ; var vRdDebit : double ; var vRdCredit : double ) ;
   function    CalculDuSolde : double ;
    protected
   procedure   SetCrit(vTOB : TOB; const Values : array of string) ; virtual ;
   function    FindFirst : TOB ; virtual ;
   function    FindNext : TOB ; virtual ;
  public
   constructor Create( vInfoEcr : TInfoEcriture ) ; override ;
   destructor  Destroy; override;

   function    GetValue(Name : string ; Index : integer=-1) : Variant ;
   function    Load(const Values : array of string ) : integer ; virtual ;
   function    LoadLigne(const Values : array of string ; vNumLigne : integer ) : TOB ; virtual ;
   procedure   RecalculGuide ( vTOB : TOB ; ARow : integer  ) ;
   function    RechercheGuideEnBase( vTOB,vTOBResult : TOB ; G : THGrid = nil ) : boolean ;

   property    Item           : TOB     read FTOBLigne ;
   property    Items          : TOB     read FTOB ;
   property    InNumDep       : integer read FInNumDep ;
   property    InNbLigneGuide : integer read FInNbLigneGuide ;

 end;


implementation

const
 cStIdCodeAFB     = '$';
 cStIdCodeAFBSupp = '&' ;
 cStJocker        = '*' ;

function CGetRowFormule ( var StFormule : hstring ) : Integer ;
var
 Pos1,
 Pos2 : Integer ;
begin

 Result := 0;

 Pos2 := Pos(':L', StFormule);
 if Pos2 > 0 then
  System.Delete( StFormule, Pos2+1 , 1 );

 Pos1 := Pos(':',  StFormule);
 if ( Pos1 <= 0 ) or ( Pos1 = Length(StFormule)) then
  Exit ;

 Result := Round(Valeur(Copy(StFormule, Pos1+1, 5))) ;

 System.Delete(StFormule, Pos1, 5) ;

end ;

constructor TZGuide.Create ( vInfoEcr : TInfoEcriture ) ;
begin
 inherited Create(vInfoEcr) ;
 FTOB          := TOB.Create('',nil,-1) ;
 FInIndex      := -1 ;
 FStPrefixe    := 'E' ;
end;

destructor TZGuide.Destroy;
begin
 try
  if assigned(FTOB) then FTOB.Free ;
 finally
  FTOB := nil ;
 end;
 inherited ;
end;


function TZGuide.GetValue(Name : string ; Index : integer=-1): Variant;
begin
 Result:=#0 ;
 if (Index<>-1) and (Index<=FTOB.Detail.Count - 1) then FInIndex:=Index ;
 if (FTOB<>nil) and (FInIndex<>-1) then Result:=FTOB.Detail[FInIndex].GetValue(Name) ;
end;

function TZGuide.LoadLigne(const Values : array of string ; vNumLigne : integer ) : TOB ;
begin
 result := nil ;
 if Load(Values) = - 1 then exit ;
 if ( vNumLigne <= Item.detail.Count ) and ( vNumLigne > - 1 ) then
  result := item.Detail[vNumLigne - 1] ;
end ;

function TZGuide.Load(const Values : array of string) : integer;
var
 lStCode : string ;
 i       : integer ;
 lTOB    : TOB ;
begin

 Result:=-1 ;
 FTOBLigne := nil ;
 if (SizeOf(Values)=0) then exit ; lStCode:='' ;
 for i:=low(Values) to High(Values) do lStCode:=lStCode+Values[i] ; lStCode:=UpperCase(lStCode) ;
 if GetValue('CRIT')=lStCode then // on commence directement par l'index
  begin
   Result          := FInIndex ;
   FTOBLigne       := FTOB.Detail[FInIndex] ;
   FInNbLigneGuide := FTOBLigne.Detail.Count ;
   Exit ;
  end
   else
    for i:=0 to FTOB.Detail.Count-1 do
     if FTOB.Detail[i].GetValue('CRIT')=lStCode then
      begin
       FInIndex        := i ;
       Result          := FInIndex ;
       FTOBLigne       := FTOB.Detail[FInIndex] ;
       FInNbLigneGuide := FTOBLigne.Detail.Count ;
       Exit ;
      end ;

 lStCode := '' ;

 for i := low(Values) to High(Values) do
  lStCode := lStCode + '"' + Values[i] + '";';

 lTOB     := TOB.CreateDB('GUIDE', FTOB , -1, nil) ;
 lTOB.AddChampSupValeur('CRIT',lStCode) ;
 lTOB.SelectDB(lStCode , nil) ;
 lTOB.LoadDetailDB('ECRGUI', lStCode , '', nil , false ) ;

 if lTOB.Detail.Count = 0 then exit ;

 for i := 0 to lTOB.Detail.Count - 1 do
  begin
   lTOB.Detail[i].AddChampSupValeur('_JOURNAL'       , lTOB.GetValue('GU_JOURNAL') ) ;
   lTOB.Detail[i].AddChampSupValeur('_NATUREPIECE' ,   lTOB.GetValue('GU_NATUREPIECE') ) ;
  end ;
 SetCrit(lTOB,Values) ;
 FInIndex             := FTOB.Detail.Count-1 ;
 FTOBLigne            := lTOB ;
 FInNbLigneGuide      := FTOBLigne.Detail.Count ;
 result               := FInIndex ;

end;

procedure TZGuide.SetCrit(vTOB : TOB; const Values : array of string);
var
 lStCode : string ;
 i       : integer ;
begin
 for i:=low(Values) to High(Values) do lStCode:=lStCode+Values[i] ; lStCode:=UpperCase(lStCode) ;
 vTOB.PutValue('CRIT',lStCode) ;
end;


function TZGuide.GetFormule( StFormule : hstring) : Variant ;
var
 lValue     : variant ;
 lInCalcRow : integer ; // ligne sur laquel s'applique la formule
 lInRefRow  : integer ;
 lTOB       : TOB ;
begin

 result    := #0 ;
 lValue    := #0 ;
 StFormule := UpperCase(StFormule) ;

 if ( Pos('E_', StFormule) = 0 ) and ( StFormule <> 'SOLDE' ) then exit ;

 StFormule := UpperCase(Trim(StFormule));

 if StFormule = '' then Exit ;

 // on recupere la ligne d'application de la formule
 lInCalcRow := CGetRowFormule(StFormule) ;
 if lInCalcRow < 0 then
  lInRefRow := FInNumLigne + lInCalcRow + FInOffset - 1 
   else
    lInRefRow := lInCalcRow + FInOffset - 2 ;

 if (lInRefRow < 0 ) or ( lInRefRow > FTOBResult.Detail.Count - 1 ) then
  lInRefRow := 0 ;

 lTOB := FTOBResult.Detail[lInRefRow] ;

 if StFormule = 'SOLDE' then
  lValue := 'SOLDE'
   else
    lValue := VarToStr( lTOB.GetValue( StFormule) ) ;

 result := lValue;

end;


function  TZGuide.CalculDuSolde : double ;
var
 i                : integer ;
 lRdSoldeDebit    : double ;
 lRdSoldeCredit   : double ;
begin

 lRdSoldeDebit    := 0 ;
 lRdSoldeCredit   := 0 ;

 for i := FInNumDep-1 to FTOBResult.Detail.Count - 2 do
  begin
   lRdSoldeDebit  := lRdSoldeDebit  + FTOBResult.Detail[i].GetValue('E_DEBIT') ;
   lRdSoldeCredit := lRdSoldeCredit + FTOBResult.Detail[i].GetValue('E_CREDIT') ;
  end; // for

 result           := Arrondi(lRdSoldeDebit - lRdSoldeCredit , V_PGI.OkDecV ) ;

end ;


procedure TZGuide.CalculDebitCredit ( vTOBLigneEcr,vTOBLigneGuide : TOB ; var vRdDebit : double ; var vRdCredit : double ) ;
var
 lStValeurCell : string ;
 lRdSolde      : double ;
 lStMess : string;

 function IsNumericLG ( Value : string ) : boolean;
 begin
  result := ( pos('-', Copy( Value, 2 , length(Value) ) ) = 0 ) and ( pos('/',Value) = 0 ) and ( pos('*',Value) = 0 ) and ( pos('+',Value) = 0 );
 end;

 function _GetMontant ( vStChamps1 ,vStChamps2 : string ) : double ;
  begin
   result := 0 ;
   if vTOBLigneGuide.GetValue(vStChamps1) <> '' then
    begin
     if IsNumeric(lStValeurCell) and not IsNumericLG(lStValeurCell) then
       begin
         lStMess := TraduireMemoire('Erreur de formule dans le Guide n°')          + FStEG_GUIDE
                     + ' ' + RechDom('TTGUIDEECR', FStEG_GUIDE ,false)              + #13#10
                     + TraduireMemoire('La formule ') + vTOBLigneGuide.GetValue(vStChamps1) + TraduireMemoire(' est incorrecte');
         NotifyError( lStMess ,'', '' );
         exit;
       end; // if
     result := Valeur(lStValeurCell) ;
    end
     else
      if vTOBLigneEcr.GetValue( vStChamps2) <> 0 then
       result := vTOBLigneEcr.GetValue( vStChamps2) ;
  end ; // function

begin

 vRdDebit  := 0 ;
 vRdCredit := 0 ;

 if vTOBLigneGuide.GetValue('EG_DEBITDEV') <> '' then
  lStValeurCell  := GFormule(vTOBLigneGuide.GetValue('EG_DEBITDEV'), GetFormule, nil, 1)
   else
    if vTOBLigneGuide.GetValue('EG_CREDITDEV') <> '' then
     lStValeurCell  := GFormule(vTOBLigneGuide.GetValue('EG_CREDITDEV'), GetFormule, nil, 1) ;

// if lStValeurCell = '' then exit ;

 if lStValeurCell = 'SOLDE' then
  begin

   lRdSolde := CalculDuSolde ;
   if vTOBLigneGuide.GetValue('EG_DEBITDEV') <> '' then
    begin
     if lRdSolde <= 0 then
      vRdDebit := lRdSolde * (-1)
       else
        vRdCredit := lRdSolde ;
     exit ;
    end ;

   if vTOBLigneGuide.GetValue('EG_CREDITDEV') <> '' then
    begin
     if lRdSolde <= 0 then
      vRdDebit := lRdSolde * (-1)
       else
        vRdCredit := lRdSolde ;
     exit ;
    end ;

  end ; // if solde

 vRdDebit  :=  _GetMontant('EG_DEBITDEV'  ,'E_DEBIT') ;
 vRdCredit :=  _GetMontant('EG_CREDITDEV' ,'E_CREDIT') ;

 // gestion des montants negatifs
 if ( vRdDebit < 0 ) and not( VH^.MontantNegatif ) then
  begin
   vRdCredit := vRdDebit * (-1);
   vRdDebit  := 0;
  end // lRdDebit > 0
   else
    if ( vRdCredit < 0 ) then
     begin
       vRdDebit  := vRdCredit * (-1);
       vRdCredit := 0;
     end;

end ;


function TZGuide.MoteurCalculGuide ( vTOBLigneEcr,vTOBLigneGuide : TOB ) : boolean ;
var
 lStNumeroPiece     : string;
 lStLibelle         : string;
 lRdDebit           : double;
 lRdCredit          : double;
 lStCritRech        : string;
 lStCritLibelle     : string;
 lStValeurCell      : string;
 lStLibelleGuide    : string;
begin

 result                  := false;

 try

   lStLibelleGuide := vTOBLigneGuide.GetValue('EG_LIBELLE') ;
   lStLibelle      := '';

   // gestion des libelles dans le guide
   lStCritRech     := UpperCase(copy( vTOBLigneGuide.GetValue('EG_LIBELLE'),1,pos(';',vTOBLigneGuide.GetValue('EG_LIBELLE') ) - 1  ));
   lStCritLibelle  := trim(copy( vTOBLigneGuide.GetValue('EG_LIBELLE'),pos(';',vTOBLigneGuide.GetValue('EG_LIBELLE')) + 1 , length(vTOBLigneGuide.GetValue('EG_LIBELLE')) )) ;

   // suppression des caractères de recherche dans la chaine de recherche
   if Copy( lStCritLibelle, 1 , 1) = '*' then

    lStCritLibelle := Copy(lStCritLibelle, 2, length( lStCritLibelle ));

   if Copy( lStCritLibelle , length(lStCritLibelle) , 1) = '*' then
    lStCritLibelle := Copy(lStCritLibelle, 1, length( lStCritLibelle ) - 1);

   if Copy( lStLibelleGuide, 1 , 1) = '*' then
    lStLibelleGuide := Copy(lStLibelleGuide, 2, length( lStLibelleGuide ));

   if Copy( lStLibelleGuide , length(lStLibelleGuide) , 1) = '*' then
     lStLibelleGuide := Copy(lStLibelleGuide, 1, length( lStLibelleGuide ) - 1);


   if lStCritRech <> ''  then
    begin
     if ( lStCritLibelle = '>' ) then
      begin
       if ( length(trim(vTOBLigneEcr.GetValue('E_LIBELLE'))) <= length(lStCritRech) ) then
        lStLibelle := vTOBLigneEcr.GetValue('E_LIBELLE')
        else
         lStLibelle := trim ( copy(
                                   vTOBLigneEcr.GetValue('E_LIBELLE'),
                                   pos( lStCritRech , UpperCase(vTOBLigneEcr.GetValue('E_LIBELLE') )) + length(lStCritRech) ,
                                   length(vTOBLigneEcr.GetValue('E_LIBELLE') )
                             ));
        end
       else
        if lStCritLibelle <> '' then
         begin
          lStValeurCell                       := trim(GFormule(lStCritLibelle, GetFormule, nil, 1));
           if ( lStValeurCell <> '' ) then
              lStLibelle          := lStValeurCell
               else
               lStLibelle         := lStCritRech;
         end; //if
    end // critere de recherche
     else
      if trim(lStLibelleGuide) <> '' then
       begin
        //s := GetFormule(lStLibelleGuide);
        lStValeurCell           := GFormule(lStLibelleGuide, GetFormule, nil, 1);
        if ( lStValeurCell <> '' ) and ( pos(cStIdCodeAFB,lStValeurCell) > 0 )  then
         lStLibelle             := vTOBLigneEcr.GetValue('E_LIBELLE')
          else
           if ( lStValeurCell <> '' ) and ( pos(cStIdCodeAFB,lStValeurCell) = 0 ) then
            lStLibelle          := trim(lStValeurCell)
             else
              lStLibelle        := lStValeurCell; //trim(lStLibelleGuide);
       end; //if

  CalculDebitCredit ( vTOBLigneEcr,vTOBLigneGuide , lRdDebit , lRdCredit ) ;

   if vTOBLigneGuide.GetValue('EG_REFINTERNE')  <> '' then
    lStNumeroPiece                       := GFormule(vTOBLigneGuide.GetValue('EG_REFINTERNE'), GetFormule, nil, 1)
     else
      lStNumeroPiece                     := '';

//   lStARRET := vTOBLigneGuide.GetValue('EG_ARRET') ;

   vTOBLigneEcr.PutValue('_GUIDE'                      , vTOBLigneGuide.GetValue('EG_GUIDE'));
   vTOBLigneEcr.PutValue('_NUMLIGNE'                   , vTOBLigneGuide.GetValue('EG_NUMLIGNE'));
   vTOBLigneEcr.PutValue( FStPrefixe + '_NUMLIGNE'     , FInNumLigne + FInOffset);
   if ( vTOBLigneGuide.GetValue('EG_GENERAL') <> '' ) and (vTOBLigneEcr.GetValue( FStPrefixe + '_GENERAL')= '' )  then
    vTOBLigneEcr.PutValue( FStPrefixe + '_GENERAL'     , vTOBLigneGuide.GetValue('EG_GENERAL') );
   if ( vTOBLigneGuide.GetValue('EG_AUXILIAIRE') <> '' ) then
    vTOBLigneEcr.PutValue( FStPrefixe + '_AUXILIAIRE'  , vTOBLigneGuide.GetValue('EG_AUXILIAIRE') );
   if lStNumeroPiece <> '' then
    vTOBLigneEcr.PutValue( FStPrefixe + '_REFINTERNE'  , lStNumeroPiece);
   if lStLibelle <> '' then
    vTOBLigneEcr.PutValue( FStPrefixe + '_LIBELLE'     , lStLibelle);
   vTOBLigneEcr.PutValue( FStPrefixe + '_DEBIT'        , lRdDebit);
   vTOBLigneEcr.PutValue( FStPrefixe + '_CREDIT'       , lRdCredit);
   vTOBLigneEcr.PutValue( FStPrefixe + '_NATUREPIECE'  , vTOBLigneGuide.GetValue('_NATUREPIECE')) ;
   vTOBLigneEcr.PutValue( FStPrefixe + '_JOURNAL'      , vTOBLigneGuide.GetValue('_JOURNAL')) ;
   result          := true;

  except
   on E:Exception do
    begin
      NotifyError( TraduireMemoire('Erreur de formule dans le Guide n°')          + FStEG_GUIDE
                    + ' ' + RechDom('TTGUIDEECR', FStEG_GUIDE ,false)  + #13#10 + E.message , '' , '' );
    end;
  end;

end;


function  TZGuide.FindFirst : TOB ;
begin
 result := Item.Detail[0] ;
end ;

function  TZGuide.FindNext : TOB ;
begin
 if FInNumLigne < ( Item.Detail.Count ) then
  result := Item.Detail[FInNumLigne]
   else
    result := nil ;
end ;

procedure TZGuide.RecalculGuide ( vTOB : TOB ; ARow : integer ) ;
var
 lCurrentGuide : TOB ;
 lTOBLigneEcr  : TOB ;
begin

  FInNumLigne    := ARow - FInNumDep ;
  FInOffset      := FInNumDep ;
  FTOBResult     := vTOB ;
  lCurrentGuide  := FindNext ;

  while assigned(lCurrentGuide) do
   begin

    lTOBLigneEcr := FTOBResult.Detail[FInNumLigne + FInOffset - 1];

    MoteurCalculGuide( lTOBLigneEcr , lCurrentGuide ) ;

    Inc(FInNumLigne);
    lCurrentGuide  := FindNext ;

  end; // while

end;


function TZGuide.CalculeGuide : boolean;
var
 lTOBLigneEcr  : TOB ;
 lCurrentGuide : TOB ;
begin

  result       := false;

  try

  FInNumLigne  := 0 ;
  // on se place sur la première ligne du guide
  lCurrentGuide  := FindFirst ;

  while assigned(lCurrentGuide) do
  begin

    lTOBLigneEcr  := TOB.Create( 'ECRITURE' , FTOBResult  , -1);
    lTOBLigneEcr.AddChampSupValeur('_GUIDE'     , '000' ) ;
    lTOBLigneEcr.AddChampSupValeur('_NUMLIGNE'  , '0' ) ;

    if not MoteurCalculGuide( lTOBLigneEcr , lCurrentGuide ) then
     begin
      result     := false ;
      exit;
     end;

    Inc(FInNumLigne);
    lCurrentGuide  := FindNext ;

  end; // while

  result := true;

 except
  on E : Exception do
   begin
    MessageAlerte('Erreur lors de l''affectation des guides '+ #10#13 + E.message);
   end;
 end; // if

end;


function TZGuide.RechercheGuideEnBase( vTOB,vTOBResult : TOB ; G : THGrid = nil ) : boolean ;
var
 lTOBLookUp       : TOB ;
 lTOBEnteteGuide  : TOB;
 lStSql           : string ;
 lStCol           : string ;
 lQ               : TQuery ;
 lStLib           : string ;
 lTOB             : TOB ;
begin

 result          := false ;

 if ( vTOBResult = nil )  then exit;

 FInOffset  := 0 ;

 // récuperation des valeurs de la grille
 FListe     := G ;
 if FListe <> nil then
  begin
   FInNumDep := FListe.Row ;
   lTOB      := vTOB.Detail[ FInNumDep - 1 ]
  end
   else
    lTOB := vTOB ;

 FTOBResult  := vTOBResult ;

 lStCol := 'EG_GUIDE;GU_LIBELLE;' ;
 lStLib := 'Code;Libellé;' ;
 lStSql := 'SELECT GU_JOURNAL,GU_LIBELLE,GU_NATUREPIECE,EG_GUIDE,EG_GENERAL FROM ECRGUI, GUIDE WHERE EG_NUMLIGNE=1 AND GU_TYPE="NOR" AND GU_TRESORERIE<>"X" AND GU_GUIDE=EG_GUIDE ' ;

 if lTOB.GetValue('E_JOURNAL') <> '' then
  lStSql := lStSql + ' AND GU_JOURNAL="' + lTOB.GetValue('E_JOURNAL') + '" '
   else
    begin
     lStCol := lStCol + 'GU_JOURNAL;' ;
     lStLib := lStLib + 'Journal;' ;
   end ;  // if

 if ( lTOB.GetValue('E_JOURNAL') <> '' ) and ( lTOB.GetValue('E_NATUREPIECE') <> '' ) then
  lStSql := lStSql + ' AND GU_NATUREPIECE="' + lTOB.GetValue('E_NATUREPIECE') + '" '
   else
    if ( lTOB.GetValue('E_JOURNAL') = '' ) and ( lTOB.GetValue('E_NATUREPIECE') = '' ) then
     begin
      lStCol := lStCol + 'GU_NATUREPIECE;' ;
      lStLib := lStLib + 'Nature;' ;
     end ; // if

 if lTOB.GetValue('E_GENERAL') <> '' then
  lStSql := lStSql + ' AND EG_GENERAL="' + lTOB.GetValue('E_GENERAL') + '" ' ;

 lQ              := OpenSQL(lStSql,true) ; 
 lTOBEnteteGuide := TOB.Create('', nil,-1) ;
 lTOBEnteteGuide.LoadDetailDB('ECRGUI','','',lQ,true) ;
 Ferme(lQ) ;

 if lTOBEnteteGuide.Detail.Count = 0 then
  exit ;

 if lTOBEnteteGuide.Detail.Count = 1 then
   FStEG_GUIDE := lTOBEnteteGuide.Detail[0].GetValue('EG_GUIDE')
    else
     if assigned( FListe ) then
      begin
       lTOBLookUp := LookUpTob ( FListe ,
                                 lTOBEnteteGuide ,
                                 'Guide pour : ' + vTOB.GetValue('E_GENERAL') ,
                                 lStCol ,
                                 lStLib );

       if assigned(lTOBLookUp) then
        FStEG_GUIDE  := lTOBLookUp.GetValue('EG_GUIDE') ;
      end; // if

 lTOBEnteteGuide.Free ;

 if FStEG_GUIDE  <> '' then
  begin
   Load(['NOR',FStEG_GUIDE ]) ;
   result := CalculeGuide ;
  end ;

end ;

end.
