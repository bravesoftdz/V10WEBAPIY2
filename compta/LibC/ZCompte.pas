unit ZCompte;

//=======================================================
//=================== Clés primaires ====================
//=======================================================
// GENERAUX : G_GENERAL

interface

uses
     Classes,
     {$IFNDEF EAGLCLIENT}
     Db,  ComCtrls,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     {$ENDIF}
     SaisComm, // pour le TTYpeExo
     UTob,uEntCommun ;

type
  TZCompte = class
  private
   Compte               : TOB ;
   FStRegimeDefaut      : string;  // code regime par defaut du guide
   FInIndex             : integer ;
   FDossier             : string ;
   FLastCompte          : string ;
   function    GetCount : Integer ;
   function    GetTOB : TOB ;
   function    GetTOBByIndex(Index : integer) : TOB ;
  public
   constructor Create( vDossier : String = '' ) ;
   destructor  Destroy ; override ;
   function    GetCompte(var NumCompte : string ; vBoEnBase : boolean = false) : Integer ;
   procedure   GetCompteParFolio(vStCle : string ) ;
   function    GetValue(Nom : string; Niveau : integer = - 1 ) : Variant ;
   function    GetString(Nom : string; Niveau : integer = - 1 ) : string ;
   procedure   PutValue(Nom : string; Value : variant ; Niveau : integer = - 1) ;
   function    IsCollectif(Niveau : Integer = -1) : Boolean ; overload ; // true si le compte est collectif
   function    IsCollectif( Value : string ) : boolean ; overload ; // idem
   function    IsVentilable(Axe, Niveau : Integer) : Boolean ; overload ;// true si le compte est ventilable sur l'axe <Axe>
   function    IsVentilable(Niveau : Integer = -1 ) : Boolean ; overload ;
   function    IsTiers(Niveau : Integer = -1) : Boolean ;
   function    IsHT(Niveau : Integer = -1) : Boolean ;
   function    IsBqe(Niveau : Integer = -1) : Boolean ; // true pour un compte de banque
   function    IsTvaAutorise( value : string ; vInIndex  : integer = 0 ) : boolean; // true si la saisie de tva est valide su ce compte
   function    IsLettrable( vIndex : integer = -1 ) : boolean; // true si le compte est lettrable
   function    IsPointable( vIndex : integer = -1 ) : boolean; // true si le compte est pointable
   function    IsTICTID(Niveau : Integer = -1 ) : Boolean ;
   function    GetCodeTVAPourUnCompte ( Value : string ) : string; // retourne le code tva du compte
{$IFNDEF NOVH}
   procedure   RecupInfoTVA ( vStCompte : string ; vRdDebit : double ; E_NATUREPIECE,J_NATUREJAL, T_NATUREAUXI : string ; var vStCompteTva : string ; var vRdTauxTva : double; vStRegimeTva : string = '' ; vBoPourTpf : Boolean = False ; vBoPourEnc : Boolean = False ; vBoForcer : boolean = false ) ;
{$ENDIF}
   function    GetRegimeTva( Value : string ) : string; // retourne le regime fiscal du compte
   function    GetTabletteTiers( vIndex : integer = -1 ) : string; // retourne le nom de la tablette des tiers en fonction de la nature des generaux
   function    GetTabletteNature( vIndex : integer = -1 ) : string; // retourne le nom de la tablette des nature en fonction de la nature du compte
   procedure   RAZSolde ;
   procedure   Solde(var D, C: double ; vTypeExo : TTypeExo ) ;
   procedure   Clear ;
  // Chargement sans SQL
  function    GetCompteOpti(var NumCompte : string ; vTobSource : tob ) : Integer ;

  property    Count     : Integer read GetCount ;
  property    InIndex   : integer read FInIndex ;
  property    Item      : TOB     read GetTOB ;
  property    Items     : TOB     read Compte ;

  end ;

implementation

uses
  ParamSoc, // pour le GetParamSoc
  SysUtils ,
  HCtrls,  // pour les ferme OpenSQL
  UtilSais ,
  HEnt1 ,
  HMsgBox , // pour les PGIError
  UtilPGI
  , Ent1 
  {$IFDEF MODENT1}
  , CPProcMetier // pour les fct de bourrage
  , CPTypeCons
  {$ENDIF MODENT1}
   ; // EstMultiSoc


const ROWS = 'G_GENERAL, G_LIBELLE, G_TOTALDEBIT, G_TOTALCREDIT, G_NATUREGENE, '
           + 'G_COLLECTIF, G_SENS, G_REGIMETVA, G_TVA, G_TPF, G_LETTRABLE' ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 14/02/2002
Modifié le ... :   /  /
Description .. : Rajout de FStRegimeDefaut
Mots clefs ... :
*****************************************************************}
constructor TZCompte.Create( vDossier : string = '' ) ;
begin
Compte:=TOB.Create('VCOMPTE', nil, -1) ; FStRegimeDefaut:='';
FInIndex:=-1 ;
if vDossier <> '' then
 FDossier := vDossier
  else
   FDossier := V_PGI.SchemaName ;
FLastCompte := '' ;
end ;

destructor TZCompte.Destroy ;
begin
if Compte<>nil then Compte.Free ;
end ;

function TZCompte.GetCount : Integer ;
begin
Result:=Compte.Detail.Count ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/11/2002
Modifié le ... :   /  /
Description .. : - suppression de warning
Mots clefs ... :
*****************************************************************}
function TZCompte.GetCompte(var NumCompte : string ; vBoEnBase : boolean = false ) : Integer ;
var Q : TQuery ; i : integer ;
    lTob : TOB ;
    lBoLoad : Boolean ;
begin
Result:=-1 ; Q:=nil ;
//debug('+++ GetCompte Debut ==> FLastCompte = ' + FLastCompte + '(' + IntToStr(FinIndex) + ')' + '/// NumCompte = ' + NumCompte ) ;
if not vBoEnBase and ( FLastCompte = NumCompte ) then
 begin
  result := FinIndex ;
  exit ;
 end ;
try
  FinIndex:=-1 ;
//  FLastCompte := NumCompte ;
//  debug('+++ GetCompte Après affectation ==> FLastCompte = ' + FLastCompte + '(' + IntToStr(FinIndex) + ')' + '/// NumCompte = ' + NumCompte ) ;
  NumCompte:=Trim(NumCompte) ; if NumCompte='' then Exit ;
  NumCompte:=BourreLaDonc(NumCompte, fbGene) ;

   for i:=0 to Compte.Detail.Count-1 do
    if Compte.Detail[i].GetString('G_GENERAL')=UpperCase(NumCompte) then
        if Compte.Detail[i].GetString('_LOADDB')='X' then
          begin
          Result:=i ;
          FInIndex:=i ;
          end
           else
             begin
             Result:=-1 ;
             FInIndex:=-1 ;
             if not vBoEnBase then Exit
             end ;


//    debug('+++ GetCompte Après Rech Mémoire ==> FLastCompte = ' + FLastCompte + '(' + IntToStr(FinIndex) + ')' + '/// NumCompte = ' + NumCompte ) ;
    if not vBoEnBase and (result <> -1) then Exit ;

  try

  Q := OpenSelect('SELECT * FROM GENERAUX WHERE G_GENERAL="'+UpperCase(NumCompte)+'" ORDER BY G_GENERAL', FDossier) ;
  lBoLoad := not Q.Eof ;

  if lBoLoad then
   begin

     if result<>-1 then
      lTob   := Compte.Detail[ result ]
       else
        begin
        lTob   := Tob.Create('GENERAUX', Compte, -1) ;
        result := Compte.Detail.count-1 ;
        end ;
     lTob.SelectDB('', Q); //   Compte.LoadDetailDB('GENERAUX', '', '', Q, True) ;
     lTob.AddChampSupValeur('_SOLDED',-1,false) ;
     lTob.AddChampSupValeur('_SOLDEC',-1,false) ;
     lTob.AddChampSupValeur('_LOADDB','X',false) ;
     // gestion table locale en multisoc
     CChargeCumulsMS( fbGene, UpperCase(NumCompte), '', lTob, FDossier ) ;
   end
    else
      begin
      lTob := TOB.Create('GENERAUX', Compte, -1 );
      lTob.PutValue('G_GENERAL', UpperCase(NumCompte) );
      lTOb.AddChampSupValeur('_LOADDB','-',false) ;
      end ;
  FInIndex:=result ;
  finally
//    debug('+++ GetCompte Après Rech Base ==> FLastCompte = ' + FLastCompte + '(' + IntToStr(FinIndex) + ')' + '/// NumCompte = ' + NumCompte ) ;

   if assigned(Q) then Ferme(Q) ;
  end;
finally
  FLastCompte := NumCompte ;
//  debug('+++ GetCompte FIN ==> FLastCompte = ' + FLastCompte + '(' + IntToStr(FinIndex) + ')' + '/// NumCompte = ' + NumCompte ) ;
end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 06/11/2006
Modifié le ... :   /  /    
Description .. : - LG - 06/11/2006 - FB 16744 - on vide la liste avant de 
Suite ........ : la remplir
Mots clefs ... : 
*****************************************************************}
procedure TZCompte.GetCompteParFolio(vStCle : string ) ;
//var lQ  : TQuery ;
//    lTOB : TOB ;
begin
Clear ;
{
lQ := OpenSql('SELECT * FROM GENERAUX WHERE G_GENERAL IN  ( SELECT E_GENERAL FROM ECRITURE WHERE ' + vStCle + ' ) ', true ) ;
while not lQ.EOF do
 begin
  lTOB := TOB.Create('GENERAUX',Compte,-1 ) ;
  lTOB.SelectDB('',lQ,true) ;
  lQ.Next ;
 end ; // while
Ferme(lQ) ;
}
Compte.LoadDetailDBFromSQL( 'GENERAUX',
                            'SELECT * FROM GENERAUX WHERE G_GENERAL IN  ( SELECT E_GENERAL FROM ECRITURE WHERE ' + vStCle + ' ) ' ) ;
if Compte.Detail.Count >0 then
 begin
  Compte.Detail[0].AddChampSupValeur('_SOLDED',-1,true) ;
  Compte.Detail[0].AddChampSupValeur('_SOLDEC',-1,true) ;
  Compte.Detail[0].AddChampSupValeur('_LOADDB','X',true) ;
 end ;
end ;

function TZCompte.GetValue(Nom : string; Niveau : integer = - 1) : Variant ;
begin
result:=#0 ;
if Niveau<>-1 then Result:=Compte.Detail[Niveau].GetValue(Nom)
 else
  if FInIndex<>-1 then Result:=Compte.Detail[FInIndex].GetValue(Nom) ;
end ;

function TZCompte.GetString(Nom : string; Niveau : integer = - 1) : string ;
begin
result:= GetValue(Nom,Niveau) ;
if result = #0 then
 result := '' ;
end ;


procedure TZCompte.PutValue(Nom : string; Value : variant ; Niveau : integer = - 1) ;
begin
if Niveau<>-1 then Compte.Detail[Niveau].PutValue(Nom,Value)
 else
  if FInIndex<>-1 then Compte.Detail[FInIndex].PutValue(Nom,Value) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 14/02/2002
Modifié le ... :   /  /    
Description .. : Suppression warning
Mots clefs ... : 
*****************************************************************}
function TZCompte.IsVentilable(Axe, Niveau : Integer) : Boolean ;
var j : Integer ; zz : Boolean ;
begin
if Axe>0 then Result:=Compte.Detail[Niveau].GetValue('G_VENTILABLE'+IntToStr(Axe))='X'
         else begin
         Result:=FALSE ;
         for j:=1 to MAXAXE do
             begin
             zz:=Compte.Detail[Niveau].GetValue('G_VENTILABLE'+IntToStr(j))='X' ;
             Result:=(Result) or (zz) ;
             end ;
         end ;
end ;

function TZCompte.IsVentilable(Niveau : Integer = -1 ) : Boolean ;
begin
if Niveau<>-1 then
 Result:= Compte.Detail[Niveau].GetValue('G_VENTILABLE')='X'
  else
   if FInIndex<>-1 then
    Result:= Compte.Detail[FInIndex].GetValue('G_VENTILABLE')='X'
     else
      begin
       result := false ;
       PGIError('Compte non chargé en mémoire','Attention') ;
      end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 11/05/2004
Modifié le ... :   /  /
Description .. : - LG - 11/05/2004 - FB 13539 - Test sur FInIndex <> - 1
Mots clefs ... : 
*****************************************************************}
function TZCompte.IsCollectif(Niveau : Integer = -1 ) : Boolean ;
begin
result := false ;
if Niveau<>-1 then
 Result:= Compte.Detail[Niveau].GetValue('G_COLLECTIF')='X'
  else
   if FInIndex<>-1 then
    Result:= Compte.Detail[FInIndex].GetValue('G_COLLECTIF')='X';
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 14/02/2002
Modifié le ... :   /  /
Description .. : reecriture du code
Mots clefs ... :
*****************************************************************}
function TZCompte.IsTiers(Niveau : Integer = - 1) : Boolean ;
var Nat : string ;
begin
if Niveau<>-1 then Nat:=Compte.Detail[Niveau].GetValue('G_NATUREGENE')
else Nat:=Compte.Detail[FInIndex].GetValue('G_NATUREGENE') ;
result:= ((Nat='TID') or (Nat='TIC'));
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 14/02/2002
Modifié le ... :   /  /
Description .. : reecriture du code
Mots clefs ... :
*****************************************************************}
function TZCompte.IsHT(Niveau : Integer = -1 ) : Boolean ;
var Nat : string ;
begin
if Niveau<>-1 then Nat:=Compte.Detail[Niveau].GetValue('G_NATUREGENE')
else Nat:=Compte.Detail[FInIndex].GetValue('G_NATUREGENE') ;
result:= ((Nat='CHA') or (Nat='PRO') or (Nat='IMO'));
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 14/02/2002
Modifié le ... :   /  /
Description .. : reecriture du code
Mots clefs ... :
*****************************************************************}
function TZCompte.IsBqe(Niveau : Integer) : Boolean ;
var Nat : string ;
begin
Nat:=Compte.Detail[Niveau].GetValue('G_NATUREGENE') ;
result:=((Nat='BQE') or (Nat='CAI'));
end ;

//LG* 04/12/2001
{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 14/02/2002
Modifié le ... : 16/05/2005
Description .. : retourne true si le code si on peut saisir un code tva
Suite ........ : - LG - 04/01/2004 - FB 13678 - les TIC ou TID gerent aussi 
Suite ........ : la tva
Suite ........ : - LG - 16/05/2005 - FB 13678 - ajout des comptes de 
Suite ........ : nature divers
Mots clefs ... : 
*****************************************************************}
function TZCompte.IsTvaAutorise( value : string ; vInIndex  : integer = 0 ) : boolean;
var
 lIndex : integer;
begin
 result := true; if value = '' then exit; result := false;
 if vInIndex = 0 then lIndex := GetCompte(Value)
 else lIndex := vInIndex;

 if ( lIndex > - 1 ) then
 result := ( GetValue('G_NATUREGENE', lIndex) = 'CHA' ) or  // Charges
           ( GetValue('G_NATUREGENE', lIndex) = 'IMO' ) or  // Immobilisation
           ( GetValue('G_NATUREGENE', lIndex) = 'TIC' ) or
           ( GetValue('G_NATUREGENE', lIndex) = 'TID' ) or
           ( GetValue('G_NATUREGENE', lIndex) = 'COD' ) or
           ( GetValue('G_NATUREGENE', lIndex) = 'PRO' );    // Produit
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 14/02/2002
Modifié le ... :   /  /
Description .. : retourne true si le compte est collectif
Mots clefs ... :
*****************************************************************}
function TZCompte.IsCollectif( Value : string ) : boolean;
var
 lIndex : integer;
begin
 result := true; if Value =  '' then exit; lIndex := GetCompte(Value);
 if ( lIndex > -1 ) then result := IsCollectif(lIndex)
 else result := false;
end;

function TZCompte.IsTICTID(Niveau : Integer = -1 ) : Boolean ;
begin
result := false ;
if Niveau<>-1 then
 Result:= (Compte.Detail[Niveau].GetValue('G_NATUREGENE')='TID') or (Compte.Detail[Niveau].GetValue('G_NATUREGENE')='TIC') 
  else
   if FInIndex<>-1 then
    Result:= (Compte.Detail[FInIndex].GetValue('G_NATUREGENE')='TID') or (Compte.Detail[FInIndex].GetValue('G_NATUREGENE')='TIC') ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 14/02/2002
Modifié le ... : 20/03/2003
Description .. : retourne le code tva d'un compte ou le parametre genral si
Suite ........ : inexistant
Suite ........ : 
Suite ........ : - LG - 20/03/2003 - le test sur lIndex etait faux
Mots clefs ... : 
*****************************************************************}
function TZCompte.GetCodeTVAPourUnCompte ( Value : string ) : string;
var
 lIndex     : integer;
begin
 result:=''; lIndex:=GetCompte(Value);
 if ( lIndex > -1 ) then result:=GetValue('G_TVA', lIndex) ;
 if result = '' then result:=GetParamSoc('SO_CODETVAGENEDEFAULT');
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 23/06/2004
Modifié le ... : 21/12/2004
Description .. : - 23/06/2004 - FB 13669 - LG  -  la détermination du
Suite ........ : compte et du taux de TVA ne fonctionne pas si la nature du 
Suite ........ : journal et la nature de pièce choisie sont "OD".
Suite ........ :
Suite ........ : - 21/12/2004 - SBO - Ajout d'un paramètre facultatif
Suite ........ : permettant de choisir si l'on récupère les infos pour la TVA
Suite ........ : ou la TPF. (par défaut : pour la TVA )
Suite ........ :
Suite ........ : - 04/09/2006 - SBO - FQ 18525 :
Suite ........ : Ajout d'un paramètre facultatif pour gestion de
Suite ........ :  la tva/encaissement avec Accélérateur de saisie
Mots clefs ... :
*****************************************************************}
{$IFNDEF NOVH}
procedure TZCompte.RecupInfoTVA ( vStCompte : string ; vRdDebit : double ; E_NATUREPIECE,J_NATUREJAL, T_NATUREAUXI : string ; var vStCompteTva : string ; var vRdTauxTva : double; vStRegimeTva : string = '' ; vBoPourTpf : Boolean = False; vBoPourEnc : Boolean = false ; vBoForcer : boolean = false) ;
var
 lTOB        : TOB;
 lStCodeTVA  : string;
 lStRegimeTva : string;
 lStCatTVA    : string;
begin
 vRdTauxTva:=0 ; vStCompteTva:='' ;

 if T_NATUREAUXI = 'SAL' then exit ;

 if vStCompte = '' then
  begin
   MessageAlerte('Le compte doit être définie !' + #10#13 + 'Fonction RecupInfoTVA') ;
   exit ;
  end;

 if vStRegimeTva = '' then
  lStRegimeTva := GetParamSoc('SO_REGIMEDEFAUT')
   else
    lStRegimeTva := vStRegimeTva;

 if IsTvaAutorise(vStCompte) or vBoForcer then
  begin
   lStCodeTVA:=GetCodeTVAPourUnCompte(vStCompte);
   // Détermination du code de catégorie TVA ou TPF
   if vBoPourTpf
     then lStCatTVA := VH^.DefCatTPF
     else lStCatTVA := VH^.DefCatTVA ;
   lTOB:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_CODETAUX','TV_REGIME'],[lStCatTVA,lStCodeTVA,lStRegimeTva],true);
   if assigned(lTOB) then
    begin
     if J_NATUREJAL = 'ACH' then
      begin
        if vBoPourEnc
          then vStCompteTva:=lTOB.GetValue('TV_ENCAISACH')
          else vStCompteTva:=lTOB.GetValue('TV_CPTEACH') ;
       vRdTauxTva:=lTOB.GetValue('TV_TAUXACH') ;
      end
       else
        if J_NATUREJAL = 'VTE' then
         begin
          vRdTauxTva:=lTOB.GetValue('TV_TAUXVTE') ;
          if vBoPourEnc 
            then vStCompteTva:=lTOB.GetValue('TV_ENCAISVTE')
            else vStCompteTva:=lTOB.GetValue('TV_CPTEVTE') ;
         end
          else
           if ( J_NATUREJAL = 'OD' ) or ( J_NATUREJAL = 'REG' ) then
            begin
               case CaseNatP(E_NATUREPIECE) of
                1,2,3 : begin
                        vRdTauxTva:=lTOB.GetValue('TV_TAUXVTE') ;
                        if vBoPourEnc
                          then vStCompteTva:=lTOB.GetValue('TV_ENCAISVTE')
                          else vStCompteTva:=lTOB.GetValue('TV_CPTEVTE') ;
                        end ;
                4,5,6 : begin
                        vRdTauxTva:=lTOB.GetValue('TV_TAUXACH') ;
                        if vBoPourEnc 
                          then vStCompteTva:=lTOB.GetValue('TV_ENCAISACH')
                          else vStCompteTva:=lTOB.GetValue('TV_CPTEACH') ;
                        end ;
                7     : begin // OD
                         Case CASENATA(T_NATUREAUXI) of
                          1 : begin
                              vRdTauxTva:=lTOB.GetValue('TV_TAUXVTE') ;
                              if vBoPourEnc 
                                then vStCompteTva:=lTOB.GetValue('TV_ENCAISVTE')
                                else vStCompteTva:=lTOB.GetValue('TV_CPTEVTE') ;
                              end ;
                          2 : begin
                              vRdTauxTva:=lTOB.GetValue('TV_TAUXACH') ;
                              if vBoPourEnc 
                                then vStCompteTva:=lTOB.GetValue('TV_ENCAISACH')
                                else vStCompteTva:=lTOB.GetValue('TV_CPTEACH') ;
                              end ;
                          else if ( vRdDebit > 0 ) then
                                 begin
                                 vRdTauxTva:=lTOB.GetValue('TV_TAUXACH') ;
                                 if vBoPourEnc 
                                   then vStCompteTva:=lTOB.GetValue('TV_ENCAISACH')
                                   else vStCompteTva:=lTOB.GetValue('TV_CPTEACH') ;
                                 end
                                else begin
                                     vRdTauxTva:=lTOB.GetValue('TV_TAUXVTE') ;
                                     if vBoPourEnc 
                                       then vStCompteTva:=lTOB.GetValue('TV_ENCAISVTE')
                                       else vStCompteTva:=lTOB.GetValue('TV_CPTEVTE') ;
                                     end
                         end ; // case
                        end ; // OD
               end ; // case
            end  // if
             else // pour tout les autres cas
              if ( vRdDebit > 0 ) then
                begin
                vRdTauxTva:=lTOB.GetValue('TV_TAUXACH') ;
                if vBoPourEnc 
                  then vStCompteTva:=lTOB.GetValue('TV_ENCAISACH')
                  else vStCompteTva:=lTOB.GetValue('TV_CPTEACH') ;
                end
              else
                begin
                vRdTauxTva:=lTOB.GetValue('TV_TAUXVTE') ;
                if vBoPourEnc 
                  then vStCompteTva:=lTOB.GetValue('TV_ENCAISVTE')
                  else vStCompteTva:=lTOB.GetValue('TV_CPTEVTE') ;
                end
    end // if assigned
  end ;
end ;
{$ENDIF}


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 14/02/2002
Modifié le ... :   /  /
Description .. : retourne le regime fiscal d'un compte ou le parametre
Suite ........ : general
Mots clefs ... :
*****************************************************************}
function TZCompte.GetRegimeTva(Value : string) : string;
var
 lIndex     : integer;
begin
 result:='' ; lIndex:=GetCompte(Value) ;
 if (lIndex>0) then result:=GetValue('G_REGIMETVA', lIndex) ;
 if result='' then result:=FStRegimeDefaut ;
 if result='' then BEGIN FStRegimeDefaut:=GetParamSoc('SO_REGIMEDEFAUT') ; result:=FStRegimeDefaut ; END;
end;


function TZCompte.IsLettrable( vIndex : integer ) : boolean;
begin
 result:=GetValue('G_LETTRABLE',vIndex) = 'X' ;
end;

function TZCompte.IsPointable( vIndex : integer ) : boolean;
begin
 result:=GetValue('G_POINTABLE',vIndex) = 'X' ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 14/02/2002
Modifié le ... :   /  /
Description .. : retourne le nom de la tablette des tiers en fonction de la
Suite ........ : nature du compte ( on ne verifie pas que le code est
Suite ........ : collectif ! )
Mots clefs ... :
*****************************************************************}
function TZCompte.GetTabletteTiers( vIndex : integer ) : string;
var
 lStNatureGene : string;
begin
 lStNatureGene:=GetValue('G_NATUREGENE',vIndex);
 if lStNatureGene='COC' then result:='TZTCLIENT'
 else if lStNatureGene='COF' then result:='TZTFOURN'
 else if lStNatureGene='COD' then result:='TZTTOUS'
 else if lStNatureGene='COS' then result:='TZTSALARIE'
 else result:='TZTTOUS';
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 20/02/2002
Modifié le ... :   /  /
Description .. : Retourne le nom de la tablette des natures de piece
Suite ........ : autorise pour ce compte
Mots clefs ... :
*****************************************************************}
function TZCompte.GetTabletteNature( vIndex : integer ) : string;
var
 lStNatureGene : string;
begin
 lStNatureGene:=GetValue('G_NATUREGENE',vIndex);
 if lStNatureGene='COC' then result:='ttNatPieceVente'
 else if lStNatureGene='COF' then result:='ttNatPieceAchat'
 else if lStNatureGene='COD' then result:='ttNaturePiece'
 else result:='ttNaturePiece';
end;


procedure TZCompte.RAZSolde ;
var
 i : integer ;
begin
  for i:=0 to Compte.Detail.Count-1 do
   begin
    Compte.Detail[i].PutValue('_SOLDED',-1) ;
    Compte.Detail[i].PutValue('_SOLDEC',-1) ;
   end ;
end ;

procedure TZCompte.Solde(var D, C: double ; vTypeExo : TTypeExo );
var
 lStLettre : string ;
begin

 case vTypeExo of
  teEncours   : lStLettre := 'E' ;
  teSuivant   : lStLettre := 'S' ;
  tePrecedent : lStLettre := 'P' ;
 end ;// case

 D := GetValue('G_TOTDEB' + lStLettre) ;
 C := GetValue('G_TOTCRE' + lStLettre) ;

end;

function TZCompte.GetTOB: TOB;
begin
 result := GetTOBByIndex( FInIndex ) ;
end;

function TZCompte.GetTOBByIndex(Index: integer): TOB;
begin
 result:=nil ;
 if ( Index <> -1 ) and ( Index < GetCount ) then
   begin
   result   := Compte.Detail[ Index ] ;
   FInIndex := Index ;
   end ;
end;


function TZCompte.GetCompteOpti(var NumCompte: string ; vTobSource : tob ): Integer;
var i       : integer ;
    lTob    : TOB ;
    lStChp  : string ;
begin
  Result:=-1 ;

  if FLastCompte = NumCompte then
    begin
    result := FinIndex ;
    exit ;
    end ;

  FinIndex:=-1 ;
  FLastCompte := NumCompte ;

  NumCompte:=Trim(NumCompte) ;
  if NumCompte='' then Exit ;

  NumCompte:=BourreLaDonc(NumCompte, fbGene) ;
  for i:=0 to Compte.Detail.Count-1 do
    if Compte.Detail[i].GetValue('G_GENERAL')=UpperCase(NumCompte) then
      begin
      Result:=i ;
      FInIndex:=result ;
      Exit ;
      end ;

  if Assigned( vTobSource ) and ( vTobSource.GetString('G_GENERAL') = NumCompte ) then
    begin

    // Création objet
    lTOB := TOB.Create( 'GENERAUX', Compte, -1 ) ;
    lTOB.AddChampSupValeur('_SOLDED',-1,false) ;
    lTOB.AddChampSupValeur('_SOLDEC',-1,false) ;
    lTOB.AddChampSupValeur('_LOADDB','X',false) ;

    // Recopie des infos
    for i:=1 to lTob.NbChamps do
      begin
      lStChp := lTob.GetNomChamp( i ) ;
      if vTobSource.GetNumChamp( lstChp ) > 0 then
        lTOB.PutValue( lStChp, vTobSource.GetValue( lStChp ) ) ;
      end ;

    // indicateur
    FInIndex := Compte.Detail.Count-1 ;
    result   := FInIndex ;

    end
  else result := GetCompte( NumCompte ) ;

end;


procedure TZCompte.Clear ;
begin
 Compte.ClearDetail ;
 FLastCompte := '' ;
 FInIndex    := -1 ;
end;

end.


