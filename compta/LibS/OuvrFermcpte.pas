unit OuvrFermcpte;

interface
uses
  Windows,
  SysUtils,  // Date, FreeAndNil
  Classes,
  Controls,  // mrYes
  Forms,     // TForm
  hmsgbox,   // PGIAsk, PGIError, MessageAlerte, HShowMessage
  HCtrls,    // ExecuteSQL, UsDateTime
{$IFDEF EAGLCLIENT}
  eMul,      // TFMul
{$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  HDB,       // THDBGrid, NbSelected, AllSelected
  Mul,       // TFMul
{$ENDIF}
  Ent1,      // TFichierBase, fbSect, fbBudSec1, fbBudSec5, fbGene, fbAux, fbJal, fbBudJal
  HStatus,   // MoveCur
  M3FP,      // RegisterAglProc
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  uTob;      // Tob

procedure CPOuvreFermeListeCpte( vTobFListe: Tob ; vType : TFichierBase ; vFermeture : Boolean );
procedure AGLOuvrFermListCpte(parms: array of variant; nb: integer );

type
TFouv = class
    StCode  : string;
    LaTable : string;
    Stc     : string;
    Pref    : string;
    Sta     : string;
    Lefb  : TFichierBase ;
    private
    Fermer: Boolean ;
    function  EstDansVentilType : Boolean ;
    function  EstDansGuideAbo : Boolean ;
    function  FabriqMess        : String ;
    procedure OuvrirOuFermer;
    procedure CPModifieComptes( vTobFListe : Tob );
{$IFDEF EAGLCLIENT}
    procedure ModifieComptes(L : THGrid; Q : TQuery; Titre : String);
{$ELSE}
    procedure ModifieComptes(L : THDBGrid; Q : TQuery; Titre : String);
{$ENDIF}
end;

implementation


uses
  {$IFDEF MODENT1}
  CPProcGen,
  {$ENDIF MODENT1}
  HEnt1;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 07/10/2003
Modifié le ... : 12/12/2003
Description .. : - LG - 07/10/2003 - on recharge les tablettes suite a la mise
Suite ........ : a jour des infos
Suite ........ : - CA - 12/12/2003 - On met à jour les tablettes uniquement
Suite ........ : à la fin du traitement
Mots clefs ... : 
*****************************************************************}
Procedure TFouv.OuvrirOuFermer;
BEGIN
if Fermer then
   BEGIN

   if lefb=fbSect then
      BEGIN
      if ExecuteSql('UPDATE '+LaTable+' SET '+Pref+'FERME="-", '+Pref+'DATEOUVERTURE="'+UsDateTime(Date)+'", '+
                    ''+Pref+'DATEMODIF="'+UsDateTime(Date)+'" Where '+StCode+'="'+Stc+'" and S_AXE="' + StA +'"')<>1 then V_PGI.IoError:=oeUnknown ;
      END else
      if lefb in [fbBudSec1..fbBudSec5] then
         BEGIN
//         StA := Q.FindField('BS_AXE').AsString ;
         if ExecuteSql('UPDATE '+LaTable+' SET '+Pref+'FERME="-", '+Pref+'DATEOUVERTURE="'+UsDateTime(Date)+'", '+
                       ''+Pref+'DATEMODIF="'+UsDateTime(Date)+'" Where '+StCode+'="'+Stc+'" and BS_AXE="' + StA +'"')<>1 then V_PGI.IoError:=oeUnknown ;
         END else
         BEGIN
         if ExecuteSql('UPDATE '+LaTable+' SET '+Pref+'FERME="-", '+Pref+'DATEOUVERTURE="'+UsDateTime(Date)+'", '+
                       ''+Pref+'DATEMODIF="'+UsDateTime(Date)+'" Where '+StCode+'="'+Stc+'"')<>1 then
            V_PGI.IoError:=oeUnknown ;
         END ;
   END else
   BEGIN
   if lefb=fbSect then
      BEGIN
//     StA := Q.FindField('S_AXE').AsString ;
      if ExecuteSql('UPDATE '+LaTable+' SET '+Pref+'FERME="X", '+Pref+'DATEFERMETURE="'+UsDateTime(Date)+'", '+
                    ''+Pref+'DATEMODIF="'+UsDateTime(Date)+'" Where '+StCode+'="'+Stc+'" and S_AXE="' + StA +'"')<>1 then V_PGI.IoError:=oeUnknown ;
      END else
      if lefb in [fbBudSec1..fbBudSec5] then
         BEGIN
//         StA := Q.FindField('BS_AXE').AsString ;
         if ExecuteSql('UPDATE '+LaTable+' SET '+Pref+'FERME="X", '+Pref+'DATEFERMETURE="'+UsDateTime(Date)+'", '+
                       ''+Pref+'DATEMODIF="'+UsDateTime(Date)+'" Where '+StCode+'="'+Stc+'" and BS_AXE="' + StA +'"')<>1 then V_PGI.IoError:=oeUnknown ;
         END else
         BEGIN
         if ExecuteSql('UPDATE '+LaTable+' SET '+Pref+'FERME="X", '+Pref+'DATEFERMETURE="'+UsDateTime(Date)+'", '+
                       ''+Pref+'DATEMODIF="'+UsDateTime(Date)+'" Where '+StCode+'="'+Stc+'"')<>1 then
                       V_PGI.IoError:=oeUnknown ;
         END ;
   END ;
//  AvertirMultiTable('TTJOURNAL') ;
END ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/11/2002
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
procedure CPOuvreFermeListeCpte( vTobFListe: Tob ; vType : TFichierBase ; vFermeture : Boolean );
var
  Evtouv  : TFouv;
begin
  Evtouv := TFouv.Create;

  if not vFermeture then
    Evtouv.Fermer := FALSE
  else
    Evtouv.Fermer := TRUE;

  Case vType of
    fbGene : begin
               Evtouv.Pref    := 'G_' ;
               Evtouv.StCode  := 'G_GENERAL';
               Evtouv.LaTable := 'GENERAUX';
               EvtOuv.Lefb    := fbGene;
               EvtOuv.CPModifieComptes( vTobFListe );
             end;

    fbAux :  begin
               Evtouv.Pref    := 'T_' ;
               Evtouv.StCode  := 'T_AUXILIAIRE';
               Evtouv.LaTable := 'TIERS';
               EvtOuv.Lefb    := fbAux;
               EvtOuv.CPModifieComptes( vTobFListe );
             end;

    fbSect : begin
               Evtouv.Pref   := 'S_' ;
               Evtouv.StCode := 'S_SECTION';
               Evtouv.LaTable:='SECTION';
               EvtOuv.Lefb := fbSect;
               EvtOuv.CPModifieComptes( vTobFListe );
             end;

    fbJal :  begin
               Evtouv.Pref    := 'J_' ;
               Evtouv.StCode  := 'J_JOURNAL';
               Evtouv.LaTable := 'JOURNAL';
               EvtOuv.Lefb    := fbJal;
               EvtOuv.CPModifieComptes( vTobFListe );
      END;

  else Exit ;

 end; // End du Case

 FreeAndNil( EvtOuv );

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/11/2002
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFouv.CPModifieComptes( vTobFListe : Tob );
var i : integer;
    lMsg : string;
begin

  if Fermer then
    lMsg := TraduireMemoire('Désirez-vous ré-ouvrir les comptes sélectionnés ?')
  else
    lMsg := TraduireMemoire('Désirez-vous fermer les comptes sélectionnés ?');

  if (Lefb = fbSect) or (Lefb in [fbBudsec1..fbBudsec5]) then
  begin
    if Fermer then
      lMsg := TraduireMemoire('Désirez-vous ré-ouvrir les sections sélectionnées ?')
    else
      lMsg := TraduireMemoire('Désirez-vous fermer les sections sélectionnées ?');
  end;

  if (Lefb = fbJal) then //or (Lefb=fbBudjal) then
  begin
    if Fermer then
      lMsg := TraduireMemoire('Désirez-vous ré-ouvrir les journaux sélectionnés ?')
    else
      lMsg := TraduireMemoire('Désirez-vous fermer les journaux sélectionnés ?');
  end;

  if (Lefb = fbAux) then
  begin
    if Fermer then
      lMsg := TraduireMemoire('Désirez-vous activer les comptes sélectionnés ?')
    else
      lMsg := TraduireMemoire('Désirez-vous mettre en sommeil les comptes sélectionnés ?');
  end;

  if (Lefb = fbBudJal) then
  begin
    if Fermer then
      lMsg := TraduireMemoire('Désirez-vous ré-ouvrir les budgets sélectionnés ?')
    else
      lMsg := TraduireMemoire('Désirez-vous fermer les budgets sélectionnés ?');
  end;

  if PGIAsk( lMsg , 'Attention') <> mrYes then Exit;

  for i := 0 to vTobFListe.Detail.Count -1 do
  begin
    if vTobFListe.Detail[i].GetValue('SELECT') = 'X' then
    begin
      Stc := vTobFListe.Detail[i].GetValue(StCode);
      if (Lefb = fbSect) then
        StA := vTobFListe.Detail[i].GetValue('S_AXE');

      // Vérification avant fermeture : FQ 16037 SBO 22/09/2005
      if (not Fermer) and EstDansVentilType then
        begin
        PgiBox( FabriqMess ) ;
        Break ;
        end ;

      // Vérification avant fermeture : FQ 22070  19.12.2007  YMO
      if (not Fermer) and EstDansGuideAbo then
        begin
        PgiBox( 'Le compte ne peut être fermé. Il est utilisé dans un abonnement ou dans un guide' ) ;
        Break ;
        end ;

      if Transactions(OuvrirOuFermer,5)<> oeOK then
        begin
        PgiError('Ouverture/Fermeture impossible', '' ) ;
        Break ;
        end ;

    end;
  end;

  // MAJ des tablettes - FQ 16036 SBO 21/09/2005
  if Lefb = fbJal then
    begin
    AvertirTable('TTJALSAISIE');
    AvertirTable('TTJALANOUVEAU');
    AvertirTable('TTJALBANQUE');
    AvertirTable('TTJOURNAUX');
    AvertirTable('CPJOURNALIFRS');
    AvertirTable('TTJALSANSECART');
    AvertirTable('TTJALSAISIE');
    AvertirTable('TTJOURNAL');
    end ;

end;

////////////////////////////////////////////////////////////////////////////////
{$IFDEF EAGLCLIENT}
Procedure TFouv.ModifieComptes(L : THGrid; Q :TQuery; Titre : String);
{$ELSE}
Procedure TFouv.ModifieComptes(L : THDBGrid; Q :TQuery; Titre : String);
{$ENDIF}
Var i : Byte ;
    Msg   : string;
BEGIN
  if (L.NbSelected=0) and (not L.AllSelected) then
  begin
    MessageAlerte('Aucun élément sélectionné.');
    Exit;
  end;

if Fermer then Msg:='1;'+Titre+';Désirez-vous ré-ouvrir les comptes sélectionnés ?;Q;YN;N;'
          else Msg:='2;'+Titre+';Désirez-vous fermer les comptes sélectionnés ?;Q;YN;N;';

if (Lefb=fbSect) or (Lefb in [fbBudsec1..fbBudsec5]) then BEGIN
   if Fermer then Msg:='14;'+Titre+';Désirez-vous ré-ouvrir les sections sélectionnées ?;Q;YN;N;'
             else Msg:='15;'+Titre+';Désirez-vous fermer les sections sélectionnées ?;Q;YN;N;' ; END ;

if (Lefb=fbJal) or (Lefb=fbBudjal) then BEGIN
  if Fermer then Msg:='16;'+Titre+';Désirez-vous ré-ouvrir les journaux sélectionnés ?;Q;YN;N;'
            else Msg:='17;'+Titre+';Désirez-vous fermer les journaux sélectionnés ?;Q;YN;N;' ; END ;

if Lefb=fbAux then BEGIN
  if Fermer then Msg:='27;'+Titre+';Désirez-vous activer les comptes sélectionnés ?;Q;YN;N;'
            else Msg:='28;'+Titre+';Désirez-vous mettre en sommeil les comptes sélectionnés ?;Q;YN;N;'; END ;

if Lefb=fbBudJal then BEGIN
  if Fermer then Msg:='50;'+Titre+';Désirez-vous ré-ouvrir les budgets sélectionnés ?;Q;YN;N;'
            else Msg:='51;'+Titre+';Désirez-vous fermer les budgets sélectionnés ?;Q;YN;N;'; END ;

if HShowMessage(Msg,'','')<>mrYes then exit ;
if L.AllSelected then
BEGIN
    Q.First;
    while Not Q.EOF do
    BEGIN
      MoveCur(False);
        Stc:=Q.FindField(StCode).AsString;
        if (Lefb=fbSect) then
        StA := Q.FindField('S_AXE').AsString ;

      // Vérification avant fermeture : FQ 16037 SBO 22/09/2005
      if (not Fermer) and EstDansVentilType then
        begin
        PgiBox( FabriqMess ) ;
        Break ;
        end ;

      // Vérification avant fermeture : FQ 22070  19.12.2007  YMO
      if (not Fermer) and EstDansGuideAbo then
        begin
        PgiBox( 'Le compte est utilisé dans un abonnement ou un guide' ) ;
        Break ;
        end ;

        if Transactions(OuvrirOuFermer,5)<>oeOK then
        BEGIN
          MessageAlerte('Ouverture/Fermeture impossible' ) ;
          Break ;
        END;
      Q.Next;
    END;
    L.AllSelected:=False;
END
ELSE
BEGIN
    for i:=0 to L.NbSelected-1 do
      BEGIN
        L.GotoLeBookMark(i) ;
      {$IFDEF EAGLCLIENT}
        Q.Seek(L.Row - 1);
      {$ENDIF}
        Stc:=Q.FindField(StCode).AsString;
        if (Lefb=fbSect) then
        StA := Q.FindField('S_AXE').AsString ;

      // Vérification avant fermeture : FQ 16037 SBO 22/09/2005
      if (not Fermer) and EstDansVentilType then
        begin
        PgiBox( FabriqMess ) ;
        Break ;
        end ;

      // Vérification avant fermeture : FQ 22070  19.12.2007  YMO
      if (not Fermer) and EstDansGuideAbo then
        begin
        PgiBox( 'Le compte est utilisé dans un abonnement ou un guide' ) ;
        Break ;
        end ;  

        if Transactions(OuvrirOuFermer,5)<>oeOK then
        BEGIN
          MessageAlerte('Ouverture/Fermeture impossible' ) ;
          Break ;
        END
      END ;
END;

  L.ClearSelected ;

  // MAJ des tablettes - FQ 16036 SBO 21/09/2005
  if Lefb = fbJal then
    AvertirMultiTable('TTJOURNAL');
END;

function  TFOuv.EstDansVentilType : Boolean ;
begin
  result := False ;
  if Lefb <> fbSect then Exit ;
  result := ExisteSQL('SELECT V_COMPTE FROM VENTIL WHERE V_SECTION="' + Stc + '" AND V_NATURE LIKE "%' + Copy(Sta, 2, 1) + '"') ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 19/12/2007
Modifié le ... :   /  /    
Description .. : Test d'existence du compte general dans un abonnement 
Suite ........ : ou un guide
Mots clefs ... : FQ 22070
*****************************************************************}
function  TFOuv.EstDansGuideAbo : Boolean ;
begin
  result := False ;
  result := ExisteSQL('SELECT 1 FROM ECRGUI WHERE EG_GENERAL="' + Stc + '"') ;
end ;

function  TFOuv.FabriqMess : String ;
var lQVentil  : TQuery ;
    lStNat    : String ;
    lStCompte : String ;
    lStLibVT  : String ;
begin
  result := '' ;
  if Lefb <> fbSect then Exit ;
  lQVentil := openSQL('SELECT V_NATURE, V_COMPTE, V_SECTION FROM VENTIL WHERE V_SECTION="' + Stc + '" AND V_NATURE LIKE "%' + Copy(Sta, 2, 1) + '"', True) ;
  if not lQVentil.Eof then
    begin
    result := 'La section ' + Stc + ' ne peut pas être fermée.' ;
    lStNat    := lQVentil.FindField('V_NATURE').AsString ;
    lStCompte := lQVentil.FindField('V_COMPTE').AsString ;
    lStLibVT  := RechDom( 'YYNATUREVENTIL', copy(lStNat, 1, 2), False ) ;
    if pos( 'GE', lStNat ) > 0 then // Ventilation des généraux
      result := result + ' Elle est utilisée dans la ventilation par défaut du compte ' + lStCompte + '.'
    else if pos( 'TY', lStNat ) > 0 then // Ventilation type comptable
      result := result + ' Elle est utilisée dans la ventilation type ' + RechDom( 'TTVENTILTYPE', lStCompte, False ) + '.'
    else if lStLibVT <> '' then // Autres ventilations types
      result := result + ' Elle est utilisée dans une ventilation par défaut de type ' + lStLibVT + '.'
    else result := result + ' Elle est utilisée par la table VENTIL.' ;

    end ;
  Ferme( lQVentil ) ;
end ;

procedure AGLOuvrFermListCpte(parms: array of variant; nb: integer );
var
  F : TFMul;
//  F : TForm;
{$IFDEF EAGLCLIENT}
  Liste : THGrid;
{$ELSE}
  Liste : THDBGrid;
{$ENDIF}
  Query   : TQuery;
  Evtouv  : TFouv;
  lf,fer  : string;
  Lefb  : TFichierBase ;
begin
  Lefb := fbGene;
//  F := TForm(Longint(Parms[0])) ;
  F := TFMul(Longint(Parms[0]));
  if F=Nil then exit ;

{$IFDEF EAGLCLIENT}
  Liste := THGrid(F.FindComponent('FListe'));
  if Liste=Nil then Exit;
  Query := F.Q.TQ;
{$ELSE}
  Liste := THDBGrid(F.FindComponent('FListe'));
  if Liste=Nil then Exit;
  Query := F.Q;
{$ENDIF}
  if (Query=Nil) then exit;

  Evtouv := TFouv.Create;
  lf :=  parms[1];
  if lf = 'fbGene' then Lefb := fbGene;
  if lf = 'fbAux'  then Lefb := fbAux;
  if lf = 'fbSect' then Lefb := fbSect;
  if lf = 'fbJal'  then Lefb := fbJal;
  fer := parms[2];
  if fer = 'FALSE' then Evtouv.Fermer := FALSE
                   else Evtouv.Fermer := TRUE;

  Case Lefb of
  fbGene  :BEGIN
      Evtouv.Pref:='G_' ; Evtouv.StCode:='G_GENERAL';
      Evtouv.LaTable:='GENERAUX';
      EvtOuv.Lefb := fbGene;
      EvtOuv.ModifieComptes(Liste,Query, F.Caption);
      END;
  fbAux  :BEGIN
      Evtouv.Pref:='T_' ; Evtouv.StCode:='T_AUXILIAIRE';
      Evtouv.LaTable:='TIERS';
      EvtOuv.Lefb := fbAux;
      EvtOuv.ModifieComptes(Liste,Query, F.Caption);
      END;
  fbSect  :BEGIN
      Evtouv.Pref:='S_' ; Evtouv.StCode:='S_SECTION';
      Evtouv.LaTable:='SECTION';
      EvtOuv.Lefb := fbSect;
      EvtOuv.ModifieComptes(Liste,Query, F.Caption);
      END;
  fbJal  :BEGIN
      Evtouv.Pref:='J_' ; Evtouv.StCode:='J_JOURNAL';
      Evtouv.LaTable:='JOURNAL';
      EvtOuv.Lefb := fbJal;
      EvtOuv.ModifieComptes(Liste,Query, F.Caption);
      END;

  else Exit ;
 End ;
EvtOuv.Free;
end;

Initialization
  RegisterAglProc( 'OuvrFermListCpte', TRUE , 1, AGLOuvrFermListCpte);
finalization


end.
