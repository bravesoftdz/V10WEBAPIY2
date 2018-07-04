unit UControlParam;

interface

uses Sysutils, UTOB, HCtrls, UDMIMP,
{$IFNDEF EAGLCLIENT}
   {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet, Variants, {$ENDIF}
{$ENDIF}
{$IFDEF CISXPGI}
  uYFILESTD, HEnt1, Uscript,
{$ENDIF}
{$IFDEF EAGLCLIENT}
  UScriptTob,
{$ELSE}
  HQRY,
{$ENDIF}
  stdctrls, HMsgBox;


type
  TSVControle = class(TObject)
  private
      Trouve            : Boolean;

  public
      TOBParam          : TOB;
      Valdefaut         : string;
      LePays,LeMode     : string;
      procedure ChargeTobParam (Domaine : string; ContenuObl : Boolean=TRUE; Listchoix : TListBox=nil);
      function parcourschampobligatoire (Val, car, Contenu : string) : Boolean;
      Function Conroleparam(var Val : variant; Name : string) : Boolean;
      procedure TOBParamFree;
{$IFDEF CISXPGI}
      procedure ChargementPGI (Domaine : string; Listchoix : TListBox=nil);
{$ENDIF}
  end;

var
	TPControle: TSVControle;

implementation

{$IFDEF CISXPGI}
procedure TSVControle.ChargementPGI (Domaine : string; Listchoix : TListBox=nil);
var
Q  : TQuery;
TA : TOB;
ST, ST1  : string;
begin
         if TOBParam <> nil then TOBParamFree;
        TOBParam := TOB.create ('', nil, -1);
(*         Q := OpenSQL ('SELECT * FROM DETABLES, DECHAMPS WHERE DT_NOMTABLE not like "%REF%" and '+
         ' DT_NOMTABLE not like "%CA3%" and DT_NOMTABLE not like "%F%" and DT_NOMTABLE not like "ICC%" '+
         'and DT_NOMTABLE not like "%LIEN%" and DH_TYPECHAMP<> "BLOB" and  '+
         'DH_TYPECHAMP<> "EXTENDED" and  DH_TYPECHAMP<> "DATA" and DT_DOMAINE="'+Domaine+'" and DH_PREFIXE=DT_PREFIXE order by DH_NUMCHAMP', TRUE);
*)
         Q := OpenSQL ('SELECT * FROM DETABLES, DECHAMPS WHERE (DT_NOMTABLE like "%EEXBQ%" or '+
         ' DT_NOMTABLE like "%ETEBAC%" or  DT_NOMTABLE="BANQUECP")  '+
         ' and DH_TYPECHAMP<> "BLOB" and  '+
         'DH_TYPECHAMP<> "EXTENDED" and  DH_TYPECHAMP<> "DATA" and DT_DOMAINE="'+Domaine+'" and DH_PREFIXE=DT_PREFIXE order by DH_NUMCHAMP', TRUE);

         while not Q.EOF do
         begin
              if Listchoix <> nil then
              begin
                    if Listchoix.Items.IndexOf (Q.FindField ('DT_NOMTABLE').asstring) < 0 then
                    begin
                       Q.next; continue;
                    end;
              end;
              TA := TOB.Create ('',TOBParam,-1);
              TA.AddChampSupValeur('Domaine', 'X');
              TA.AddChampSupValeur('TableName', Q.FindField ('DT_NOMTABLE').asstring);
              ST :=  Q.FindField ('DH_NOMCHAMP').asstring;
              St1 := ReadTokenPipe (St, '_');
              TA.AddChampSupValeur('Nom', ST);
              TA.AddChampSupValeur('Contenu', '');
              if (Q.FindField ('DH_TYPECHAMP').asstring = 'INTEGER') or
              (Q.FindField ('DH_TYPECHAMP').asstring = 'LONGCHAR') then
              begin
                 TA.AddChampSupValeur('Longueur', 8);
                 TA.AddChampSupValeur('TypeAlpha', 'N');
              end
              else
              if Q.FindField ('DH_TYPECHAMP').asstring = 'COMBO' then
              begin
                 TA.AddChampSupValeur('Longueur', 3);
                 TA.AddChampSupValeur('TypeAlpha', 'A');
              end
              else
              if Q.FindField ('DH_TYPECHAMP').asstring = 'BOOLEAN' then
              begin
                 TA.AddChampSupValeur('Longueur', 3);
                 TA.AddChampSupValeur('TypeAlpha', 'B');
              end
              else
              if (Q.FindField ('DH_TYPECHAMP').asstring = 'DOUBLE') or
              (Q.FindField ('DH_TYPECHAMP').asstring = 'RATE') then
              begin
                 TA.AddChampSupValeur('Longueur', 20);
                 TA.AddChampSupValeur('TypeAlpha', 'F');
              end
              else
              if Q.FindField ('DH_TYPECHAMP').asstring = 'DATE' then
              begin
                 TA.AddChampSupValeur('Longueur', 10);
                 TA.AddChampSupValeur('TypeAlpha', 'D');
              end
              else
              if pos ('CHAR', Q.FindField ('DH_TYPECHAMP').asstring) <> 0 then
              begin
                   St := Q.FindField ('DH_TYPECHAMP').asstring;
                   St1 := ReadTokenPipe (St, '(');
                   St1 := ReadTokenPipe (St, ')');
                   TA.AddChampSupValeur('Longueur', StrToint(St1));
                   TA.AddChampSupValeur('TypeAlpha', 'A');
              end
              else
                   TA.AddChampSupValeur('TypeAlpha', '');
              TA.AddChampSupValeur('Obligatoire', Q.FindField ('DH_OBLIGATOIRE').asstring);
              TA.AddChampSupValeur('NomPGI', '');
              TA.AddChampSupValeur('Contenuoblig', '');
              TA.AddChampSupValeur('MajusMin', '');

              TA.AddChampSupValeur('Fige', '0');
              TA.AddChampSupValeur('calcul', '');
              TA.AddChampSupValeur('Nbdecimal', '2');
              TA.AddChampSupValeur('Ordretri', '');
              TA.AddChampSupValeur('Famcorresp', '');
              TA.AddChampSupValeur('Commentaire', '');
              Q.Next;
         end;
         Ferme (Q);
end;
{$ENDIF}

procedure TSVControle.ChargeTobParam (Domaine : string; ContenuObl : Boolean=TRUE; Listchoix : TListBox=nil);
    var
    TA                     : TOB;
{$IFDEF CISXPGI}
    TobTemp                : TOB;
    CodeRetour             : integer;
    i                      : integer;
{$ELSE}
    QP                     : TQuery;
    ConnectionString       : string;
{$ENDIF}
begin
{$IFDEF CISXPGI}
        if (Domaine = 'X') and (LeMode='IMPORT') then  ChargementPGI ('C', Listchoix)
        else
        CodeRetour := RendTobparametre (Domaine, TobTemp, 'COMPTA', 'PARAM', LePays);
        //CodeRetour := RendTobparametre (Domaine, TobTemp);
        if CodeRetour = -1 then
        begin
              if TOBParam <> nil then TOBParamFree;
              TOBParam := TOB.create ('', nil, -1);
              For i:= 0 to TobTemp.detail.count-1 do
              begin
                              if Listchoix <> nil then
                              begin
                                    if Listchoix.Items.IndexOf (TobTemp.detail[i].Getvalue('TableName')) < 0 then
                                       continue;
                              end;
                             if not ContenuObl then
                             begin
                                  if  (TobTemp.detail[i].Getvalue('Contenuoblig') = '')
                                  and (TobTemp.detail[i].Getvalue('MajusMin') = '')
                                  and (TobTemp.detail[i].Getvalue('TypeAlpha') = '') then
                                       continue;
                              end;

                              TA := TOB.Create ('',TOBParam,-1);
                              TA.AddChampSupValeur('Domaine', TobTemp.detail[i].Getvalue('Domaine'));
                              TA.AddChampSupValeur('TableName', TobTemp.detail[i].Getvalue('TableName'));
                              TA.AddChampSupValeur('Nom', TobTemp.detail[i].Getvalue('Nom'));
                              TA.AddChampSupValeur('Contenu', TobTemp.detail[i].Getvalue('Contenu'));
                              TA.AddChampSupValeur('Longueur', TobTemp.detail[i].Getvalue('Longueur'));
                              TA.AddChampSupValeur('Obligatoire', TobTemp.detail[i].Getvalue('Obligatoire'));
                              TA.AddChampSupValeur('NomPGI', TobTemp.detail[i].Getvalue('NomPGI'));
                              TA.AddChampSupValeur('Contenuoblig', TobTemp.detail[i].Getvalue('Contenuoblig'));
                              TA.AddChampSupValeur('MajusMin', TobTemp.detail[i].Getvalue('MajusMin'));
                              TA.AddChampSupValeur('TypeAlpha', TobTemp.detail[i].Getvalue('TypeAlpha'));
                              TA.AddChampSupValeur('Fige', TobTemp.detail[i].Getvalue('Fige'));
                              TA.AddChampSupValeur('calcul', TobTemp.detail[i].Getvalue('calcul'));
                              TA.AddChampSupValeur('Nbdecimal', TobTemp.detail[i].Getvalue('Nbdecimal'));
                              TA.AddChampSupValeur('Ordretri', TobTemp.detail[i].Getvalue('Ordretri'));
                              TA.AddChampSupValeur('Famcorresp', TobTemp.detail[i].Getvalue('Famcorresp'));
                              TA.AddChampSupValeur('Commentaire', TobTemp.detail[i].Getvalue('Commentaire'))
              end;
              TobTemp.free;
        end;
{$ELSE}
        QP := nil;
        try
              if TOBParam <> nil then TOBParamFree;
              TOBParam := TOB.create ('', nil, -1);
{$IFDEF DBXPRESS}
              ConnectionString := DMImport.DBGlobal.ConnectionString;
{$ELSE}
              ConnectionString := DMImport.DBGlobal.Name;
{$ENDIF}
              QP := OpenSQLADO ('Select * from '+DMImport.GzImpPar.TableName
              + ' Where Domaine="'+Domaine+'"'
              +' order by Numero', ConnectionString);

              While  not (QP.Eof) do
              begin
                   with QP do
                   begin
                       if not ContenuObl then
                       begin
                            if  (FindField('Contenuoblig').AsString = '')
                            and (FindField('MajusMin').AsString = '')
                            and (FindField('TypeAlpha').AsString = '') then
                            begin
                                 QP.next;
                                 continue;
                            end;
                        end;
                        if Listchoix <> nil then
                        begin
                              if Listchoix.Items.IndexOf (FindField('TableName').AsString) < 0 then
                              begin
                                 QP.next;
                                 continue;
                              end;

                        end;

                        TA := TOB.Create ('',TOBParam,-1);
                        TA.AddChampSupValeur('Domaine', FindField('Domaine').AsString);
                        TA.AddChampSupValeur('TableName', FindField('TableName').AsString);
                        TA.AddChampSupValeur('Nom', FindField('Nom').AsString);
                        TA.AddChampSupValeur('Contenu', FindField('Contenu').AsString);
                        TA.AddChampSupValeur('Longueur', FindField('Longueur').Asinteger);
                        TA.AddChampSupValeur('Obligatoire', FindField('Obligatoire').AsString);
                        TA.AddChampSupValeur('NomPGI', FindField('NomPGI').AsString);
                        TA.AddChampSupValeur('Contenuoblig', FindField('Contenuoblig').AsString);
                        TA.AddChampSupValeur('MajusMin', FindField('MajusMin').AsString);
                        TA.AddChampSupValeur('TypeAlpha', FindField('TypeAlpha').AsString);
                        TA.AddChampSupValeur('Fige', FindField('Fige').Asinteger);
                        TA.AddChampSupValeur('calcul', FindField('calcul').Asvariant);
                        TA.AddChampSupValeur('Nbdecimal', FindField('Nbdecimal').Asinteger);
                        TA.AddChampSupValeur('Ordretri', FindField('Ordretri').AsString);
                        TA.AddChampSupValeur('Famcorresp', FindField('Famcorresp').AsString);
                        TA.AddChampSupValeur('Commentaire', FindField('Commentaire').AsString);
                        QP.Next;
                   end;
              end;
        finally
               QP.Free;
        end;
{$ENDIF}
end;

function TSVControle.parcourschampobligatoire (Val, car, Contenu : string) : Boolean;
var
NN,NM          : string;
ind, ind3      : integer;
begin
     NM := Contenu;
     repeat
           NN := ReadTokenPipe(NM, car);
           if (car = '-') and (NN <> '') then
           begin
                ind := StrToint(Val);
                NN := ReadTokenPipe(NM, car);
                if NN <> '' then
                begin
                     ind3 := StrToint (NN);
                     if (ind >= ind)  and (ind <= ind3) then
                     begin Result := TRUE; exit; end;
                end;
           end;
           if NN = Val then begin Result := TRUE; exit; end
           else Trouve := FALSE;
     until ( NN='' );
     Result := FALSE;
end;

Function TSVControle.Conroleparam(var Val : variant; Name : string) : Boolean;
var
NN,Tbl     : string;
T          : TOB;
begin
     Result := TRUE;
     NN := Name;
     if pos('_', NN) <> 0 then
          Tbl := ReadTokenPipe (NN, '_');
     Valdefaut := '';
     if TobParam = nil then exit;
     T := TobParam.FindFirst(['TableName', 'Nom'], [Tbl, NN], FALSE);
     if T <> nil then
     begin
               if T.GetValue('MajusMin') = 'MAJ' then
                  Val := UpperCase(Val);
               if T.GetValue('MajusMin') = 'Min' then
                  Val := LowerCase(Val);

               if T.getvalue ('Contenuoblig') = '' then exit;
               if T.getvalue ('Contenuoblig') = Val then exit
               else
               begin
                   NN := T.getvalue ('Contenuoblig');
                   if parcourschampobligatoire (Val, ';', NN) then exit;
                   if pos('-', NN) <> 0 then
                      if parcourschampobligatoire (Val, '-', NN) then exit;
               end;
               if not trouve then
               begin
                    Valdefaut := T.getvalue ('Contenuoblig'); Result := FALSE; exit;
               end;
     end;
end;

procedure TSVControle.TOBParamFree;
begin
     if TOBParam <> nil then TOBParam.free;
     TOBParam := nil;
end;




end.
