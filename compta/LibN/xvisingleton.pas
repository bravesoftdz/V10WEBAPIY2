unit XVISingleton;

interface

uses Classes ; 

Type
     TXVIBaseSingleton = Class(TObject)
      protected
       Class Function IsSingleton : Boolean ; virtual ; Abstract ;
       Class Function DoNewInstance : TObject ; virtual ;
       Procedure DoFreeInstance ; virtual ;
      public
       class function NewInstance: TObject; override ;
       Class Function RefCount: Integer ;
       procedure FreeInstance; override ;
      end ;

     TXVISingleList = Class (TList)
      protected
       Class Function IsSingleton : Boolean ; virtual ; Abstract ;
       Class Function DoNewInstance : TObject ; virtual ;
       Procedure DoFreeInstance ; virtual ;
      public
       class function NewInstance: TObject; override ;
       Class Function RefCount: Integer ;
       procedure FreeInstance; override ;
      end ;

     TXVISingleStringList = Class (TStringList)
      protected
       Class Function IsSingleton : Boolean ; virtual ; Abstract ;
       Class Function DoNewInstance : TObject ; virtual ;
       Procedure DoFreeInstance ; virtual ;
      public
       class function NewInstance: TObject; override ;
       Class Function RefCount: Integer ;
       procedure FreeInstance; override ;
      end ;
implementation

uses sysutils, Dialogs ;


///////////////////////////////////////////////////////////////////////////////////
//                   TXVISingletonManager                                        //
///////////////////////////////////////////////////////////////////////////////////

type
   pXVISingleton = ^TXVISingleton ;
   TXVISingleton = Record
    Instance : TObject ;
    Count    : Integer ;
    End ;

Type
   TXVISingletonManager = Class(TStringList)
    public
     destructor Destroy ; override ;
    end ;

Var SingleManager : TXVISingletonManager ;

//////////////////////////////////////////////////////////////////////////////////
Destructor TXVISingletonManager.Destroy ;
var
   Single : pXVISingleton ;
Begin
   if Count>0 then Begin
     //ShowMessage('Quedan '+inttostr(Count)+' Singleton(s) por suprimir!!!') ;
     While Count>0 do Begin
        Single:=pXVISingleton(Objects[Count-1]) ;
        Single.Instance.Free ;
     End
   End ;
   Inherited ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Function SingletonManager : TXVISingletonManager ;
Begin
   if Not Assigned(SingleManager) then
      SingleManager:=TXVISingletonManager.Create ;
   Result:=SingleManager ;
End ;
//////////////////////////////////////////////////////////////////////////////////
(*class function TXVIBaseSingleton.IsSingleton : Boolean ;
Begin
Result:=TRUE ;
End ;*)
///////////////////////////////////////////////////////////////////////////////////
//                   TXVIBaseSingleton                                           //
///////////////////////////////////////////////////////////////////////////////////

Class Function TXVIBaseSingleton.RefCount: Integer ;
var
   Idx    : integer ;
Begin
   result:=0 ;
   Idx:=SingletonManager.Indexof(Classname) ;
   if (IsSingleton) and (idx>-1) then
      Result:=pXVISingleton(SingletonManager.Objects[idx]).Count ;
End ;
//////////////////////////////////////////////////////////////////////////////////
class function TXVIBaseSingleton.NewInstance: TObject;
var
   Idx    : integer ;
   Single : pXVISingleton ;
Begin
   Idx:=SingletonManager.Indexof(Classname) ;
   if (Not IsSingleton) or (idx<0) then Begin
      New(Single) ;
      Single^.Instance:=DoNewInstance ;
      Single^.Count:= 0 ;
      if IsSingleton then
         SingletonManager.AddObject(Classname,TObject(Single)) ;
   End else
      Single:=pXVISingleton(SingletonManager.Objects[idx]) ;
   inc(Single^.Count) ;
   Result:=Single^.Instance ;
   if Not IsSingleton then
      Dispose(Single) ;
End ;
//////////////////////////////////////////////////////////////////////////////////
procedure TXVIBaseSingleton.FreeInstance;
var
   Idx    : integer ;
   Single : pXVISingleton ;
Begin
   Idx:=SingletonManager.Indexof(Classname) ;
   single:=nil ;
   if idx>-1 then Begin
      Single:=pXVISingleton(SingletonManager.Objects[idx]) ;
      Dec(Single^.Count) ;
   end ;
   if (idx<0) or (Assigned(Single) and (single^.Count<=0)) then begin
      doFreeInstance ;
      if idx>-1 then Begin
         Dispose(Single) ;
         SingletonManager.delete(idx) ;
      End ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Class Function TXVIBaseSingleton.DoNewInstance : TObject ;
Begin
   Result:=inherited NewInstance ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Procedure TXVIBaseSingleton.DoFreeInstance ;
Begin
   inherited freeinstance ;
End ;
//////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////
//                   TXViSingleList                                           //
///////////////////////////////////////////////////////////////////////////////////

Class Function TXViSingleList.RefCount: Integer ;
var
   Idx    : integer ;
Begin
   result:=0 ;
   Idx:=SingletonManager.Indexof(Classname) ;
   if (IsSingleton) and (idx>-1) then
      Result:=pXVISingleton(SingletonManager.Objects[idx]).Count ;
End ;
//////////////////////////////////////////////////////////////////////////////////
class function TXViSingleList.NewInstance: TObject;
var
   Idx    : integer ;
   Single : pXVISingleton ;
Begin
   Idx:=SingletonManager.Indexof(Classname) ;
   if (Not IsSingleton) or (idx<0) then Begin
      New(Single) ;
      Single^.Instance:=DoNewInstance ;
      Single^.Count:= 0 ;
      if IsSingleton then
         SingletonManager.AddObject(Classname,TObject(Single)) ;
   End else
      Single:=pXVISingleton(SingletonManager.Objects[idx]) ;
   inc(Single^.Count) ;
   Result:=Single^.Instance ;
   if Not IsSingleton then
      Dispose(Single) ;
End ;
//////////////////////////////////////////////////////////////////////////////////
procedure TXViSingleList.FreeInstance;
var
   Idx    : integer ;
   Single : pXVISingleton ;
Begin
   Idx:=SingletonManager.Indexof(Classname) ;
   single:=nil ;
   if idx>-1 then Begin
      Single:=pXVISingleton(SingletonManager.Objects[idx]) ;
      Dec(Single^.Count) ;
   end ;
   if (idx<0) or (Assigned(Single) and (single^.Count<=0)) then begin
      doFreeInstance ;
      if idx>-1 then Begin
         Dispose(Single) ;
         SingletonManager.delete(idx) ;
      End ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Class Function TXVISingleList.DoNewInstance : TObject ;
Begin
   Result:=inherited NewInstance ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Procedure TXVISingleList.DoFreeInstance ;
Begin
   inherited freeinstance ;
End ;
//////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////
//                   TXViSingleStringList                                           //
///////////////////////////////////////////////////////////////////////////////////

Class Function TXViSingleStringList.RefCount: Integer ;
var
   Idx    : integer ;
Begin
   result:=0 ;
   Idx:=SingletonManager.Indexof(Classname) ;
   if (IsSingleton) and (idx>-1) then
      Result:=pXVISingleton(SingletonManager.Objects[idx]).Count ;
End ;
//////////////////////////////////////////////////////////////////////////////////
class function TXViSingleStringList.NewInstance: TObject;
var
   Idx    : integer ;
   Single : pXVISingleton ;
Begin
   Idx:=SingletonManager.Indexof(Classname) ;
   if (Not IsSingleton) or (idx<0) then Begin
      New(Single) ;
      Single^.Instance:=DoNewInstance ;
      Single^.Count:= 0 ;
      if IsSingleton then
         SingletonManager.AddObject(Classname,TObject(Single)) ;
   End else
      Single:=pXVISingleton(SingletonManager.Objects[idx]) ;
   inc(Single^.Count) ;
   Result:=Single^.Instance ;
   if Not IsSingleton then
      Dispose(Single) ;
End ;
//////////////////////////////////////////////////////////////////////////////////
procedure TXViSingleStringList.FreeInstance;
var
   Idx    : integer ;
   Single : pXVISingleton ;
Begin
   Idx:=SingletonManager.Indexof(Classname) ;
   single:=nil ;
   if idx>-1 then Begin
      Single:=pXVISingleton(SingletonManager.Objects[idx]) ;
      Dec(Single^.Count) ;
   end ;
   if (idx<0) or (Assigned(Single) and (single^.Count<=0)) then begin
      doFreeInstance ;
      if idx>-1 then Begin
         Dispose(Single) ;
         SingletonManager.delete(idx) ;
      End ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Class Function TXVISingleStringList.DoNewInstance : TObject ;
Begin
   Result:=inherited NewInstance ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Procedure TXVISingleStringList.DoFreeInstance ;
Begin
   inherited freeinstance ;
End ;
//////////////////////////////////////////////////////////////////////////////////
initialization
   SingleManager:=nil ;

Finalization
   if Assigned(SingleManager) then
      SingleManager.Free ;

end.
