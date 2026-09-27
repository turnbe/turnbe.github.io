var UN="Unlabelled Sensitivity";

function flat(a,o){
  o=o||[];
  if(!a)return o;
  for(var i=0;i<a.length;i++){
    o.push(a[i]);
    if(a[i].children)flat(a[i].children,o);
  }
  return o;
}
function nameById(c,id){
  var a=flat(c,[]);
  for(var i=0;i<a.length;i++)if(a[i].id===id)return a[i].name;
  return null;
}
function openPickerIfUnlabelled(event){
  try{
    Office.context.sensitivityLabelsCatalog.getIsEnabledAsync({asyncContext:event},function(en){
      var ev=en.asyncContext;
      if(en.status!==Office.AsyncResultStatus.Succeeded||!en.value){ev.completed();return;}
      Office.context.sensitivityLabelsCatalog.getAsync({asyncContext:ev},function(cat){
        var ev2=cat.asyncContext;
        if(cat.status!==Office.AsyncResultStatus.Succeeded){ev2.completed();return;}
        Office.context.mailbox.item.sensitivityLabel.getAsync({asyncContext:{event:ev2,catalog:cat.value}},function(cur){
          var c=cur.asyncContext;
          if(cur.status!==Office.AsyncResultStatus.Succeeded){c.event.completed();return;}
          var n=nameById(c.catalog,cur.value)||"";
          if(n.toLowerCase()===UN.toLowerCase() &&
             Office.addin && Office.addin.showAsTaskpane){
            Office.addin.showAsTaskpane().then(function(){c.event.completed();})
              .catch(function(){c.event.completed();});
          }else{
            c.event.completed();
          }
        });
      });
    });
  }catch(x){event.completed();}
}

function onMessageCompose(e){openPickerIfUnlabelled(e);}

function onMessageSend(e){
  try{
    Office.context.sensitivityLabelsCatalog.getIsEnabledAsync({asyncContext:e},function(en){
      var ev=en.asyncContext;
      if(en.status!==Office.AsyncResultStatus.Succeeded||!en.value){ev.completed({allowEvent:true});return;}
      Office.context.sensitivityLabelsCatalog.getAsync({asyncContext:ev},function(cat){
        var ev2=cat.asyncContext;
        if(cat.status!==Office.AsyncResultStatus.Succeeded){ev2.completed({allowEvent:true});return;}
        Office.context.mailbox.item.sensitivityLabel.getAsync({asyncContext:{event:ev2,catalog:cat.value}},function(cur){
          var c=cur.asyncContext;
          if(cur.status!==Office.AsyncResultStatus.Succeeded){c.event.completed({allowEvent:true});return;}
          var n=nameById(c.catalog,cur.value)||"";
          if(n.toLowerCase()===UN.toLowerCase()){
            c.event.completed({
              allowEvent:false,
              errorMessage:"This email is still labelled Unlabelled Sensitivity. Select the appropriate sensitivity label before sending.",
              cancelLabel:"Choose label",
              commandId:"Prompt.Open"
            });
          }else{
            c.event.completed({allowEvent:true});
          }
        });
      });
    });
  }catch(x){e.completed({allowEvent:true});}
}

Office.actions.associate("onMessageCompose",onMessageCompose);
Office.actions.associate("onMessageSend",onMessageSend);