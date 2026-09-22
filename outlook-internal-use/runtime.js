var TARGET_LABEL_NAME="Internal Use";
Office.onReady(function(){});
function done(e){if(e&&e.completed)e.completed();}
function flat(a,o){o=o||[];if(!a)return o;for(var i=0;i<a.length;i++){o.push(a[i]);if(a[i].children)flat(a[i].children,o);}return o;}
function findLabel(a,n){a=flat(a,[]);for(var i=0;i<a.length;i++)if(a[i].name&&a[i].name.toLowerCase()===n.toLowerCase())return a[i];return null;}
function onMessageCompose(event){
 try{
  if(!Office.context.requirements.isSetSupported("Mailbox","1.13")){console.log("Mailbox 1.13 unsupported");done(event);return;}
  Office.context.sensitivityLabelsCatalog.getIsEnabledAsync({asyncContext:event},function(er){
   var ev=er.asyncContext;if(er.status!==Office.AsyncResultStatus.Succeeded||!er.value){console.log("Label catalog unavailable");done(ev);return;}
   Office.context.sensitivityLabelsCatalog.getAsync({asyncContext:ev},function(cr){
    var ev2=cr.asyncContext;if(cr.status!==Office.AsyncResultStatus.Succeeded){console.log("Cannot read catalog");done(ev2);return;}
    var target=findLabel(cr.value,TARGET_LABEL_NAME);if(!target){console.log("Internal Use label not found");done(ev2);return;}
    Office.context.mailbox.item.sensitivityLabel.getAsync({asyncContext:{event:ev2,target:target}},function(rr){
     var c=rr.asyncContext;if(rr.status!==Office.AsyncResultStatus.Succeeded){console.log("Cannot read current label");done(c.event);return;}
     if(rr.value){console.log("Existing label present; unchanged: "+rr.value);done(c.event);return;}
     Office.context.mailbox.item.sensitivityLabel.setAsync(c.target,{asyncContext:c.event},function(sr){
      console.log(sr.status===Office.AsyncResultStatus.Succeeded?"Applied Internal Use":"Unable to set label");
      done(sr.asyncContext);
     });
    });
   });
  });
 }catch(x){console.log("Label helper error: "+x.message);done(event);}
}
Office.actions.associate("onMessageCompose",onMessageCompose);