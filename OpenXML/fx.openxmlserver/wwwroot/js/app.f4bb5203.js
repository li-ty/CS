(function(t){function e(e){for(var a,l,i=e[0],s=e[1],d=e[2],p=0,f=[];p<i.length;p++)l=i[p],Object.prototype.hasOwnProperty.call(o,l)&&o[l]&&f.push(o[l][0]),o[l]=0;for(a in s)Object.prototype.hasOwnProperty.call(s,a)&&(t[a]=s[a]);c&&c(e);while(f.length)f.shift()();return r.push.apply(r,d||[]),n()}function n(){for(var t,e=0;e<r.length;e++){for(var n=r[e],a=!0,i=1;i<n.length;i++){var s=n[i];0!==o[s]&&(a=!1)}a&&(r.splice(e--,1),t=l(l.s=n[0]))}return t}var a={},o={app:0},r=[];function l(e){if(a[e])return a[e].exports;var n=a[e]={i:e,l:!1,exports:{}};return t[e].call(n.exports,n,n.exports,l),n.l=!0,n.exports}l.m=t,l.c=a,l.d=function(t,e,n){l.o(t,e)||Object.defineProperty(t,e,{enumerable:!0,get:n})},l.r=function(t){"undefined"!==typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(t,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(t,"__esModule",{value:!0})},l.t=function(t,e){if(1&e&&(t=l(t)),8&e)return t;if(4&e&&"object"===typeof t&&t&&t.__esModule)return t;var n=Object.create(null);if(l.r(n),Object.defineProperty(n,"default",{enumerable:!0,value:t}),2&e&&"string"!=typeof t)for(var a in t)l.d(n,a,function(e){return t[e]}.bind(null,a));return n},l.n=function(t){var e=t&&t.__esModule?function(){return t["default"]}:function(){return t};return l.d(e,"a",e),e},l.o=function(t,e){return Object.prototype.hasOwnProperty.call(t,e)},l.p="";var i=window["webpackJsonp"]=window["webpackJsonp"]||[],s=i.push.bind(i);i.push=e,i=i.slice();for(var d=0;d<i.length;d++)e(i[d]);var c=s;r.push([0,"chunk-vendors"]),n()})({0:function(t,e,n){t.exports=n("56d7")},"034f":function(t,e,n){"use strict";n("85ec")},"56d7":function(t,e,n){"use strict";n.r(e);n("e623"),n("e379"),n("5dc8"),n("37e1");var a=n("2b0e"),o=function(){var t=this,e=t.$createElement,n=t._self._c||e;return n("div",{attrs:{id:"app"}},[n("Form",{staticClass:"form",attrs:{"v-model":t.form,"label-width":90,"label-colon":""}},[n("div",{staticClass:"split-pane left"},[n("FormItem",{attrs:{label:"JSON","label-width":0}},[n("Input",{attrs:{type:"textarea",rows:23,autosize:{minRows:23,maxRows:23}},model:{value:t.form.textarea,callback:function(e){t.$set(t.form,"textarea",e)},expression:"form.textarea"}})],1)],1),n("div",{staticClass:"split-pane right"},[n("FormItem",{attrs:{label:"模板文件"}},[n("Upload",{staticStyle:{display:"inline-block"},attrs:{"before-upload":t.handelUpload0,action:"//jsonplaceholder.typicode.com/posts/"}},[n("Button",[t._v("选择文件")])],1),t._v(" "+t._s(t.filesName0)+" ")],1),n("FormItem",{attrs:{label:"生成文件名"}},[n("Input",{staticStyle:{display:"inline-block",width:"200px"},model:{value:t.form.fileName,callback:function(e){t.$set(t.form,"fileName",e)},expression:"form.fileName"}}),t._v(".docx ")],1),n("FormItem",{attrs:{label:"图片1"}},[n("Upload",{staticStyle:{display:"inline-block"},attrs:{"before-upload":t.handelUpload1,action:"//jsonplaceholder.typicode.com/posts/",format:["jpg","jpeg","png"]}},[n("Button",[t._v("选择图片")])],1),t._v(" "+t._s(t.filesName1)+" ")],1),n("FormItem",{attrs:{label:"图片2"}},[n("Upload",{staticStyle:{display:"inline-block"},attrs:{"before-upload":t.handelUpload2,action:"//jsonplaceholder.typicode.com/posts/",format:["jpg","jpeg","png"]}},[n("Button",[t._v("选择图片")])],1),t._v(" "+t._s(t.filesName2)+" ")],1),n("FormItem",[n("Button",{attrs:{type:"info"},on:{click:t.upLoad}},[t._v(" 提交 ")])],1),n("FormItem",[n("a",{attrs:{href:"http://192.168.0.121:10045/%E6%A8%A1%E6%9D%BF.docx"}},[n("Button",{attrs:{type:"info"}},[t._v(" 下载模板文件 ")])],1)])],1)])],1)},r=[],l=(n("b0c0"),n("d3b7"),n("3ca3"),n("ddb0"),n("2b3d"),n("96cf"),n("1da1")),i=n("bc3a"),s=n.n(i),d={name:"App",data:function(){return{form:{textarea:'{\n\t"name": "插入的文本",\n\t"outer": {\n\t\t"inner": {\n\t\t\t"str": "OK"\n\t\t}\n\t},\n\t"image": "img1",\n\t"checkBox": false,\n\t"para": "May there be enough clouds to make a beautiful sunset",\n\t"tb": [{\n\t\t"name": "Lee",\n\t\t"age": 20,\n\t\t"sex": "MALE"\n\t}, {\n\t\t"name": "Li",\n\t\t"age": 20,\n\t\t"sex": "FEMALE"\n\t}],\n\t"table": {\n\t\t"columns": [{\n\t\t\t"width": 1000\n\t\t}, {\n\t\t\t"width": 2000\n\t\t}, {\n\t\t\t"width": 1000\n\t\t}, {\n\t\t\t"width": 2000\n\t\t}, {\n\t\t\t"width": 3000\n\t\t}],\n\t\t"rows": [{\n\t\t\t"height": 1000\n\t\t}, {\n\t\t\t"height": 1000\n\t\t}, {}, {}, {}],\n\t\t"cells": [\n\t\t\t[{\n\t\t\t\t"DataType": "1",\n\t\t\t\t"data": "img2",\n\t\t\t\t"colspan": 2,\n\t\t\t\t"rowspan": 2\n\t\t\t}, {\n\t\t\t\t"data": "data"\n\t\t\t}, {\n\t\t\t\t"data": "data"\n\t\t\t}, {\n\t\t\t\t"data": "data",\n\t\t\t\t"rowspan": 5\n\t\t\t}],\n\t\t\t[{\n\t\t\t\t"data": "data",\n\t\t\t\t"colspan": 2\n\t\t\t}, {\n\t\t\t\t"data": "data"\n\t\t\t}, {\n\t\t\t\t"data": "data",\n\t\t\t\t"rowspan": 4\n\t\t\t}, {\n\t\t\t\t"data": "data"\n\t\t\t}],\n\t\t\t[{\n\t\t\t\t"data": "data",\n\t\t\t\t"rowspan": 3\n\t\t\t}, {\n\t\t\t\t"data": "data"\n\t\t\t}, {\n\t\t\t\t"data": "data"\n\t\t\t}, {\n\t\t\t\t"data": "data"\n\t\t\t}, {\n\t\t\t\t"data": "data"\n\t\t\t}],\n\t\t\t[{\n\t\t\t\t"data": "data"\n\t\t\t}, {\n\t\t\t\t"data": "data"\n\t\t\t}, {\n\t\t\t\t"data": "data"\n\t\t\t}, {\n\t\t\t\t"data": "data"\n\t\t\t}, {\n\t\t\t\t"data": "data"\n\t\t\t}],\n\t\t\t[{\n\t\t\t\t"data": "data"\n\t\t\t}, {\n\t\t\t\t"data": "data",\n\t\t\t\t"colspan": 2\n\t\t\t}, {\n\t\t\t\t"data": "data"\n\t\t\t}, {\n\t\t\t\t"data": "data"\n\t\t\t}]\n\t\t]\n\t}\n}',fileName:""},files:[],filesName0:"",filesName1:"",filesName2:""}},methods:{handelUpload0:function(t){return this.files[0]=t,this.filesName0=t.name,console.log(this.filesName0),console.log(t),!1},handelUpload1:function(t){return this.files[1]=t,this.filesName1=t.name,console.log(t),!1},handelUpload2:function(t){return this.files[2]=t,this.filesName2=t.name,console.log(t),!1},upLoad:function(){var t=this;if(console.log(this.files),0!=this.files.length){var e=new FormData;e.append("data",t.form.textarea),e.append("template",t.files[0]),t.files[1]&&e.append("img1",t.files[1]),t.files[2]&&e.append("img2",t.files[2]),s()({url:"/api/officemake/DocxGenarate",method:"post",data:e,processData:!1,contentType:!1,responseType:"blob"}).then(function(){var e=Object(l["a"])(regeneratorRuntime.mark((function e(n){var a,o,r,l;return regeneratorRuntime.wrap((function(e){while(1)switch(e.prev=e.next){case 0:if(a=n.data,"text/plain"!=a.type){e.next=7;break}return e.next=4,new Response(a).text();case 4:return o=e.sent,t.$Message.error(o),e.abrupt("return");case 7:r=window.URL.createObjectURL(a),l=document.createElement("a"),document.body.appendChild(l),l.href=r,t.form.fileName&&(l.download=t.form.fileName+".docx"),l.click(),window.URL.revokeObjectURL(r);case 14:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}()).catch((function(){t.$Message.error("请求失败！")}))}else this.$Message.error("请上传模板文件")}}},c=d,p=(n("034f"),n("2877")),f=Object(p["a"])(c,o,r,!1,null,null,null),u=f.exports,m=n("f825"),h=n.n(m);n("f8ce");a["default"].config.productionTip=!1,a["default"].use(h.a),new a["default"]({render:function(t){return t(u)}}).$mount("#app")},"85ec":function(t,e,n){}});
//# sourceMappingURL=app.f4bb5203.js.map