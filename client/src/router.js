import Vue from "vue";
import Router from "vue-router";
import Store from "./store";
import Login from "./views/Login.vue";
import Competencies from "./views/Competencies.vue";
import Publish from "./views/Publish.vue";

Vue.use(Router);

export default new Router({
  mode: "history",
  base: process.env.BASE_URL,
  routes: [
    {
      path: "/login",
      name: "login",
      component: Login
    },
    {
      path: "/publish",
      name: "publish",
      component: Publish
    },
    {
      path: "/",
      name: "competencies",
      component: Competencies,
      beforeEnter(to, from, next) {
        const user = Store.getters["users/current"];
        if (!user) {
          next({ name: "login" });
        }
        next();
      }
    }
  ]
});
