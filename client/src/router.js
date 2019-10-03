import Vue from "vue";
import Router from "vue-router";
import Store from "./store";
import Login from "./views/Login.vue";

Vue.use(Router);

export default new Router({
  mode: "history",
  base: process.env.BASE_URL,
  routes: [
    {
      path: "/login",
      name: "login",
      component: Login
    }
    // {
    //   path: "/",
    //   name: "competencies",
    //   component: Competencies,
    //   beforeEnter(to, from, next) {
    //     const user = Store.getters["users/current"];
    //     if (!user) {
    //       next({ name: "login" });
    //     }
    //     next();
    //   }
    // }
  ]
});
