import Vue from "vue";
import Router from "vue-router";
import Home from "./views/Home.vue";
import Competencies from "./views/Competencies.vue";
import Publish from "./views/Publish.vue";

Vue.use(Router);

export default new Router({
  mode: "history",
  base: process.env.BASE_URL,
  routes: [
    {
      path: "/",
      name: "home",
      component: Home
    },
    {
      path: "/competencies",
      name: "competencies",
      component: Competencies
    },
    {
      path: "/publish",
      name: "publish",
      component: Publish
    }
  ]
});
