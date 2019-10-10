<template>
  <v-app class="indigo darken-3">
    <Toolbar />
    <v-content class="mx-8 mb-8 mt-10">
      <router-view />
    </v-content>
  </v-app>
</template>


<script>
import Toolbar from "@/components/Toolbar";

export default {
  name: "App",
  components: {
    toolbar: Toolbar
  },
  data: () => ({
    //
  }),
  computed: {
    hide() {
      return this.$route.path === "/login" || this.$route.path === "/register";
    }
  }
};
</script>
