<template>
  <nav>
    <v-toolbar flat app color="indigo darken-4" height="150">
        <v-app-bar-nav-icon @click="drawer = !drawer"></v-app-bar-nav-icon>
      <v-toolbar-title>
        <a>
          <img src="@/assets/uwa-logo.png" alt />
        </a>
      </v-toolbar-title>
      <v-toolbar-title right>Logout</v-toolbar-title>
    </v-toolbar>

    <v-navigation-drawer v-model="drawer" color="grey lighten-1">
      <v-btn block>Home</v-btn>
      <v-btn block>Competencies</v-btn>
      <v-btn block>Search</v-btn>
      <v-btn block>Logout</v-btn>
    </v-navigation-drawer>
  </nav>
</template>

<script>
export default {
  data() {
    return {
      drawer: false
    }
  }
}
</script>