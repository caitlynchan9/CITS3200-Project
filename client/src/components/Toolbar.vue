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
      <v-list-tile v-for="link in links" :key="link.text" router :to="link.route">
          <v-list-tile-action>
              <v-icon> {{ link.icon }} </v-icon>
          </v-list-tile-action>

          <v-list-tile-content>
              <v-list-tile-title> {{ link.text }} </v-list-tile-title>
          </v-list-tile-content>

      </v-list-tile>
    </v-navigation-drawer>
  </nav>
</template>

<script>
export default {
  data() {
    return {
      drawer: false,
      links: [
          {icon: 'fas fa-lock', text: 'Home', route: '/'},
          {icon: 'fas fa-edit', text: 'Competencies', route: '/competencies'},
          {icon: 'fas fa-list', text: 'Publish', route: '/publish'},
      ]
    }
  }
}
</script>