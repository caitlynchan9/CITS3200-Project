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
      <v-list rounded>
        <v-list-item-group v-model="item" color="indigo darken-4">
          <v-list-item v-for="(item, i) in items" :key="i" router :to="items.route">
            <v-list-item-icon>
              <v-icon v-text="item.icon"></v-icon>
            </v-list-item-icon>
            <v-list-item-content>
              <v-list-item-title v-text="item.text"></v-list-item-title>
            </v-list-item-content>
          </v-list-item>
        </v-list-item-group>
      </v-list>
    </v-navigation-drawer>
  </nav>
</template>

<script>
  export default {
    data: () => ({
      item: 1,
      items: [
        { text: 'Home', icon: 'mdi-clock', route: '/' },
        { text: 'Competencies', icon: 'mdi-account', route: '/competencies' },
        { text: 'Search', icon: 'mdi-flag', route: '/publish' },
      ],
    }),
  }
</script>