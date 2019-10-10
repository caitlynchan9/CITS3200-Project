<template>
  <v-card>
    <v-app-bar app clipped-left class="indigo darken-4">
      <v-app-bar-nav-icon @click="drawer = !drawer" class="white--text"></v-app-bar-nav-icon>
      <v-toolbar-title class="white--text">
        <!-- <img class="mt-12" :src="require('@/assets/uwa-crest.png')" height="80" /> -->
        <v-btn text dark>Talent Acquisition</v-btn>
      </v-toolbar-title>

      <div class="flex-grow-1"></div>

      <v-toolbar-items>
        <v-btn text class="white--text">
          <v-icon>mdi-account-badge-horizontal-outline</v-icon>Competencies
        </v-btn>
        <v-btn text class="white--text">
          <v-icon>mdi-magnify</v-icon>Search
        </v-btn>
        <v-btn text class="white--text">
          <v-icon>mdi-account</v-icon>Logout
        </v-btn>
      </v-toolbar-items>
    </v-app-bar>

    <v-navigation-drawer v-model="drawer" app class="indigo darken-1" clipped dark>
      <v-list>
        <v-list-item v-for="comp in competencies" :key="comp.title" @click.stop :to="comp.route">
          <v-list-item-icon>
            <v-icon>{{ comp.icon }}</v-icon>
          </v-list-item-icon>

          <v-list-item-content>
            <v-list-item-title>{{ comp.id }}</v-list-item-title>
          </v-list-item-content>
        </v-list-item>
      </v-list>

      <v-list-group sub-group no-action>
        <template v-slot:activator>
          <v-list-item-content>
            <v-list-item-title>Actions</v-list-item-title>
          </v-list-item-content>
        </template>
        <v-list-item v-for="(crud, i) in cruds" :key="i" @click.stop>
          <v-list-item-title v-text="crud[0]"></v-list-item-title>
          <v-list-item-action>
            <v-icon v-text="crud[1]"></v-icon>
          </v-list-item-action>
        </v-list-item>
      </v-list-group>

      <template v-slot:append>
        <div class="pa-4">
          <v-btn block>
            <v-icon>mdi-account</v-icon>Logout
          </v-btn>
        </div>
      </template>
    </v-navigation-drawer>
  </v-card>
</template>


<script>
export default {
  data() {
    return {
      drawer: true,
      competencies: [
        // { id: "Home", icon: "mdi-home", route: "/" },
        {
          id: "Competencies",
          icon: "mdi-account-badge-horizontal-outline",
          route: "/competencies"
        }
        // { id: "Search", icon: "mdi-magnify", route: "/search" }
      ],
      cruds: [
        ["Create new", "mdi-account-plus"],
        ["View previous", "mdi-account-search"],
        ["Delete old", "mdi-account-remove"]
      ]
    };
  }
};
</script>