<template>
  <v-row justify="center">
    <v-col cols="12" sm="10" md="8" lg="6">
      <v-card>
        <v-form v-model="valid" @keyup.native.enter="valid && login($event)">
          <v-layout align-center justify-center column>
            <v-card-text>
              <v-text-field
                ref="username"
                v-model="username"
                :rules="[() => !!username || 'Please enter your username']"
                :error-messages="errorMessages"
                label
                placeholder="Username"
                required
              ></v-text-field>
              <v-text-field
                ref="password"
                v-model="password"
                :rules="[() => !!password || 'Please enter your password']"
                label
                required
                placeholder="Password"
              ></v-text-field>
            </v-card-text>
            <v-divider class="mt-1"></v-divider>
            <v-card-actions>
              <v-btn
                dark
                class="grey darken-4--text"
                color="yellow darken-1"
                :disabled="!valid || loading"
                :loading="loading"
                @click.stop.prevent="login"
              >Login</v-btn>
              <v-btn color="yellow darken-1" text @click="login">Login</v-btn>
            </v-card-actions>
          </v-layout>
        </v-form>
      </v-card>
    </v-col>
  </v-row>
</template>

<script>
import { mapState, mapActions } from "vuex";
import uwaLogin from "../assets/uwalogin.svg";

export default {
  name: "Login",
  data: () => ({
    user: {
      username: "",
      password: ""
    },
    valid: false,
    show: false,
    alert: false,
    uwaLogin: true,
    error: "",
    rules: {
      required: v => !!v || "This field is required",
      username: v => /.+@.+/.test(v) || "Username must be valid"
    }
  }),
  computed: {
    ...mapState("auth", { loading: "isAuthenticatePending" }),
    hide() {
      return this.$route.path === "/login" || this.$route.path === "/register";
    }
  },
  methods: {
    ...mapActions("auth", ["authenticate"]),
    async login() {
      if (this.valid) {
        await this.authenticate({
          strategy: "local",
          ...this.user
        })
          .then(async () => {
            // logged in
            this.$router.push({ name: "competencies" });
          })
          .catch(async e => {
            // Error on page
            this.alert = true;
            this.error = e.message;
          });
      }
    }
  }
};
</script>
