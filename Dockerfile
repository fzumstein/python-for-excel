FROM ghcr.io/astral-sh/uv:python3.13-trixie

# COPY --from=ghcr.io/astral-sh/uv:latest /uv /uvx /bin/

COPY .python-version .
COPY pyproject.toml .
COPY uv.lock .

RUN uv sync --locked

# create user with a home directory
ARG NB_USER=jovyan
ARG NB_UID=1000
ENV USER ${NB_USER}
ENV NB_UID ${NB_UID}
ENV HOME /home/${NB_USER}

RUN adduser --disabled-password \
    --gecos "Default user" \
    --uid ${NB_UID} \
    ${NB_USER}
